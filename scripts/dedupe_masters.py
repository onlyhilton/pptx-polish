# -*- coding: utf-8 -*-
"""母版/版式层收口（换模板后必做）

解决什么问题
------------
用 PowerPoint「设计 → 应用模板」换过一次模板后，包里会出现 **两套母版**：

    presentation.xml  sldMasterIdLst
      rId1 -> slideMaster1.xml   （旧的一套，版式 slideLayout1..16，可能缺装饰形状）
      rId2 -> slideMaster2.xml   （PowerPoint 新建的一套，形状完整）

34 页全部迁到新母版的版式上，**但旧母版仍列在 sldMasterIdLst 里**，
于是 PowerPoint 的「版式」面板里排在最前、用户随手能选到的正是那批**残缺的旧版式**。
用户下一次「重设版式」或换版式，模板的标志性装饰（标题左侧绿条、标题下分隔线、
封底 THANKS!）立刻打回原形。

本脚本做两件事：

  1) **孤儿母版清理**（默认分析，`--out-deck` 执行）
     一个母版若"它名下的版式没有任何一页引用"，即为孤儿。
     整批删掉：该母版 + 它的全部版式（含伴生 .rels）+ 只被它引用的 theme / tags /
     embeddings / media，并同步修 presentation.xml 的 sldMasterIdLst、
     presentation.xml.rels、[Content_Types].xml。
     **演示文稿级主题按"血统"重指**：presentation.xml.rels 指的 theme 若挂在被删母版下，
     改为指向存活母版的主题。光删母版不重指，主题库里会残留一个过期主题，
     新建形状仍继承旧字体/配色 —— 又一次"看起来换过模板但没真换"。
     （判定按血统而非引用计数：旧主题常被 presentation 与旧母版同时引用，
      引用计数=2，按计数扫会被判成"有人用"而漏掉。本轮实测踩过。）

  2) **重设版式（Reset Slide）之后的收口**（需显式开 flag）
     PowerPoint 的 Reset Slide 在 OOXML 层等价于「删掉 slide 上占位符自己的 <a:xfrm>」，
     并有两个副作用，都靠这里清掉：
       --unify-titles      标题占位符删 <a:xfrm>，位置统一回归版式（不再"上上下下"）
       --drop-stray-body   删掉"文本为空且无图片"的 body 占位符（Reset 补出来的空实例）；
                           版式自带的留空行可用 --keep-empty-body-on "<版式名>,..." 豁免
       --pagenum-x N       没有 sldNum 槽位的模板，Reset 会把页码 off.x 写成 0（贴左边缘），
                           改回正确值（EMU）。配套 --pagenum-y N

用法
----
    # ① 只分析
    python dedupe_masters.py "<deck.pptx>"

    # ② 清孤儿母版
    python dedupe_masters.py "<deck.pptx>" --out-deck "<new.pptx>"

    # ③ 清孤儿母版 + Reset 之后收口（T3T4 ISP 实战参数）
    python dedupe_masters.py "<deck.pptx>" --out-deck "<new.pptx>" \
        --unify-titles --drop-stray-body --keep-empty-body-on "目录,Workshop index" \
        --pagenum-x 236220 --pagenum-y 6153150

安全约定
--------
- 只删"引用计数为 0"的部件；任何仍被别处引用的 theme/tags/media 都不动。
  唯一例外：`presentation.xml.rels` 的主题若挂在被删母版下，先重指再回收
  （此时会把该 .rels 排除在引用扫描之外，因为它即将被改写）。
- `[Content_Types].xml` 在产物里**保持为第一个 ZIP 条目**（否则 PowerPoint 拒收）。
- 产出后请跑 `verify_pptx.py` + `open_test.py`；期望 `悬空=0`、母版只剩一个。
"""
from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ns = {"a": A, "p": P}

CT_PART = "[Content_Types].xml"
PRES = "ppt/presentation.xml"
PRES_RELS = "ppt/_rels/presentation.xml.rels"


# ------------------------------------------------------------------ 工具
def zread(z, name):
    return z.read(name)


def resolve(host, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(host), target))


def part_rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def host_of_rels(rp):
    """由 `X/_rels/Y.rels` 反推宿主部件 `X/Y`。"""
    d, b = posixpath.split(rp)                  # d = X/_rels, b = Y.rels
    return posixpath.join(posixpath.dirname(d), b[: -len(".rels")])


def build_referenced(z, names, skip=()):
    """全包"被谁引用"集合。

    **必须扫遍所有 .rels**（含 slideMasters / slideLayouts 自己的 rels），
    且 target 要相对**宿主部件**所在目录解析（不是相对 .rels 文件）。
    早期版本漏了这两点：跳过 slideLayouts 目录 + 用 .rels 当宿主，
    于是"只被存活版式引用的媒体"被判成无引用而删掉 → 52 条悬空引用（verify_pptx 抓到）。
    `skip` 里的宿主部件（即将被删的）不算引用来源。
    """
    ref = set()
    for n in names:
        if not n.endswith(".rels"):
            continue
        host = host_of_rels(n)
        if host in skip:
            continue
        base = posixpath.dirname(host)
        for r in etree.fromstring(z.read(n)):
            t = r.get("Target", "")
            ref.add(t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t)))
    return ref


def read_rels(z, part):
    """返回 {rId: (type_local, abs_target)}；无 rels 返回 {}"""
    rp = part_rels_path(part)
    if rp not in z.namelist():
        return {}
    out = {}
    for r in etree.fromstring(z.read(rp)):
        out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1],
                            resolve(part, r.get("Target", "")))
    return out


def to_bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ------------------------------------------------------------------ 分析
def analyze(deck):
    z = zipfile.ZipFile(deck)
    names = set(z.namelist())
    pres = etree.fromstring(z.read(PRES))
    prels = read_rels(z, PRES)

    # 母版清单（顺序 = PowerPoint 里的顺序）
    masters = []
    for m in pres.findall(".//p:sldMasterIdLst/p:sldMasterId", ns):
        rid = m.get("{%s}id" % R)
        tgt = prels.get(rid, (None, None))[1]
        if tgt:
            masters.append({"rId": rid, "id": m.get("id"), "part": tgt})

    # 每个母版的版式 + 每页实际引用
    slide_use = defaultdict(int)          # layout part -> 引用页数
    slide_to_layout = {}
    for n in sorted(names):
        if re.match(r"ppt/slides/slide\d+\.xml$", n):
            for _, (typ, tgt) in read_rels(z, n).items():
                if typ == "slideLayout":
                    slide_use[tgt] += 1
                    slide_to_layout[n] = tgt

    report_masters = []
    for mm in masters:
        mrels = read_rels(z, mm["part"])
        lay_ids = [s.get("{%s}id" % R) for s in
                   etree.fromstring(z.read(mm["part"])).findall(".//p:sldLayoutIdLst/p:sldLayoutId", ns)]
        lays = [mrels[i][1] for i in lay_ids if i in mrels]
        used = [l for l in lays if slide_use.get(l, 0) > 0]
        # 母版自身是否被计入（presentation 引用它，但可删）
        shapes = {}
        for lp in lays:
            if lp in names:
                sp = etree.fromstring(z.read(lp)).find(".//p:cSld/p:spTree", ns)
                cnt = 0 if sp is None else len(
                    [c for c in sp if etree.QName(c).localname not in ("nvGrpSpPr", "grpSpPr")])
                shapes[lp] = (etree.fromstring(z.read(lp)).find(".//p:cSld", ns).get("name"), cnt)
        report_masters.append({
            "part": mm["part"], "rId": mm["rId"],
            "layouts": lays, "layouts_used": used,
            "orphan": len(used) == 0,
            "layout_names": {k: v for k, v in shapes.items()},
        })
    return z, names, pres, prels, report_masters, slide_use, slide_to_layout


def do_analyze(deck, out_dir=None):
    z, names, pres, prels, rm, slide_use, _ = analyze(deck)
    print("# 母版/版式层分析")
    print()
    print("- 文件：`%s`" % deck)
    print("- 母版数：**%d**　slideLayout 部件数：**%d**　幻灯片数：**%d**"
          % (len(rm), len([n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)]),
             len([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)])))
    print()
    for mm in rm:
        flag = "**孤儿（可删）**" if mm["orphan"] else "保留"
        print("## %s → %s" % (mm["part"], flag))
        print()
        print("| 版式 | 名称 | 形状数 | 引用页数 |")
        print("|---|---|---|---|")
        for lp in mm["layouts"]:
            nm, cnt = mm["layout_names"].get(lp, ("?!", "?"))
            print("| `%s` | %s | %s | %d |" % (posixpath.basename(lp), nm, cnt, slide_use.get(lp, 0)))
        print()
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        p = os.path.join(out_dir, "master_analysis.json")
        with open(p, "w", encoding="utf-8") as f:
            json.dump(rm, f, ensure_ascii=False, indent=2)
        print("[dedupe-masters] 已写: %s" % p)
    return rm


# ------------------------------------------------------------------ 执行
def apply_fix(deck, out_deck, unify_titles=False, drop_stray_body=False,
              keep_empty_body_on=(), pagenum_x=None, pagenum_y=None,
              content_layout=None, content_layout_except=(), log=print):
    z, names, pres, prels, rm, slide_use, slide_to_layout = analyze(deck)

    orphans = [m for m in rm if m["orphan"]]
    if not orphans:
        log("没有孤儿母版，无需清理" + ("（仍执行 Reset 收口）" if (unify_titles or drop_stray_body) else ""))
    # ---- 主题血统：被删母版的主题（dead_themes） vs 存活母版的主题（live_theme）
    # 陷阱（本轮实测踩到）：presentation.xml.rels 的"演示文稿级主题"经常还指在
    # 旧母版的主题上。此时该 theme 被**旧母版 + presentation 同时引用**，引用计数=2，
    # 单纯按"引用计数为 0"做的孤儿扫描会判成"有人用"→ 既不删、也不重指 →
    # 产物里主题库残留一个过期主题，新建形状/取色取字体仍继承旧主题，
    # 用户下次"重设版式"又被打回原形。
    # 判定必须按**血统**（这个 theme 挂在被删母版下），不能只按引用计数。
    dead_themes = {t for mm in orphans
                   for _, (typ, t) in read_rels(z, mm["part"]).items() if typ == "theme"}
    dead_themes |= {t for mm in orphans for lp in mm["layouts"]
                    for _, (typ, t) in read_rels(z, lp).items() if typ == "theme"}
    live_theme = None
    for mm in rm:
        if mm["orphan"]:
            continue
        for _, (typ, t) in read_rels(z, mm["part"]).items():
            if typ == "theme":
                live_theme = t
                break
        if live_theme:
            break
    if live_theme is None:
        for r in etree.fromstring(z.read(PRES_RELS)):
            if r.get("Type", "").endswith("/theme"):
                live_theme = resolve(PRES, r.get("Target", ""))
                break
    pres_theme = None
    for r in etree.fromstring(z.read(PRES_RELS)):
        if r.get("Type", "").endswith("/theme"):
            pres_theme = resolve(PRES, r.get("Target", ""))
            break
    repoint_pres_theme = bool(live_theme and pres_theme
                              and pres_theme in dead_themes and pres_theme != live_theme)
    if repoint_pres_theme:
        log("主题血统：presentation 主题 %s 挂在被删母版下 -> 将改指 %s"
            % (pres_theme, live_theme))

    doomed = set()
    for mm in orphans:
        doomed.add(mm["part"])
        doomed.add(part_rels_path(mm["part"]))
        for lp in mm["layouts"]:
            doomed.add(lp)
            doomed.add(part_rels_path(lp))
        # 孤儿母版专属的 theme / tags / embeddings / media
        own = {t for _, (typ, t) in read_rels(z, mm["part"]).items()
               if typ in ("theme", "tags", "oleObject", "image", "font")}
        for lp in mm["layouts"]:
            for _, (typ, t) in read_rels(z, lp).items():
                if typ in ("theme", "tags", "oleObject", "image", "font"):
                    own.add(t)
        # 只有"除孤儿这条线外无人引用"的才删 —— 全包扫一遍才敢下这个结论
        skip = set(doomed)
        if repoint_pres_theme:
            # build_referenced 的 skip 按**宿主部件**比对，要塞 PRES 而不是 PRES_RELS
            skip.add(PRES)          # 该 rels 的主题引用即将被改写，旧引用不算数
        ref = build_referenced(z, names, skip=skip)
        if repoint_pres_theme:
            ref.add(live_theme)     # 重指后 presentation 引的是存活母版的主题
        for t in own:
            if t in names and t not in ref:
                doomed.add(t)
                doomed.add(part_rels_path(t))
    doomed = {d for d in doomed if d in names}

    # ---- presentation.xml：移除孤儿母版的 sldMasterId
    oids = {m["rId"] for m in orphans}
    px = etree.fromstring(z.read(PRES))
    for m in list(px.findall(".//p:sldMasterIdLst/p:sldMasterId", ns)):
        if m.get("{%s}id" % R) in oids:
            m.getparent().remove(m)
    pres_new = to_bytes(px)

    # ---- presentation.xml.rels：移除孤儿母版关系；theme 若指向被删的 theme 则改指存活者
    prx = etree.fromstring(z.read(PRES_RELS))
    # live_theme / dead_themes 已在前面按"主题血统"算出
    for r in list(prx):
        if r.get("Id") in oids:
            prx.remove(r)
            log("presentation.xml.rels: 删除 %s" % r.get("Id"))
        elif r.get("Type", "").endswith("/theme"):
            cur = resolve(PRES, r.get("Target", ""))
            # 这里**不能**只写 `cur in doomed`：旧主题常被 presentation 与旧母版同时引用，
            # 血统判定才对。只在重指之后 `cur in doomed` 才成立（本行两条件并存即为此）。
            if live_theme and cur != live_theme and (cur in dead_themes or cur in doomed):
                # 宿主是 **presentation.xml**，不是 .rels 文件；
                # 写成 PRES_RELS 会把 "theme/theme1.xml" 解析成
                # "ppt/_rels/theme/theme1.xml"，判定永远不命中（原脚本埋的坑）。
                r.set("Target", posixpath.relpath(live_theme, posixpath.dirname(PRES)))
                log("presentation.xml.rels: theme -> %s" % r.get("Target"))
    prels_new = to_bytes(prx)

    # ---- 版式 rels 里指向被删母版的，随版式一起删；存活版式不动（它们指向存活母版）
    # ---- Content_Types：删掉被删部件的 Override
    cx = etree.fromstring(z.read(CT_PART))
    n_ct = 0
    for ov in list(cx.findall("{%s}Override" % CTNS)):
        pn = (ov.get("PartName") or "").lstrip("/")
        if pn in doomed:
            cx.remove(ov)
            n_ct += 1
    ct_new = to_bytes(cx)
    log("Content_Types: 删除 %d 个 Override" % n_ct)

    # ---- slide 层收口
    slide_new, rel_new = {}, {}
    if unify_titles or drop_stray_body or pagenum_x is not None:
        keep = {k.strip() for k in keep_empty_body_on if k.strip()}
        for n in sorted(names):
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
            if not m:
                continue
            i = m.group(1)
            x = etree.fromstring(z.read(n))
            spTree = x.find(".//p:spTree", ns)
            changed = []
            lay = slide_to_layout.get(n)
            lay_name = ""
            if lay and lay in names:
                lay_name = etree.fromstring(z.read(lay)).find(".//p:cSld", ns).get("name") or ""
            is_keep = lay_name in keep

            for sp in list(spTree):
                if etree.QName(sp).localname not in ("sp", "pic", "graphicFrame"):
                    continue
                ph = sp.find(".//p:nvPr/p:ph", ns)
                if ph is None:
                    continue
                t = ph.get("type", "body")

                if t == "sldNum" and pagenum_x is not None:
                    xf = sp.find(".//a:xfrm", ns)
                    off = xf.find("a:off", ns) if xf is not None else None
                    if off is not None and off.get("x") != str(pagenum_x):
                        changed.append("页码x %s->%s" % (off.get("x"), pagenum_x))
                        off.set("x", str(pagenum_x))
                        if pagenum_y is not None:
                            off.set("y", str(pagenum_y))

                if unify_titles and t in ("title", "ctrTitle"):
                    spPr = sp.find("p:spPr", ns)
                    if spPr is not None:
                        xf = spPr.find("a:xfrm", ns)
                        if xf is not None:
                            spPr.remove(xf)
                            changed.append("标题删xfrm")

                if drop_stray_body and t == "body" and not is_keep:
                    txt = "".join(e.text or "" for e in sp.iter("{%s}t" % A)).strip()
                    if not txt and sp.find(".//a:blip", ns) is None:
                        spTree.remove(sp)
                        changed.append("删空body(idx=%s)" % ph.get("idx"))

            if changed:
                slide_new[n] = to_bytes(x)
                log("  slide%s  %s" % (i, ", ".join(changed)))

    # ---- 内容页版式归位（把内容页统一挂到模板的内容版式）
    if content_layout:
        target = None
        for mm in rm:
            if mm["orphan"]:
                continue
            for lp, (nm, _c) in mm["layout_names"].items():
                if nm == content_layout:
                    target = lp
                    break
            if target:
                break
        if not target:
            log("!! 找不到名为「%s」的版式，跳过内容页版式归位" % content_layout)
        else:
            exc = {int(s) for s in content_layout_except if str(s).strip().isdigit()}
            n_sw = 0
            for n in sorted(names):
                m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
                if not m:
                    continue
                i = int(m.group(1))
                if i in exc:
                    continue
                rp = part_rels_path(n)
                if rp not in names:
                    continue
                rx = etree.fromstring(z.read(rp))
                hit = False
                for r in rx:
                    if r.get("Type", "").endswith("/slideLayout"):
                        if posixpath.basename(resolve(n, r.get("Target", ""))) != \
                                posixpath.basename(target):
                            r.set("Target", posixpath.relpath(target, posixpath.dirname(n)))
                            hit = True
                            n_sw += 1
                if hit:
                    rel_new[rp] = to_bytes(rx)
            log("版式归位：%d 页 -> %s" % (n_sw, content_layout))

    # ---- 写包
    if os.path.exists(out_deck):
        os.remove(out_deck)
    zo = zipfile.ZipFile(out_deck, "w", zipfile.ZIP_DEFLATED)
    zo.writestr(zipfile.ZipInfo(CT_PART), ct_new)          # CT 必须第一个条目
    for item in z.infolist():
        n = item.filename
        if n == CT_PART or n in doomed or n.endswith("/"):
            continue
        if n == PRES:
            zo.writestr(item, pres_new)
        elif n == PRES_RELS:
            zo.writestr(item, prels_new)
        elif n in slide_new:
            zo.writestr(item, slide_new[n])
        elif n in rel_new:
            zo.writestr(item, rel_new[n])
        else:
            zo.writestr(item, z.read(n))
    zo.close()
    z.close()
    log("产出 %s (%.1f MB)" % (out_deck, os.path.getsize(out_deck) / 1048576))
    return out_deck


def main():
    ap = argparse.ArgumentParser(description="母版/版式层收口（换模板后必做）")
    ap.add_argument("deck")
    ap.add_argument("--out", default=None, help="分析结果的落盘目录")
    ap.add_argument("--out-deck", default=None, help="执行清理并输出到该文件")
    ap.add_argument("--unify-titles", action="store_true",
                    help="标题占位符删 <a:xfrm>，位置回归版式")
    ap.add_argument("--drop-stray-body", action="store_true",
                    help="删掉空的 body 占位符（Reset 残留）")
    ap.add_argument("--keep-empty-body-on", default="",
                    help="豁免的版式名（逗号分隔），如 \"目录,Workshop index\"")
    ap.add_argument("--pagenum-x", type=int, default=None, help="页码 off.x（EMU）")
    ap.add_argument("--pagenum-y", type=int, default=None, help="页码 off.y（EMU）")
    ap.add_argument("--content-layout", default=None,
                    help="把内容页统一挂到该版式名，如 \"标题和文本\"")
    ap.add_argument("--content-layout-except", default="",
                    help="不改版式的页码（逗号分隔），如 \"1,2,3,7,10,32,34\"")
    args = ap.parse_args()

    if args.out_deck:
        apply_fix(args.deck, args.out_deck,
                  unify_titles=args.unify_titles,
                  drop_stray_body=args.drop_stray_body,
                  keep_empty_body_on=args.keep_empty_body_on.split(",") if args.keep_empty_body_on else (),
                  pagenum_x=args.pagenum_x, pagenum_y=args.pagenum_y,
                  content_layout=args.content_layout,
                  content_layout_except=args.content_layout_except.split(",") if args.content_layout_except else ())
    else:
        do_analyze(args.deck, args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
