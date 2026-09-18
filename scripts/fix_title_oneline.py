# -*- coding: utf-8 -*-
"""fix_title_oneline.py — 标题强制单行：放不下就缩字号，绝不换行。

为什么需要它
------------
模板的标题框是**一行高**的（`标题和文本` 版式：框高 50.5pt、字号 28pt、容行数 1）。
标题一超宽就会折成两行，第二行会**穿过标题下的分隔线**，并且把下面的内容顶下去 ——
模板设计当场破掉。目录/分节名/序号那几个槽位同理，框也只够一行。

做法：**按真实字体度量算出需要的缩放比，把 `fontScale` 写到页面上。**
为什么不是"在版式上放个自动缩排"：实测无效 —— PowerPoint 的溢出缩排**只在编辑时计算**，
并把结果写回页面的 `bodyPr`；只在版式里放一个空的 `normAutofit`，显示时不会重算。
所以这里自己算好、写进页面，等价于用户在 PowerPoint 里打开"溢出时缩排文字"后的结果。

为什么默认用 `fontScale` 而不是改写 `sz`
--------------------------------------
  · `fontScale`：**声明字号仍是模板值**（28pt），只在渲染时等比缩小。
    于是 `verify_template_fidelity.py`（R7）与 `verify_pptx.py` 的"标题回归版式槽位"
    两条体检都照样通过 —— 改的是"这一页怎么显示"，不是"模板长什么样"。
  · `sz`（`--mode sz`）：直接改 run 的字号，是**页面级硬覆盖**；标题文字如果以后改短了，
    字号不会自己恢复。要硬覆盖就显式选它，别当默认。

硬边界（不缩）
-------------
  · 标题里带**硬换行**（`<a:br/>` 或人工分段）：缩字号救不了，标 `HARD` 让人改文字；
  · 需要的缩放比 < `--min-scale`（默认 0.70，即 28pt→19.6pt）：**不缩**，标 `UNRESOLVED`
    并给出"要压到单行需删到 N 字符"。字号是版面档位，缩到看不清等于把合格改成不合格。

用法
----
    # 体检（只读）：列出所有会换行的标题 + 需要缩到多少
    python fix_title_oneline.py "<deck.pptx>" --check

    # 修（默认 autofit 模式，只写页面，不碰版式/母版）
    python fix_title_oneline.py "<deck.pptx>" --out "<new.pptx>"

    # 硬改字号
    python fix_title_oneline.py "<deck.pptx>" --out "<new.pptx>" --mode sz

    # 放宽下限 / 收紧安全余量
    python fix_title_oneline.py "<deck.pptx>" --out "<new.pptx>" --min-scale 0.6 --margin 0.05

产出后**必须**跑 `verify_title_lines.py`（PowerPoint 自己数行数）复核。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_title_scope as TS                                    # noqa: E402

A = TS.A
P = TS.P
NS = TS.NS
CT_PART = "[Content_Types].xml"
# bodyPr 子元素顺序（CT_TextBodyProperties）：autofit 必须排在 prstTxWarp 之后、
# scene3d 之前，否则 PowerPoint 会因为元素顺序错误拒收。
BODYPR_AFTER_WARP = ("scene3d", "sp3d", "flatTx", "extLst")


def clear_autofit(bodypr):
    """去掉 autofit 节点 —— 文字回到"完全按声明字号显示"（即恢复模板字号）。"""
    n = 0
    for c in list(bodypr):
        if etree.QName(c).localname in ("noAutofit", "normAutofit", "spAutoFit"):
            bodypr.remove(c)
            n += 1
    return n


def set_autofit(bodypr, font_scale):
    """把 bodyPr 的 autofit 设成 normAutofit/fontScale，位置合规。"""
    for c in list(bodypr):
        if etree.QName(c).localname in ("noAutofit", "normAutofit", "spAutoFit"):
            bodypr.remove(c)
    node = etree.Element("{%s}normAutofit" % A)
    node.set("fontScale", str(int(font_scale)))
    node.set("lnSpcReduction", "0")
    pos = None
    for i, c in enumerate(bodypr):
        if etree.QName(c).localname in BODYPR_AFTER_WARP:
            pos = i
            break
    if pos is None:
        bodypr.append(node)
    else:
        bodypr.insert(pos, node)
    return node


def set_no_autofit(bodypr):
    for c in list(bodypr):
        if etree.QName(c).localname in ("noAutofit", "normAutofit", "spAutoFit"):
            bodypr.remove(c)
    etree.SubElement(bodypr, "{%s}noAutofit" % A)


def ensure_bodypr(txbody):
    bp = txbody.find("{%s}bodyPr" % A)
    if bp is None:
        bp = etree.Element("{%s}bodyPr" % A)
        txbody.insert(0, bp)
    return bp


def apply_sz_mode(txbody, size_pt):
    """把段落/运行的 rPr 都写成指定字号，并加 noAutofit（不再让 PowerPoint 自动缩）。"""
    val = str(int(round(size_pt * 100)))
    n = 0
    for rpr in txbody.iter("{%s}rPr" % A):
        rpr.set("sz", val)
        n += 1
    for d in txbody.iter("{%s}defRPr" % A):
        d.set("sz", val)
        n += 1
    for e in txbody.iter("{%s}endParaRPr" % A):
        e.set("sz", val)
        n += 1
    set_no_autofit(ensure_bodypr(txbody))
    return n


def analyse(z, index, args):
    """算出每一处需要处理的标题。返回 (rows, targets)"""
    rows = []
    for part in sorted(index["slides"], key=lambda s: index["slides"][s]["number"]):
        for t in TS.iter_targets(z, index, part,
                                 include_heading=not args.only_title,
                                 include_subtitle=args.include_subtitle,
                                 slots=set(args.slots.split(",")) if args.slots else None):
            ff, exact = TS.find_font(t["font"], args.font_map)
            size = t["size"]
            need = TS.measure_width_pt(t["text"], size, ff) if size else 0.0
            avail = t["avail_w"] or 0.0
            ratio = (need / avail) if avail else 0.0
            cur = t["font_scale"] or 100000
            # 字体度量不精确（退回同族其它字重或纯估算）时，余量翻倍
            margin = args.margin if exact else args.margin * 2.0
            # 基准必须是**声明字号**：算出"刚好单行"该用的缩放比，这就是目标值。
            # （不能用"当前有效字号"当基准 —— 那样永远算不出"缩过头"，见下方 STALE）
            if need <= avail or need <= 0:
                needed = 100000
            else:
                needed = int(round(min(1.0, (avail / need) * (1.0 - margin)) * 100000))
            final_size = (size * needed / 100000.0) if size else None
            # 差 0.2% 以内视为"已经是目标值"（避免浮点+取整导致每次都报一次）
            tol = 200

            if t["hard_break"]:
                verdict = "HARD"
                action = "标题含硬换行 —— 缩字号无效，必须改文字"
            elif abs(cur - needed) <= tol:
                verdict, action = "OK", ""
                needed = cur                   # 已达标：不写文件
            elif cur > needed:
                # 缩得不够（或还没缩）→ 缩到 needed
                if size and final_size < args.min_scale * size:
                    verdict = "UNRESOLVED"
                    per_char = need / max(1, len(t["text"]))
                    keep = int(avail / per_char * 0.97) if per_char else 0
                    action = ("需缩到 %.1fpt 才单行，低于下限 %.1fpt（%d%%）—— 不缩。"
                              "要单行请把标题压到约 %d 字符内" %
                              (final_size, args.min_scale * size,
                               int(args.min_scale * 100), keep))
                else:
                    verdict, action = "SHRINK", ""
            else:
                # cur < needed：缩过头了（上游写死，或标题后来改短了）
                # 会把"合格"变成"没必要的小字"，必须报出来，但默认不改观感
                verdict = "STALE"
                if needed >= 100000:
                    action = "缩排过头：可按模板值恢复 %.1fpt（去掉 fontScale）" % size
                else:
                    action = ("缩排过头：可放大回 %.1fpt（现 %.1fpt）"
                              % (size * needed / 100000.0, size * cur / 100000.0))
            rows.append(dict(t, need=need, avail=avail, ratio=ratio, scale=needed,
                             final_size=final_size, verdict=verdict, action=action,
                             cur_scale=cur, font_file=ff, font_exact=exact))
    return rows


def main():
    ap = argparse.ArgumentParser(description="标题强制单行（放不下就缩字号）")
    ap.add_argument("deck")
    ap.add_argument("--out", default=None, help="输出文件；不给则只体检不写")
    ap.add_argument("--check", action="store_true", help="只体检（默认行为，显式写更清楚）")
    ap.add_argument("--mode", default="autofit", choices=["autofit", "sz"],
                    help="autofit=写 fontScale（声明字号不变，默认）；sz=硬改 run 字号")
    ap.add_argument("--margin", type=float, default=0.03,
                    help="安全余量（默认 0.03）。算成恰好等于框宽时 PowerPoint 仍会换行")
    ap.add_argument("--min-scale", type=float, default=0.70,
                    help="缩放下限（默认 0.70）。低于此值不缩，改报 UNRESOLVED")
    ap.add_argument("--restore-stale", action="store_true",
                    help="把「缩过头」的标题（fontScale 比需要的还小，或标题已改短）放大回"
                         "刚好单行的字号。默认只报告不改 —— 改字号是看得见的观感变化")
    ap.add_argument("--only-title", action="store_true",
                    help="只处理 title/ctrTitle 占位符，不碰目录/分节名等标题槽位")
    ap.add_argument("--include-subtitle", action="store_true",
                    help="连封面副标题一起处理（副标题换行通常是设计意图，默认不动）")
    ap.add_argument("--slots", default=None,
                    help='只处理指定槽位，如 "body:17,body:18"')
    ap.add_argument("--font-map", default=None,
                    help='字体文件覆盖，如 "noto sans=C:\\path\\NotoSans-Regular.ttf"')
    ap.add_argument("--json", default=None)
    args = ap.parse_args()

    if args.font_map:
        extra = {}
        for kv in args.font_map.split(";"):
            if "=" in kv:
                k, v = kv.split("=", 1)
                extra[k.strip().lower()] = v.strip()
        args.font_map = extra
    else:
        args.font_map = None

    if not os.path.isfile(args.deck):
        print("ERROR: 找不到文件 %s" % args.deck, file=sys.stderr)
        return 2

    z = zipfile.ZipFile(args.deck)
    index = TS.build_index(z)
    rows = analyse(z, index, args)

    print("=" * 116)
    print("标题单行体检：%s" % os.path.basename(args.deck))
    print("版式 %d 个；标题槽位（模板驱动）%s" % (
        len(index["layouts"]),
        ", ".join("%s:%s" % k for k in sorted(index["heading_slots"])) or "无"))
    print("=" * 116)
    print("%-5s %-22s %-7s %-8s %-8s %-6s %-8s %-8s %-11s %s" % (
        "页", "形状", "字号pt", "可用pt", "需要pt", "比值", "缩放", "缩放后pt", "判定", "标题"))
    print("-" * 116)
    todo = []
    for r in rows:
        if r["verdict"] == "OK":
            continue
        todo.append(r)
        print("%-5d %-22s %-7.1f %-8.1f %-8.1f %-6.3f %-8s %-8s %-11s %s" % (
            r["page"], r["name"][:22], r["size"] or 0, r["avail"], r["need"],
            r["ratio"], ("%.3f" % (r["scale"] / 100000.0)) if r["scale"] < 100000 else "-",
            ("%.1f" % r["final_size"]) if r["final_size"] else "-",
            r["verdict"], r["text"][:44]))
    print("-" * 116)
    n_ok = len(rows) - len(todo)
    by = {}
    for r in todo:
        by[r["verdict"]] = by.get(r["verdict"], 0) + 1
    print("纳入判定 %d 处；已合规 %d 处；待处理 %d 处 %s" % (
        len(rows), n_ok, len(todo), by or ""))
    for r in todo:
        if r["verdict"] in ("HARD", "UNRESOLVED"):
            print("  [%s] P%-3d %s：%s" % (r["verdict"], r["page"], r["text"][:60], r["action"]))
        elif r["verdict"] == "STALE":
            print("  [STALE] P%-3d %s：%s" % (r["page"], r["text"][:60], r["action"]))
    if any(r["verdict"] == "STALE" for r in todo) and not args.restore_stale:
        print("  （STALE 需显式 --restore-stale 才会写回放大；默认只报告，不改观感）")

    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump(rows, fh, ensure_ascii=False, indent=2, default=str)
        print("JSON -> %s" % args.json)

    if not args.out:
        z.close()
        print("\n（体检模式，未输出文件）")
        return 0

    # ---------------- 写入 ----------------
    # 按部件分组，整篇重解析后按 cNvPr id 定位形状 —— 直接改 XML 片段会丢命名空间声明
    want = {"SHRINK"}
    if args.restore_stale:
        want.add("STALE")
    todo = {}
    for r in rows:
        if r["verdict"] in want:
            todo.setdefault(r["part"], []).append(r)

    changed = {}
    n_shrink = n_sz = n_restore = 0
    PSTR = "{%s}" % P
    for part, rs in todo.items():
        root = etree.fromstring(z.read(part))
        by_id = {}
        for sp in root.findall(".//%sspTree/%ssp" % (PSTR, PSTR)):
            cid = sp.find(".//%scNvPr" % PSTR)
            if cid is not None:
                by_id[cid.get("id")] = sp
        touched = 0
        for r in rs:
            sp = by_id.get(r["shape_id"])
            if sp is None:
                print("  !! P%d %s：找不到形状 id=%s，跳过"
                      % (r["page"], r["name"], r["shape_id"]))
                continue
            txbody = sp.find("%stxBody" % PSTR)
            if txbody is None:
                print("  !! P%d %s：无 txBody，跳过" % (r["page"], r["name"]))
                continue
            if r["verdict"] == "STALE":
                bp = ensure_bodypr(txbody)
                if r["scale"] >= 100000:
                    clear_autofit(bp)          # 完全恢复模板字号
                    n_restore += 1
                else:
                    set_autofit(bp, r["scale"])
                    n_restore += 1
                print("  P%-3d %-22s 恢复 %.1fpt -> %.1fpt" % (
                    r["page"], r["name"][:22], r["size"] * r["cur_scale"] / 100000.0,
                    r["final_size"]))
            elif args.mode == "autofit":
                set_autofit(ensure_bodypr(txbody), r["scale"])
                n_shrink += 1
            else:
                apply_sz_mode(txbody, r["final_size"])
                n_sz += 1
            touched += 1
            if r["verdict"] == "SHRINK":
                print("  P%-3d %-22s %.1fpt -> %.1fpt  (fontScale=%d)" % (
                    r["page"], r["name"][:22], r["size"], r["final_size"], r["scale"]))
        if touched:
            changed[part] = etree.tostring(root, xml_declaration=True,
                                           encoding="UTF-8", standalone=True)

    print("\n=== 写入（mode=%s）===" % args.mode)
    print("缩排 %d 处；硬改字号 %d 处；恢复 %d 处；改动部件 %d 个"
          % (n_shrink, n_sz, n_restore, len(changed)))

    if os.path.exists(args.out):
        os.remove(args.out)
    zo = zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED)
    zo.writestr(zipfile.ZipInfo(CT_PART), z.read(CT_PART))
    for item in z.infolist():
        n = item.filename
        if n == CT_PART or n.endswith("/"):
            continue
        if n in changed:
            zo.writestr(item, changed[n])
        else:
            zo.writestr(item, z.read(n))
    zo.close()
    z.close()
    print("\n产出 %s (%.1f MB)" % (args.out, os.path.getsize(args.out) / 1048576))
    return 0


if __name__ == "__main__":
    sys.exit(main())
