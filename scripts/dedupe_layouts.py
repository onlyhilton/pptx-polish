# -*- coding: utf-8 -*-
"""版式去重分析/执行器

用途：PPT 长期迭代后，包里会堆积重复或未使用的 slideLayout，既涨体积，
      也让"新建幻灯片"面板出现大量同名版式。本脚本做两件事：

  1) 分析（默认）：给每个 slideLayout 算"结构指纹"，分组报告重复情况，
     并统计每页引用了哪个版式。只读，不改文件。
  2) 执行（--apply）：把重复组统一到保留者，把未被引用的版式删除，
     同时修正 presentation.xml.rels / [Content_Types].xml / 各 slide 的 rels。

指纹维度（顺序即判定优先级）：
  A. 占位符结构：[(ph_type, idx, 左, 上, 宽, 高), ...]
  B. 非占位符形状：[(name, 左, 上, 宽, 高, 是否图片, 图片媒体名), ...]
  C. 背景：<p:bg> 规范化后的 XML
  D. 继承的母版

用法：
    python dedupe_layouts.py "<deck.pptx>"                        # 只分析
    python dedupe_layouts.py "<deck.pptx>" --out "<dir>"          # 分析 + 落 JSON/MD
    python dedupe_layouts.py "<deck.pptx>" --apply --out-deck "<new.pptx>"
"""
from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# 指纹时忽略的属性（每次保存都可能变，与视觉无关）
VOLATILE_ATTRS = {"id", "name"}


def q(tag: str):
    pre, local = tag.split(":")
    return "{%s}%s" % (NS[pre], local)


def read_xml(blob: bytes):
    return etree.fromstring(blob)


def to_bytes(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def resolve(host_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(host_part), target))


def relativize(host_part: str, abs_target: str) -> str:
    base = posixpath.dirname(host_part)
    return posixpath.relpath(abs_target, base) if base else abs_target


def num(v, default=0):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------- 指纹


def ph_signature(root) -> list:
    """占位符结构签名"""
    out = []
    for sp in root.iter(q("p:sp")):
        ph = sp.find(".//" + q("p:ph"))
        if ph is None:
            continue
        t = ph.get("type") or "body"
        idx = ph.get("idx") or ""
        xfrm = sp.find(".//" + q("a:xfrm"))
        geo = None
        if xfrm is not None:
            off = xfrm.find(q("a:off"))
            ext = xfrm.find(q("a:ext"))
            if off is not None and ext is not None:
                geo = (num(off.get("x")), num(off.get("y")),
                       num(ext.get("cx")), num(ext.get("cy")))
        out.append((t, idx, geo))
    # 占位符出现在 XML 里的顺序不稳定，按 (type, idx) 排序后再比
    return sorted(out, key=lambda r: (r[0], r[1], str(r[2])))


def shape_signature(root, rels_map: dict) -> list:
    """非占位符形状签名（含图片指向的媒体文件名）"""
    out = []
    for tag in ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp"):
        for el in root.iter(q(tag)):
            if el.find(".//" + q("p:ph")) is not None:
                continue                      # 占位符已在 A 里算过
            name = ""
            cNvPr = el.find(".//" + q("p:cNvPr"))
            if cNvPr is not None:
                name = cNvPr.get("name") or ""
            xfrm = el.find(".//" + q("a:xfrm"))
            geo = None
            if xfrm is not None:
                off = xfrm.find(q("a:off"))
                ext = xfrm.find(q("a:ext"))
                if off is not None and ext is not None:
                    geo = (num(off.get("x")), num(off.get("y")),
                           num(ext.get("cx")), num(ext.get("cy")))
            media = ""
            blip = el.find(".//" + q("a:blip"))
            if blip is not None:
                rid = blip.get("{%s}embed" % REL_NS) or blip.get("{%s}link" % REL_NS)
                if rid:
                    media = rels_map.get(rid, "")
                    media = posixpath.basename(media)
            out.append((tag, geo, media))
    out.sort(key=lambda r: (r[0], str(r[1]), r[2]))
    return out


def bg_signature(root):
    bg = root.find("." + q("p:cSld") + "/" + q("p:bg"))
    if bg is None:
        return None
    # 去掉 blip 的 r:embed（会被重编号），保留媒体名
    x = etree.tostring(bg, encoding="unicode")
    x = re.sub(r'\s*r:(?:embed|link)="[^"]*"', "", x)
    return x


def bg_media(root, rels_map: dict):
    bg = root.find("." + q("p:cSld") + "/" + q("p:bg"))
    if bg is None:
        return []
    out = []
    for blip in bg.iter(q("a:blip")):
        rid = blip.get("{%s}embed" % REL_NS)
        if rid and rid in rels_map:
            out.append(posixpath.basename(rels_map[rid]))
    return out


def fingerprint(sig: dict) -> str:
    blob = json.dumps(sig, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()[:12]


# ---------------------------------------------------------------- 分析


def load_deck(pptx: Path) -> dict:
    z = zipfile.ZipFile(pptx)
    names = set(z.namelist())

    # 每个部件的 rels
    def rels_of(part: str) -> dict:
        rp = posixpath.join(posixpath.dirname(part), "_rels",
                            posixpath.basename(part) + ".rels")
        if rp not in names:
            return {}
        out = {}
        root = read_xml(z.read(rp))
        for rel in root:
            out[rel.get("Id")] = resolve(part, rel.get("Target"))
        return out

    # 装配关系：presentation.xml -> masters/layouts
    pres_rels = {}
    if "ppt/_rels/presentation.xml.rels" in names:
        root = read_xml(z.read("ppt/_rels/presentation.xml.rels"))
        for rel in root:
            pres_rels[rel.get("Id")] = resolve("ppt/presentation.xml", rel.get("Target"))

    pres = read_xml(z.read("ppt/presentation.xml"))
    master_parts = []
    for e in pres.iter(q("p:sldMasterId")):
        rid = e.get("{%s}id" % REL_NS)
        if rid in pres_rels:
            master_parts.append(pres_rels[rid])

    layouts = []
    for mp in master_parts:
        mrels = rels_of(mp)
        mroot = read_xml(z.read(mp))
        for e in mroot.iter(q("p:sldLayoutId")):
            rid = e.get("{%s}id" % REL_NS)
            lp = mrels.get(rid)
            if lp:
                layouts.append({"part": lp, "master": mp})

    # 每页引用的版式
    slide_layout = {}
    for n in sorted(x for x in names if re.match(r"ppt/slides/slide\d+\.xml$", x)):
        srels = rels_of(n)
        sroot = read_xml(z.read(n))
        lay = None
        for e in sroot.iter(q("p:sldLayoutId")):
            lay = srels.get(e.get("{%s}id" % REL_NS))
        if lay is None:
            rel = sroot.find("." + q("p:cSld"))
            # 通过 rels 里 Type 找 layout
            rp = posixpath.join(posixpath.dirname(n), "_rels",
                                posixpath.basename(n) + ".rels")
            if rp in names:
                rroot = read_xml(z.read(rp))
                for rel in rroot:
                    if rel.get("Type", "").endswith("/slideLayout"):
                        lay = resolve(n, rel.get("Target"))
        slide_layout[n] = lay

    return {
        "zip": z,
        "names": names,
        "pres": pres,
        "pres_rels": pres_rels,
        "master_parts": master_parts,
        "layouts": layouts,
        "slide_layout": slide_layout,
        "rels_of": rels_of,
    }


def analyze(deck: dict) -> dict:
    z = deck["zip"]
    info = []
    for L in deck["layouts"]:
        part = L["part"]
        if part not in deck["names"]:
            continue
        root = read_xml(z.read(part))
        rm = deck["rels_of"](part)
        cSld = root.find("." + q("p:cSld"))
        name = cSld.get("name") if cSld is not None else None
        phs = ph_signature(root)
        shps = shape_signature(root, rm)
        bg = bg_signature(root)
        sig = {
            "ph": [[t, i, list(g) if g else None] for t, i, g in phs],
            "shapes": [[t, list(g) if g else None, m] for t, g, m in shps],
            "bg": bg,
            "bgmedia": bg_media(root, rm),
            "type": root.get("type") or "cust",
        }
        info.append({
            "part": part,
            "index": int(re.search(r"slideLayout(\d+)", part).group(1)),
            "name": name,
            "type": root.get("type") or "cust",
            "master": L["master"],
            "ph": [[t, i] for t, i, _g in phs],
            "ph_count": len(phs),
            "shape_count": len(shps),
            "bg_kind": ("image" if "blipFill" in (bg or "") else
                        ("solid" if "solidFill" in (bg or "") else
                         ("inherit" if bg is None else "other"))),
            "bg_media": bg_media(root, rm),
            "fp": fingerprint(sig),
            "size": z.getinfo(part).file_size,
        })

    # 分组：同母版 + 同指纹
    groups = defaultdict(list)
    for it in info:
        groups[(it["master"], it["fp"])].append(it)

    refs = defaultdict(list)
    for slide_part, lay in deck["slide_layout"].items():
        if lay:
            refs[lay].append(int(re.search(r"slide(\d+)", slide_part).group(1)))

    for it in info:
        it["ref_count"] = len(refs.get(it["part"], []))
        it["ref_slides"] = sorted(refs.get(it["part"], []))

    dup_groups = []
    for (m, fp), members in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        if len(members) > 1:
            total = sum(x["size"] for x in members)
            keep = max(members, key=lambda x: (x["ref_count"], -x["index"]))
            dup_groups.append({
                "fingerprint": fp,
                "master": m,
                "members": [{"index": x["index"], "part": x["part"], "name": x["name"],
                             "ref_count": x["ref_count"], "ref_slides": x["ref_slides"],
                             "size": x["size"]} for x in members],
                "keep": keep["part"],
                "removable": [x["part"] for x in members if x["part"] != keep["part"]],
                "wasted_bytes": total - keep["size"],
            })

    unused = [it for it in info if it["ref_count"] == 0]

    return {
        "layouts": info,
        "dup_groups": dup_groups,
        "unused": [{"index": x["index"], "part": x["part"], "name": x["name"],
                    "size": x["size"], "type": x["type"]} for x in unused],
        "slide_refs": {int(re.search(r"slide(\d+)", k).group(1)): v
                       for k, v in deck["slide_layout"].items()},
    }


def report_md(res: dict, pptx: Path) -> str:
    L = []
    L.append("# 版式去重分析\n")
    L.append("- 文件：`%s`" % pptx)
    L.append("- 版式总数：**%d**　重复组：**%d**　未被引用：**%d**\n"
             % (len(res["layouts"]), len(res["dup_groups"]), len(res["unused"])))

    L.append("## 一、重复版式组\n")
    if not res["dup_groups"]:
        L.append("无结构完全相同的版式。\n")
    for gi, g in enumerate(res["dup_groups"], 1):
        L.append("### 组 %d（指纹 `%s`）" % (gi, g["fingerprint"]))
        L.append("保留 `%s`，可删 %d 个，回收 %.1f KB\n"
                 % (posixpath.basename(g["keep"]), len(g["removable"]),
                    g["wasted_bytes"] / 1024))
        L.append("| idx | 名称 | 页面引用 | 尺寸 | 处置 |")
        L.append("|---|---|---|---|---|")
        for m in g["members"]:
            act = "保留" if m["part"] == g["keep"] else "合并删除"
            L.append("| %d | %s | %d 页 %s | %.1f KB | %s |" % (
                m["index"], m["name"] or "-", m["ref_count"],
                m["ref_slides"] if m["ref_slides"] else "（无）",
                m["size"] / 1024, act))
        L.append("")

    L.append("## 二、未被任何页面引用\n")
    if not res["unused"]:
        L.append("无。\n")
    else:
        L.append("| idx | 名称 | 类型 | 尺寸 |")
        L.append("|---|---|---|---|")
        for u in res["unused"]:
            L.append("| %d | %s | %s | %.1f KB |" % (u["index"], u["name"] or "-",
                                                     u["type"], u["size"] / 1024))
        L.append("")
        L.append("> 合计 %.1f KB\n" % (sum(u["size"] for u in res["unused"]) / 1024))

    L.append("## 三、全部版式清单\n")
    L.append("| idx | 名称 | 类型 | 占位符 | 形状 | 背景 | 引用页 |")
    L.append("|---|---|---|---|---|---|---|")
    for it in sorted(res["layouts"], key=lambda x: x["index"]):
        L.append("| %d | %s | %s | %s | %d | %s | %d |" % (
            it["index"], it["name"] or "-", it["type"],
            ",".join(sorted({t for t, _ in it["ph"]})) or "-",
            it["shape_count"], it["bg_kind"], it["ref_count"]))
    L.append("")
    return "\n".join(L)


# ---------------------------------------------------------------- 执行


def apply_dedupe(src: Path, res: dict, out: Path,
                 targets: list | None = None,
                 drop_media: bool = True,
                 redirect_refs: bool = True) -> dict:
    """删除 targets 指定的版式，并回收其专属媒体。

    targets: 版式部件路径列表（如 ppt/slideLayouts/slideLayout17.xml）；
             None 表示删除所有"未被页面引用"的版式。
    redirect_refs: 若某页引用了将被删除的版式，把引用改指向同结构保留者
                   （res["dup_groups"] 里 keep）；没有保留者则不删该版式。
    """
    z = zipfile.ZipFile(src)
    names = list(z.namelist())
    nameset = set(names)

    # ---- 确定删除清单
    if targets is None:
        targets = [u["part"] for u in res["unused"]]
    targets = [t for t in targets if t in nameset]

    # 有页面引用的不能裸删：要么重定向，要么剔除
    refs_of = {int(k): v for k, v in res["slide_refs"].items()}
    ref_slides = defaultdict(list)
    for slide_no, lay in refs_of.items():
        if lay:
            ref_slides[lay].append(slide_no)

    keep_for = {}
    for g in res["dup_groups"]:
        for m in g["removable"]:
            keep_for[m] = g["keep"]

    final_targets, refused = [], []
    for t in targets:
        users = ref_slides.get(t, [])
        if not users:
            final_targets.append(t)
        elif redirect_refs and t in keep_for and keep_for[t] not in targets:
            final_targets.append(t)                 # 允许删，但要改引用
        else:
            refused.append({"part": t, "ref_slides": users,
                            "reason": "被页面引用且无同结构保留者"})

    if refused:
        print("[dedupe] 拒绝删除 %d 个（仍被引用且无替代）：%s"
              % (len(refused), [r["part"] for r in refused]))

    # ---- 需要重写的关系
    def rels_part_for(part):
        return posixpath.join(posixpath.dirname(part), "_rels",
                              posixpath.basename(part) + ".rels")

    def rels_of(part):
        rp = rels_part_for(part)
        if rp not in nameset:
            return {}
        out = {}
        for rel in read_xml(z.read(rp)):
            out[rel.get("Id")] = (resolve(part, rel.get("Target")),
                                  rel.get("Type", "").rsplit("/", 1)[-1])
        return out

    # 与版式相关的三个枢纽
    master_parts = [n for n in nameset
                    if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]
    masters_rels = {mp: rels_of(mp) for mp in master_parts}

    # ---- 收集：master -> 哪些 rId 指向被删版式
    drop_rid = defaultdict(set)          # master part -> {rid}
    for mp, rm in masters_rels.items():
        for rid, (tgt, typ) in rm.items():
            if typ == "slideLayout" and tgt in final_targets:
                drop_rid[mp].add(rid)

    # ---- 目标媒体：仅被将删版式引用
    media_users = defaultdict(set)
    for n in names:
        if not n.endswith(".xml") or "/_rels/" in n:
            continue
        try:
            root = read_xml(z.read(n))
        except Exception:
            continue
        rm = rels_of(n)
        for blip in root.iter(q("a:blip")):
            rid = blip.get("{%s}embed" % REL_NS) or blip.get("{%s}link" % REL_NS)
            if rid and rid in rm:
                tgt = rm[rid][0]
                if tgt.startswith("ppt/media/"):
                    media_users[tgt].add(n)

    drop_media_set = set()
    if drop_media:
        for m, users in media_users.items():
            if users and users <= set(final_targets):
                drop_media_set.add(m)

    target_rels = set()
    for t in final_targets:
        rp = rels_part_for(t)
        if rp in nameset:
            target_rels.add(rp)

    drop_all = set(final_targets) | target_rels | drop_media_set

    # ---- 组装输出
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".building.pptx")
    report = {
        "removed_layouts": [],
        "removed_media": sorted(drop_media_set),
        "refused": refused,
        "freed_layout_bytes": 0,
        "freed_media_bytes": 0,
    }

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            if n in drop_all:
                continue
            blob = z.read(n)

            # master xml：摘掉 sldLayoutId
            if n in drop_rid:
                root = read_xml(blob)
                removed = []
                for e in list(root.iter(q("p:sldLayoutId"))):
                    rid = e.get("{%s}id" % REL_NS)
                    if rid in drop_rid[n]:
                        removed.append((rid, e.get("id")))
                        e.getparent().remove(e)
                cl = root.find("." + q("p:sldLayoutIdLst"))
                if cl is not None and len(cl) == 0:
                    cl.getparent().remove(cl)
                blob = to_bytes(root)
                if removed:
                    print("[dedupe] %s 摘除 %d 条 sldLayoutId" % (
                        posixpath.basename(n), len(removed)))

            # master rels：摘掉 Relationship
            elif any(n == rels_part_for(mp) for mp in drop_rid):
                host = n[:-5].replace("/_rels/", "/")
                rids = drop_rid.get(host, set())
                root = read_xml(blob)
                cnt = 0
                for rel in list(root):
                    if rel.get("Id") in rids:
                        root.remove(rel)
                        cnt += 1
                blob = to_bytes(root)
                if cnt:
                    print("[dedupe] %s 摘除 %d 条 Relationship" % (
                        posixpath.basename(n), cnt))

            # 页面 rels：把指向被删版式的引用改到保留者
            elif n.endswith(".rels") and "/slides/_rels/" in n:
                host = n[:-5].replace("/_rels/", "/")
                root = read_xml(blob)
                hit = 0
                for rel in root:
                    tgt = resolve(host, rel.get("Target"))
                    if tgt in final_targets:
                        newt = keep_for.get(tgt)
                        if newt:
                            rel.set("Target", relativize(host, newt))
                            hit += 1
                if hit:
                    blob = to_bytes(root)
                    print("[dedupe] %s 重定向 %d 条版式引用" % (
                        posixpath.basename(n), hit))

            # Content_Types：摘掉被删部件的 Override
            elif n == "[Content_Types].xml":
                root = read_xml(blob)
                cnt = 0
                for ov in list(root):
                    pn = (ov.get("PartName") or "").lstrip("/")
                    if pn in drop_all:
                        root.remove(ov)
                        cnt += 1
                blob = to_bytes(root)
                print("[dedupe] [Content_Types].xml 摘除 %d 条 Override" % cnt)

            zo.writestr(n, blob)

    tmp.replace(out)

    # ---- 统计
    for t in final_targets:
        report["removed_layouts"].append({
            "part": t,
            "xml_bytes": z.getinfo(t).file_size,
            "rels_bytes": (z.getinfo(rels_part_for(t)).file_size
                           if rels_part_for(t) in nameset else 0),
        })
    report["freed_layout_bytes"] = sum(
        x["xml_bytes"] + x["rels_bytes"] for x in report["removed_layouts"])
    report["freed_media_bytes"] = sum(
        z.getinfo(m).file_size for m in drop_media_set if m in nameset)

    print()
    print("[dedupe] 删除版式 %d 个，回收 XML %.1f KB"
          % (len(final_targets), report["freed_layout_bytes"] / 1024))
    print("[dedupe] 回收媒体 %d 个，%.1f KB"
          % (len(drop_media_set), report["freed_media_bytes"] / 1024))
    print("[dedupe] 输出 %.2f MB（原 %.2f MB），净减 %.2f MB"
          % (out.stat().st_size / 1024 / 1024, src.stat().st_size / 1024 / 1024,
             (src.stat().st_size - out.stat().st_size) / 1024 / 1024))
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--out", help="输出目录（写 analysis.json / analysis.md）")
    ap.add_argument("--json", action="store_true", help="只打印 JSON")
    ap.add_argument("--apply", action="store_true", help="执行去重（需配合 --out-deck）")
    ap.add_argument("--out-deck", help="去重后输出 pptx 路径")
    ap.add_argument("--only", help="只删这些 idx（逗号分隔）；默认删全部未引用版式")
    ap.add_argument("--keep-media", action="store_true", help="不回收媒体")
    args = ap.parse_args()

    src = Path(args.deck)
    if not src.exists():
        print("[dedupe] 文件不存在:", src)
        return 2

    deck = load_deck(src)
    res = analyze(deck)

    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        md = report_md(res, src)
        print(md)
        if args.out:
            od = Path(args.out)
            od.mkdir(parents=True, exist_ok=True)
            (od / "layout_analysis.json").write_text(
                json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
            (od / "layout_analysis.md").write_text(md, encoding="utf-8")
            print("[dedupe] 已写:", od / "layout_analysis.md")

    if args.apply:
        if not args.out_deck:
            print("[dedupe] --apply 需要 --out-deck")
            return 2
        targets = None
        if args.only:
            want = {int(x.strip()) for x in args.only.split(",") if x.strip()}
            targets = [it["part"] for it in res["layouts"] if it["index"] in want]
            miss = want - {it["index"] for it in res["layouts"]}
            if miss:
                print("[dedupe] 忽略不存在的 idx:", sorted(miss))
        rep = apply_dedupe(src, res, Path(args.out_deck),
                           targets=targets,
                           drop_media=not args.keep_media)
        if args.out:
            od = Path(args.out)
            od.mkdir(parents=True, exist_ok=True)
            (od / "dedupe_report.json").write_text(
                json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
