# -*- coding: utf-8 -*-
"""
analyze_template.py — 解析 .thmx / .pptx 模板包，产出结构清单

用途：
    接手一份新模板（.thmx 或 .pptx）时，先跑这个脚本，拿到：
      1) 画布尺寸
      2) 主题配色方案（clrScheme 全部 12 个槽位）+ 字体方案
      3) slideMaster 数量、slideLayout 数量与各自 name/type/占位符
      4) 品牌色硬编码分布（在 layout/master 里出现了多少次、在哪些文件）
      5) 嵌入字体清单
      6) 媒体文件清单

用法：
    python analyze_template.py "<template.thmx|.pptx>" [--out <dir>] [--brand 0055CD,00DCA5]

产物：
    <out>/template.json   结构化
    <out>/template.md     人读
"""

import argparse
import json
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# Reyee 品牌色
BRAND_DEFAULT = {
    "0055CD": "Reyee Blue (R0 G85 B205)",
    "00DCA5": "Reyee Green (R0 G220 B165)",
}

# 主题配色 12 槽
CLR_SLOTS = [
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]

# 版式类型 -> 中文用途（用于判定「六种版式」）
LAYOUT_TYPE_CN = {
    "title": "封面",
    "obj": "标题+内容",
    "secHead": "章节过渡",
    "twoObj": "标题+内容(双栏)",
    "titleOnly": "仅标题",
    "blank": "空白页",
    "title+4obj": "标题+四象限",
    "picTx": "图片+说明",
    "tx": "纯文本",
    "twoTxTwoObj": "双文本+双对象",
    "objTx": "对象+文本",
    "chart": "图表",
    "tbl": "表格",
    "clipArtTx": "剪贴画+文本",
    "media": "媒体",
    "dgm": "图示",
    "smartArt": "SmartArt",
    "txOverObj": "文本叠对象",
    "objOverTx": "对象叠文本",
    "cust": "自定义",
    "vertTx": "竖排文本",
    "vertTitleAndTx": "竖排标题+文本",
}


def emu_to_px(v):
    try:
        return round(int(v) / 9525.0, 1)
    except Exception:
        return None


def parse_theme(theme_xml: str) -> dict:
    root = ET.fromstring(theme_xml)
    out = {"name": root.get("name")}
    cs = root.find("a:themeElements/a:clrScheme", NS)
    if cs is not None:
        out["clrScheme_name"] = cs.get("name")
        scheme = {}
        for slot in CLR_SLOTS:
            node = cs.find("a:%s" % slot, NS)
            if node is None:
                continue
            srgb = node.find("a:srgbClr", NS)
            sysc = node.find("a:sysClr", NS)
            if srgb is not None:
                scheme[slot] = {"type": "srgb", "val": srgb.get("val")}
            elif sysc is not None:
                scheme[slot] = {
                    "type": "sys", "val": sysc.get("val"),
                    "lastClr": sysc.get("lastClr"),
                }
        out["clrScheme"] = scheme

    fs = root.find("a:themeElements/a:fontScheme", NS)
    if fs is not None:
        out["fontScheme_name"] = fs.get("name")
        fonts = {}
        for which in ("majorFont", "minorFont"):
            f = fs.find("a:%s" % which, NS)
            if f is None:
                continue
            entry = {}
            for tag, key in (("a:latin", "latin"), ("a:ea", "ea"), ("a:cs", "cs")):
                n = f.find(tag, NS)
                if n is not None and n.get("typeface"):
                    entry[key] = n.get("typeface")
            # 东亚字体通常在 <a:font script="Hans" typeface="..."/>
            subs = []
            for n in f.findall("a:font", NS):
                if n.get("typeface"):
                    subs.append({"script": n.get("script"), "typeface": n.get("typeface")})
            if subs:
                entry["script_fonts"] = subs
            fonts[which] = entry
        out["fonts"] = fonts
    return out


def parse_presentation(xml: str, rels_xml: str) -> dict:
    root = ET.fromstring(xml)
    out = {}
    sz = root.find("p:sldSz", NS)
    if sz is not None:
        out["sldSz"] = {
            "cx": int(sz.get("cx")), "cy": int(sz.get("cy")),
            "px": [emu_to_px(sz.get("cx")), emu_to_px(sz.get("cy"))],
            "ratio": round(int(sz.get("cx")) / int(sz.get("cy")), 4),
        }
    ns = root.find("p:notesSz", NS)
    if ns is not None:
        out["notesSz"] = {"cx": int(ns.get("cx")), "cy": int(ns.get("cy"))}
    idlst = root.find("p:sldIdLst", NS)
    out["slide_count"] = len(idlst) if idlst is not None else 0

    rels = ET.fromstring(rels_xml)
    rel_list = []
    for rel in rels.findall("rel:Relationship", NS):
        rtype = rel.get("Type", "").rsplit("/", 1)[-1]
        rel_list.append({"id": rel.get("Id"), "type": rtype, "target": rel.get("Target")})
    out["relationships"] = rel_list
    out["slideMaster_targets"] = [r["target"] for r in rel_list if r["type"] == "slideMaster"]
    out["font_targets"] = [r["target"] for r in rel_list if r["type"] == "font"]
    return out


def parse_master(xml: str) -> dict:
    root = ET.fromstring(xml)
    out = {"txStyles": {}, "placeholders": []}
    tx = root.find("p:txStyles", NS)
    if tx is not None:
        for key in ("titleStyle", "bodyStyle", "otherStyle"):
            node = tx.find("p:%s" % key, NS)
            if node is None:
                continue
            lvls = []
            for lvl in node.findall("a:lvl%dpPr" % 1, NS) + node.findall("a:lvl1pPr", NS):
                pass
            # 收集每个层级的关键属性
            for i in range(1, 10):
                lp = node.find("a:lvl%dpPr" % i, NS)
                if lp is None:
                    continue
                defRPr = lp.find("a:defRPr", NS)
                info = {"lvl": i, "algn": lp.get("algn"), "marL": lp.get("marL")}
                if defRPr is not None:
                    info["sz"] = int(defRPr.get("sz")) / 100.0 if defRPr.get("sz") else None
                    info["b"] = defRPr.get("b")
                    latin = defRPr.find("a:latin", NS)
                    if latin is not None:
                        info["latin"] = latin.get("typeface")
                    ea = defRPr.find("a:ea", NS)
                    if ea is not None:
                        info["ea"] = ea.get("typeface")
                    fill = defRPr.find("a:solidFill/a:srgbClr", NS)
                    if fill is not None:
                        info["srgbClr"] = fill.get("val")
                    sfill = defRPr.find("a:solidFill/a:schemeClr", NS)
                    if sfill is not None:
                        info["schemeClr"] = sfill.get("val")
                lvls.append(info)
            out["txStyles"][key] = lvls

    # 母版占位符
    spTree = root.find("p:cSld/p:spTree", NS)
    if spTree is not None:
        for sp in spTree.findall("p:sp", NS):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            if ph is None:
                continue
            xfrm = sp.find("p:spPr/a:xfrm", NS)
            off = xfrm.find("a:off", NS) if xfrm is not None else None
            ext = xfrm.find("a:ext", NS) if xfrm is not None else None
            out["placeholders"].append({
                "type": ph.get("type") or "body",
                "idx": ph.get("idx"),
                "off_px": [emu_to_px(off.get("x")), emu_to_px(off.get("y"))] if off is not None else None,
                "ext_px": [emu_to_px(ext.get("cx")), emu_to_px(ext.get("cy"))] if ext is not None else None,
            })
    return out


def parse_layout(xml: str) -> dict:
    root = ET.fromstring(xml)
    out = {"placeholders": [], "bg": None}
    cSld = root.find("p:cSld", NS)
    if cSld is not None:
        out["layout_name"] = cSld.get("name")
        bg = cSld.find("p:bg", NS)
        if bg is not None:
            srgb = bg.find(".//a:srgbClr", NS)
            scheme = bg.find(".//a:schemeClr", NS)
            if srgb is not None:
                out["bg"] = {"type": "srgb", "val": srgb.get("val")}
            elif scheme is not None:
                out["bg"] = {"type": "scheme", "val": scheme.get("val")}

    out["layout_type"] = root.get("type") or "cust"
    out["layout_type_cn"] = LAYOUT_TYPE_CN.get(out["layout_type"], out["layout_type"])
    out["preserve"] = root.get("preserve")

    spTree = root.find("p:cSld/p:spTree", NS)
    if spTree is not None:
        for sp in spTree.findall("p:sp", NS):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            xfrm = sp.find("p:spPr/a:xfrm", NS)
            off = xfrm.find("a:off", NS) if xfrm is not None else None
            ext = xfrm.find("a:ext", NS) if xfrm is not None else None
            rec = {
                "name": (sp.find("p:nvSpPr/p:cNvPr", NS).get("name")
                         if sp.find("p:nvSpPr/p:cNvPr", NS) is not None else ""),
                "is_ph": ph is not None,
                "ph_type": (ph.get("type") or "body") if ph is not None else None,
                "ph_idx": ph.get("idx") if ph is not None else None,
                "off_px": [emu_to_px(off.get("x")), emu_to_px(off.get("y"))] if off is not None else None,
                "ext_px": [emu_to_px(ext.get("cx")), emu_to_px(ext.get("cy"))] if ext is not None else None,
            }
            if ph is not None:
                out["placeholders"].append(rec)
        # 图片类形状
        pics = spTree.findall("p:pic", NS)
        if pics:
            out["picture_count"] = len(pics)
    return out


def brand_hits(blob: str, brand: dict) -> dict:
    """统计品牌色在文本块中的出现次数（大小写不敏感，兼容 # 前缀与 8 位 ARGB）"""
    hits = {}
    low = blob.lower()
    for hexv, label in brand.items():
        h = hexv.lower()
        n = low.count(h)
        if n:
            hits[hexv] = {"label": label, "count": n}
    return hits


def analyze(path: Path, brand: dict) -> dict:
    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {
        "file": str(path),
        "file_name": path.name,
        "size_bytes": path.stat().st_size,
        "package_kind": "thmx" if path.suffix.lower() == ".thmx" else "pptx",
        "parts": {},
    }

    # 定位根目录（thmx 里在 theme/ 下，pptx 里在 ppt/ 下）
    if "theme/presentation.xml" in names:
        base = "theme/"
    elif "ppt/presentation.xml" in names:
        base = "ppt/"
    else:
        raise SystemExit("未找到 presentation.xml，可能不是有效的 Office 主题/演示包")

    data["base_dir"] = base

    # presentation
    pres_xml = z.read(base + "presentation.xml").decode("utf-8")
    rels_xml = z.read(base + "_rels/presentation.xml.rels").decode("utf-8")
    data["presentation"] = parse_presentation(pres_xml, rels_xml)

    # theme
    theme_paths = [n for n in names if re.match(re.escape(base) + r"theme/theme\d+\.xml$", n)]
    data["themes"] = []
    for tp in sorted(theme_paths):
        t = parse_theme(z.read(tp).decode("utf-8"))
        t["part"] = tp
        data["themes"].append(t)

    # master
    data["masters"] = []
    for mp in sorted(n for n in names if re.match(re.escape(base) + r"slideMasters/slideMaster\d+\.xml$", n)):
        m = parse_master(z.read(mp).decode("utf-8"))
        m["part"] = mp
        mrels = mp.replace("slideMasters/", "slideMasters/_rels/") + ".rels"
        if mrels in names:
            rx = ET.fromstring(z.read(mrels).decode("utf-8"))
            m["rels"] = [
                {"id": r.get("Id"), "type": r.get("Type", "").rsplit("/", 1)[-1], "target": r.get("Target")}
                for r in rx.findall("rel:Relationship", NS)
            ]
        data["masters"].append(m)

    # layouts
    data["layouts"] = []
    for lp in sorted(
        (n for n in names if re.match(re.escape(base) + r"slideLayouts/slideLayout\d+\.xml$", n)),
        key=lambda s: int(re.search(r"slideLayout(\d+)", s).group(1)),
    ):
        L = parse_layout(z.read(lp).decode("utf-8"))
        L["part"] = lp
        L["index"] = int(re.search(r"slideLayout(\d+)", lp).group(1))
        lrels = lp.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels"
        if lrels in names:
            rx = ET.fromstring(z.read(lrels).decode("utf-8"))
            L["rels"] = [
                {"id": r.get("Id"), "type": r.get("Type", "").rsplit("/", 1)[-1], "target": r.get("Target")}
                for r in rx.findall("rel:Relationship", NS)
            ]
        data["layouts"].append(L)

    # 嵌入字体
    data["embedded_fonts"] = sorted(n for n in names if "/fonts/" in n and n.endswith(".fntdata"))
    # 媒体
    media = sorted(n for n in names if "/media/" in n)
    data["media"] = [{"part": n, "size": z.getinfo(n).file_size} for n in media]

    # 品牌色分布
    scan_parts = [n for n in names if n.endswith(".xml")]
    dist = defaultdict(dict)
    for n in scan_parts:
        try:
            blob = z.read(n).decode("utf-8", errors="ignore")
        except Exception:
            continue
        h = brand_hits(blob, brand)
        if h:
            dist[n] = h
    data["brand_color_distribution"] = dict(dist)

    # 主题配色里是否本来就是品牌色
    theme_brand = {}
    for t in data["themes"]:
        for slot, v in (t.get("clrScheme") or {}).items():
            if v.get("val") and v["val"].upper() in {k.upper() for k in brand}:
                theme_brand[slot] = v["val"]
    data["brand_in_clrScheme"] = theme_brand

    return data


def to_md(d: dict) -> str:
    L = []
    L.append("# 模板结构分析：%s\n" % d["file_name"])
    L.append("- 包类型：`%s`　大小：%.2f MB　基准目录：`%s`" % (
        d["package_kind"], d["size_bytes"] / 1024 / 1024, d["base_dir"]))
    p = d["presentation"]
    if p.get("sldSz"):
        s = p["sldSz"]
        L.append("- 画布：%d × %d EMU = %s px，比例 %.4f" % (s["cx"], s["cy"], s["px"], s["ratio"]))
    L.append("- 含幻灯片：%d 张（thmx 应为 0）" % p.get("slide_count", 0))
    L.append("")

    L.append("## 主题\n")
    for t in d["themes"]:
        L.append("**%s**　主题名 `%s`　配色方案名 `%s`　字体方案 `%s`" % (
            t["part"], t.get("name"), t.get("clrScheme_name"), t.get("fontScheme_name")))
        L.append("")
        L.append("| 槽位 | 值 | 类型 |")
        L.append("|---|---|---|")
        for slot in CLR_SLOTS:
            v = (t.get("clrScheme") or {}).get(slot)
            if v:
                L.append("| %s | `%s` | %s |" % (slot, v.get("val"), v["type"]))
        L.append("")
        f = t.get("fonts") or {}
        for which in ("majorFont", "minorFont"):
            e = f.get(which) or {}
            L.append("- %s：latin=`%s`　ea=`%s`" % (which, e.get("latin", "-"), e.get("ea", "-")))
            for sf in e.get("script_fonts", [])[:8]:
                L.append("  - script=%s → `%s`" % (sf["script"], sf["typeface"]))
        L.append("")

    L.append("## 母版\n")
    for m in d["masters"]:
        L.append("**%s**" % m["part"])
        for key in ("titleStyle", "bodyStyle", "otherStyle"):
            lvls = m["txStyles"].get(key) or []
            if lvls:
                L.append("- `%s`：" % key)
                for lv in lvls[:4]:
                    L.append("  - lvl%d　sz=%s　latin=%s　color=%s%s" % (
                        lv["lvl"], lv.get("sz"), lv.get("latin", "-"),
                        lv.get("srgbClr", ""), lv.get("schemeClr", "") and ("(scheme:%s)" % lv.get("schemeClr")) if lv.get("schemeClr") else ""))
        L.append("- 母版占位符：")
        for ph in m["placeholders"]:
            L.append("  - %s idx=%s　pos=%s　size=%s" % (ph["type"], ph["idx"], ph["off_px"], ph["ext_px"]))
        L.append("")

    L.append("## 版式（%d 个）\n" % len(d["layouts"]))
    L.append("| # | type | 用途(中文) | 名称 | 占位符 | 背景 | 图片 |")
    L.append("|---|---|---|---|---|---|---|")
    for L2 in d["layouts"]:
        phs = ", ".join(sorted({p["ph_type"] for p in L2["placeholders"] if p["is_ph"]})) or "-"
        bg = L2.get("bg")
        bg_s = ("`%s`(%s)" % (bg["val"], bg["type"])) if bg else "-"
        L.append("| %d | `%s` | %s | %s | %s | %s | %s |" % (
            L2["index"], L2["layout_type"], L2["layout_type_cn"],
            L2.get("layout_name", "-"), phs, bg_s, L2.get("picture_count", 0)))
    L.append("")
    L.append("### 版式占位符明细\n")
    for L2 in d["layouts"]:
        L.append("**layout%d（%s｜%s）**" % (L2["index"], L2["layout_type"], L2["layout_type_cn"]))
        if not L2["placeholders"]:
            L.append("- 无占位符")
        for ph in L2["placeholders"]:
            if ph["is_ph"]:
                L.append("- `%s` idx=%s　pos=%s　size=%s" % (ph["ph_type"], ph["ph_idx"], ph["off_px"], ph["ext_px"]))
            else:
                L.append("- (装饰形状) %s　pos=%s　size=%s" % (ph["name"], ph["off_px"], ph["ext_px"]))
        L.append("")

    L.append("## 品牌色分布\n")
    if d["brand_color_distribution"]:
        for part, hits in sorted(d["brand_color_distribution"].items()):
            L.append("- `%s`：" % part)
            for hexv, h in hits.items():
                L.append("  - %s → %s × %d" % (hexv, h["label"], h["count"]))
    else:
        L.append("**包内未发现任何品牌色硬编码。**")
    L.append("")
    if d["brand_in_clrScheme"]:
        L.append("品牌色已存在于 clrScheme 中：%s" % d["brand_in_clrScheme"])
    else:
        L.append("**品牌色未出现在 clrScheme 中** → 主题配色是 Office 默认，品牌色（若存在）为形状级硬编码。")
    L.append("")

    L.append("## 资源\n")
    L.append("- 嵌入字体：%d 个" % len(d["embedded_fonts"]))
    for f in d["embedded_fonts"]:
        L.append("  - `%s`" % f)
    L.append("- 媒体文件：%d 个（合计 %.2f MB）" % (
        len(d["media"]), sum(m["size"] for m in d["media"]) / 1024 / 1024))
    L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("--out", default=None)
    ap.add_argument("--brand", default=None, help="额外品牌色，逗号分隔 hex，如 0055CD,00DCA5")
    args = ap.parse_args()

    src = Path(args.template)
    if not src.exists():
        raise SystemExit("文件不存在：%s" % src)

    brand = dict(BRAND_DEFAULT)
    if args.brand:
        for h in args.brand.split(","):
            h = h.strip().lstrip("#").upper()
            if h:
                brand.setdefault(h, "自定义品牌色 %s" % h)

    out = Path(args.out) if args.out else src.parent / (src.stem + "_analysis")
    out.mkdir(parents=True, exist_ok=True)

    data = analyze(src, brand)
    data["brand_palette_used"] = brand
    (out / "template.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / "template.md").write_text(to_md(data), encoding="utf-8")

    print("文件：%s" % data["file_name"])
    print("包类型：%s　基准目录：%s" % (data["package_kind"], data["base_dir"]))
    p = data["presentation"]
    if p.get("sldSz"):
        print("画布：%s px（%.4f）" % (p["sldSz"]["px"], p["sldSz"]["ratio"]))
    print("主题：%d　母版：%d　版式：%d" % (len(data["themes"]), len(data["masters"]), len(data["layouts"])))
    for t in data["themes"]:
        print("  配色方案名：%s　字体方案：%s" % (t.get("clrScheme_name"), t.get("fontScheme_name")))
        f = t.get("fonts") or {}
        print("  major latin=%s ea=%s | minor latin=%s ea=%s" % (
            (f.get("majorFont") or {}).get("latin"), (f.get("majorFont") or {}).get("ea"),
            (f.get("minorFont") or {}).get("latin"), (f.get("minorFont") or {}).get("ea")))
    print("嵌入字体：%d　媒体：%d" % (len(data["embedded_fonts"]), len(data["media"])))
    if data["brand_color_distribution"]:
        print("品牌色命中部件：")
        for part, hits in sorted(data["brand_color_distribution"].items()):
            print("  %s → %s" % (part, {k: v["count"] for k, v in hits.items()}))
    else:
        print("品牌色命中：无")
    print("产物：%s" % out)


if __name__ == "__main__":
    main()
