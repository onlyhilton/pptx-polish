# -*- coding: utf-8 -*-
"""verify_template_fidelity.py — 模板保真体检（**只读**）

解决什么问题
------------
pptx-polish 的其它工具都在"改 PPT"，但**没有任何一步检查"模板本身有没有被改"**。
后果（2026-09-18 实测，Public Wi-Fi Hotspot 那个包）：

    .thmx 里 Section Title 版式的 `01` 占位符是
        <a:defRPr sz="13800" i="1"> + latin "Noto Sans bold"     ← 斜体
    产物里变成
        <a:defRPr sz="13800" b="0" i="0"> + latin "Noto Sans bold"  ← 斜体被取消
    同时 `Section title` 那行从 b="1" 变成 b="0" i="0"、
    字体从 "Noto Sans" 换成 "Noto Sans Black"、<p:ph sz="quarter"> 被删、
    主题从 Reyee 系换成 "ChatGPT"、16 个版式只剩 8 个。

verify_pptx.py 全过、P0=0 —— 因为它拿**已经改过的版式**当权威值去比对页面，
「模板被改」这个维度它根本不查。本脚本补的就是这一维。

判定口径
--------
以模板（.thmx / 模板 .pptx）为**唯一事实源**，逐版式、逐形状、逐 run 属性比对：

| 级别 | 触发条件 | 含义 |
|---|---|---|
| **P0** | 形状增/删；占位符 `p:ph` 属性变化；形状几何变化；run 数量变化；run 属性**有值→另一个不同值**（sz/b/i/u/spc/kern/cap/strike/baseline、latin/ea/cs 字体族、颜色）；背景图变化；主题/字体方案不是模板的；版式在模板里找不到对应 | **模板被篡改**，必须还原 |
| **P1** | run 属性 `无→"0"/"none"` 这类中性补写；版式数量不符；模板里有、本包没有的版式 | 记录，交人工判断（多为 PowerPoint 重存噪声） |

用法
----
    python verify_template_fidelity.py "<deck.pptx>"                    # 默认浅色模板
    python verify_template_fidelity.py "<deck.pptx>" --theme dark
    python verify_template_fidelity.py "<deck.pptx>" --template "<某个 .thmx/.pptx>"
    python verify_template_fidelity.py "<deck.pptx>" --json out.json

退出码：P0 > 0 → 1；否则 0。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import zipfile

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}

# 注意：本机路径一律走 os.path（Windows 是反斜杠，posixpath.dirname 会返回空串）；
# posixpath 只用于 zip 包内部路径。
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(SKILL_DIR, "assets", "master")

THEMES = {
    "default": "2025 Ruijie Reyee PPT Template-20250530.thmx",
    "light": "2025 Ruijie Reyee PPT Template-20250530.thmx",
    "dark": "2025 Ruijie Reyee PPT Template-Dark-20250620.thmx",
}

# 要逐值比对的 run 属性（含"无 → 有"这种补写也要看到）
RUN_ATTRS = ("sz", "b", "i", "u", "strike", "cap", "spc", "kern", "baseline")
RUN_FONTS = ("latin", "ea", "cs")


# ----------------------------------------------------------------------------- 基础

def resolve_template(name: str) -> str:
    if os.path.isabs(name) and os.path.exists(name):
        return name
    alias = THEMES.get(name.lower())
    cand = [alias] if alias else []
    cand.insert(0, name)
    for c in cand:
        if not c:
            continue
        p = os.path.join(TEMPLATE_DIR, c)
        if os.path.exists(p):
            return p
    raise SystemExit("找不到模板：%s（可用别名：%s）" % (name, ", ".join(sorted(THEMES))))


def detect_base(names) -> str:
    if any(n.startswith("theme/slideMasters/") for n in names):
        return "theme"
    if any(n.startswith("ppt/slideMasters/") for n in names):
        return "ppt"
    raise SystemExit("包里找不到 slideMasters，不是可用的 pptx/thmx")


def rels_of(base: str, part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def read_rels(z, base, part) -> dict:
    rp = rels_of(base, part)
    try:
        x = z.read(rp)
    except KeyError:
        return {}
    root = etree.fromstring(x)
    out = {}
    for rel in root:
        out[rel.get("Id")] = rel.get("Target")
    return out


def resolve(host_part: str, target: str):
    if not target or "://" in target:
        return None
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(host_part), target))


def sha1(z, part):
    try:
        return hashlib.sha1(z.read(part)).hexdigest()[:12]
    except KeyError:
        return "<缺失>"


def fill_of(el) -> str:
    """取 rPr 里的颜色（srgbClr / schemeClr / prstClr），返回 'kind:val' 或 ''"""
    for tag in ("srgbClr", "schemeClr", "prstClr", "sysClr"):
        c = el.find("{%s}%s" % (A, tag))
        if c is not None:
            return "%s:%s" % (tag, c.get("val"))
    return ""


def run_sig(el) -> dict:
    d = {}
    for at in RUN_ATTRS:
        v = el.get(at)
        if v is not None:
            d[at] = v
    for f in RUN_FONTS:
        e = el.find("{%s}%s" % (A, f))
        if e is not None:
            d[f] = e.get("typeface")
    c = fill_of(el)
    if c:
        d["color"] = c
    return d


def _local(el):
    return etree.QName(el).localname


def geom_of_xfrm(xfrm) -> tuple:
    """xfrm -> (x,y,cx,cy)；没有 xfrm 返回 None（=继承）"""
    if xfrm is None:
        return None
    off = xfrm.find("{%s}off" % A)
    ext = xfrm.find("{%s}ext" % A)
    if off is None or ext is None:
        return None
    return (int(off.get("x")), int(off.get("y")),
            int(ext.get("cx")), int(ext.get("cy")))


def canonical_layout(z, base, part) -> dict:
    """把一个版式抽成可比的规范结构"""
    root = etree.fromstring(z.read(part))
    rels = read_rels(z, base, part)

    csld = root.find("{%s}cSld" % P)
    name = (csld.get("name") or "") if csld is not None else ""

    # 背景图（模板的底色全靠它）
    bg_sha = None
    bg = root.find(".//{%s}bg" % P)
    if bg is not None:
        blip = bg.find(".//{%s}blip" % A)
        if blip is not None:
            rid = blip.get("{%s}embed" % R)
            tgt = resolve(part, rels.get(rid, "")) if rid else None
            bg_sha = sha1(z, tgt) if tgt else None

    shapes = []
    spTree = root.find("{%s}cSld/{%s}spTree" % (P, P))
    for el in (spTree if spTree is not None else []):
        kind = _local(el)
        if kind not in ("sp", "pic", "graphicFrame", "grpSp"):
            continue
        cnv = el.find(".//{%s}cNvPr" % P)
        ph_el = el.find(".//{%s}ph" % P)
        ph = None
        if ph_el is not None:
            ph = {"type": ph_el.get("type") or "body",
                  "idx": ph_el.get("idx") or "0",
                  "sz": ph_el.get("sz")}
        xfrm = el.find(".//{%s}xfrm" % A)
        runs = [run_sig(e) for e in el.iter()
                if _local(e) in ("defRPr", "rPr", "endParaRPr")]
        shapes.append({
            "kind": kind,
            "name": (cnv.get("name") or "") if cnv is not None else "",
            "ph": ph,
            "geom": geom_of_xfrm(xfrm),
            "runs": runs,
        })

    return {"name": name, "part": part, "bg": bg_sha, "shapes": shapes}


def canonical_master(z, base, part) -> dict:
    root = etree.fromstring(z.read(part))
    csld = root.find("{%s}cSld" % P)
    shapes = []
    if csld is not None:
        spTree = csld.find("{%s}spTree" % P)
        for el in (spTree if spTree is not None else []):
            if _local(el) not in ("sp", "pic", "graphicFrame", "grpSp"):
                continue
            cnv = el.find(".//{%s}cNvPr" % P)
            shapes.append((cnv.get("name") or "") if cnv is not None else "")
    styles = {}
    for tag in ("titleStyle", "bodyStyle", "otherStyle"):
        el = root.find(".//{%s}%s" % (P, tag))
        if el is not None:
            styles[tag] = [run_sig(e) for e in el.iter() if _local(e) == "defRPr"]
    return {"name": (csld.get("name") or "") if csld is not None else "", "part": part,
            "shapes": shapes, "txStyles": styles}


def collect(z) -> dict:
    names = z.namelist()
    base = detect_base(names)
    layouts, masters = [], []
    # ⚠ 部件名可能带注入前缀（inject_template.py 默认 `tpl`），所以**不能**只匹配裸名
    #   `slideMaster\d+.xml` —— 那样注入后的产物会被数出"0 个版式、0 个母版"，
    #   进而报假 P0「模板被改 / 版式被删」。2026-09-18 实测踩过。
    ms_re = re.compile(re.escape(base) + r"/slideMasters/\w*[Ss]lideMaster\d+\.xml$")
    for m in sorted([n for n in names if ms_re.match(n)]):
        masters.append(canonical_master(z, base, m))
        rels = read_rels(z, base, m)
        mx = z.read(m).decode("utf-8", "ignore")
        for _, rid in re.findall(r'<p:sldLayoutId id="(\d+)" r:id="([^"]+)"', mx):
            t = resolve(m, rels.get(rid, ""))
            if t and t in names:
                layouts.append(canonical_layout(z, base, t))
    # 主题必须取「演示文稿级主题」（presentation.xml.rels 指向的那份），
    # **不能**取"第一个 themeN.xml"。换过模板的包里常同时躺着旧主题
    # （挂在 notesMaster 上、或还没回收的旧母版上），按顺序取会读到旧的那份，
    # 于是把"主题没换干净"误判成"主题被换掉"（方向反了）。
    theme_part = None
    pres = base + "/presentation.xml"
    if pres in names:
        rels = read_rels(z, base, pres)
        for rid, tgt in rels.items():
            # 按**路径形状**认主题，不按文件名前缀 —— 注入后的主题叫 `tplTheme1.xml`，
            # 用 startswith("theme") 会漏掉它，转而读到包里那份旧主题，结论方向就反了。
            if tgt and re.search(r"(^|/)theme/[^/]*[Tt]heme\d+\.xml$", tgt):
                t = resolve(pres, tgt)
                if t and t in names:
                    theme_part = t
                    break
    if theme_part is None:                       # .thmx 没有 presentation.xml
        cands = sorted(n for n in names
                       if re.match(re.escape(base) + r"/theme/\w*[Tt]heme\d+\.xml$", n))
        theme_part = cands[0] if cands else None
    theme = ""
    fontscheme = ""
    if theme_part:
        t = z.read(theme_part).decode("utf-8", "ignore")
        mt = re.search(r'<a:theme[^>]*name="([^"]*)"', t)
        mf = re.search(r'<a:fontScheme name="([^"]*)"', t)
        theme = mt.group(1) if mt else ""
        fontscheme = mf.group(1) if mf else ""
    return {"base": base, "layouts": layouts, "masters": masters,
            "theme": theme, "fontScheme": fontscheme, "theme_part": theme_part,
            "slide_count": len([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)])}


# ----------------------------------------------------------------------------- 比对

def biggest_layout(tpl_layouts, deck_layout):
    """按形状名交集 + 几何交集给模板版式打分（deck 的版式名可能在模板里不存在）"""
    dn = {s["name"] for s in deck_layout["shapes"]}
    dg = {s["geom"] for s in deck_layout["shapes"] if s["geom"]}
    best, bs = None, -1
    for t in tpl_layouts:
        tn = {s["name"] for s in t["shapes"]}
        tg = {s["geom"] for s in t["shapes"] if s["geom"]}
        sc = len(dn & tn) * 10 + len(dg & tg)
        if sc > bs:
            best, bs = t, sc
    return best, bs


def geom_close(a, b) -> bool:
    if a is None or b is None:
        return a is b
    return all(abs(x - y) <= GEOM_TOL for x, y in zip(a, b))


def diff_layout(deck, tpl, findings, ctx):
    def add(sev, msg):
        findings.append({"severity": sev, "where": ctx, "msg": msg})

    if deck["bg"] != tpl["bg"]:
        add("P0", "版式背景图变了：模板 %s → 本包 %s" % (tpl["bg"], deck["bg"]))

    ds = {(s["kind"], s["name"]): s for s in deck["shapes"]}
    ts = {(s["kind"], s["name"]): s for s in tpl["shapes"]}
    for k in ts:
        if k not in ds:
            add("P0", "模板形状被删除：%s「%s」" % (k[0], k[1]))
    for k in ds:
        if k not in ts:
            add("P0", "多出非模板形状：%s「%s」" % (k[0], k[1]))

    for k in sorted(set(ds) & set(ts)):
        d, t = ds[k], ts[k]
        tag = "%s「%s」" % (k[0], k[1])
        if d["ph"] != t["ph"]:
            add("P0", "%s 占位符槽位变了：模板 %s → 本包 %s" % (tag, t["ph"], d["ph"]))
        if d["geom"] != t["geom"]:
            add("P0", "%s 几何变了：模板 %s → 本包 %s" % (tag, t["geom"], d["geom"]))
        if len(d["runs"]) != len(t["runs"]):
            add("P1", "%s run 数不同：模板 %d → 本包 %d" % (tag, len(t["runs"]), len(d["runs"])))
        for i, (dr, tr) in enumerate(zip(d["runs"], t["runs"])):
            for key in sorted(set(dr) | set(tr)):
                a, b = tr.get(key), dr.get(key)
                if a == b:
                    continue
                if a is not None and b is not None:
                    add("P0", "%s run#%d %s：模板 %s → 本包 %s" % (tag, i, key, a, b))
                else:
                    add("P1", "%s run#%d %s 被补写：模板 %s → 本包 %s" % (tag, i, key, a, b))


def main() -> int:
    ap = argparse.ArgumentParser(description="模板保真体检：产物是否改动了模板本体")
    ap.add_argument("deck")
    ap.add_argument("--template", default="default", help="模板路径或别名 default/light/dark")
    ap.add_argument("--json", default=None, help="把结论写成 json")
    args = ap.parse_args()

    tpl_path = resolve_template(args.template)
    zt = zipfile.ZipFile(tpl_path)
    zd = zipfile.ZipFile(args.deck)
    T = collect(zt)
    D = collect(zd)

    # 模板里的"历史导入副本"（1_/2_ 前缀）不参与比对
    tpl_official = [l for l in T["layouts"] if not re.match(r"^[12]_", l["name"])]
    tpl_by_name = {}
    for l in tpl_official:
        tpl_by_name.setdefault(l["name"], l)

    findings = []
    print("=" * 84)
    print("模板保真体检")
    print("  产物 : %s" % args.deck)
    print("  模板 : %s" % tpl_path)
    print("  模板主题 %s / fontScheme=%s ；产物主题 %s / fontScheme=%s"
          % (T["theme"], T["fontScheme"], D["theme"], D["fontScheme"]))
    print("  模板版式 %d（正式 %d） ；产物版式 %d ；产物页数 %d"
          % (len(T["layouts"]), len(tpl_official), len(D["layouts"]), D["slide_count"]))
    print("=" * 84)

    if D["fontScheme"] != T["fontScheme"]:
        findings.append({"severity": "P0", "where": "主题",
                         "msg": "字体方案不是模板的：模板 fontScheme=%s → 本包 %s"
                                "（说明产物根本没挂在模板主题上）" % (T["fontScheme"], D["fontScheme"])})

    for dl in D["layouts"]:
        t = tpl_by_name.get(dl["name"])
        tag = "版式「%s」" % dl["name"]
        if re.match(r"^[12]_", dl["name"]):
            findings.append({"severity": "P1", "where": tag,
                             "msg": "非模板官方版式（1_/2_ 前缀的历史导入副本），"
                                    "不在模板格式范围内 —— 常见来源是多次套模板后残留"})
            continue
        if t is None:
            best, sc = biggest_layout(tpl_official, dl)
            if best is not None and sc >= 30:
                findings.append({"severity": "P0", "where": tag,
                                 "msg": "版式名在模板里不存在，疑似由模板版式「%s」改名而来"
                                        "（形状重合分 %d）" % (best["name"], sc)})
                diff_layout(dl, best, findings, tag)
            else:
                findings.append({"severity": "P1", "where": tag,
                                 "msg": "模板里没有这个版式，且与任何模板版式都不相似 ——"
                                        "属于自建版式或另一世代模板的版式，需人工确认"})
            continue
        diff_layout(dl, t, findings, tag)

    for tl in tpl_official:
        if tl["name"] not in {l["name"] for l in D["layouts"]}:
            findings.append({"severity": "P1", "where": "版式「%s」" % tl["name"],
                             "msg": "模板里有、产物里没有（版式被删或未被引用）"})

    # 母版
    if D["masters"] and T["masters"]:
        dm, tm = D["masters"][0], T["masters"][0]
        if sorted(dm["shapes"]) != sorted(tm["shapes"]):
            findings.append({"severity": "P0", "where": "母版",
                             "msg": "母版形状清单变了：模板 %s → 本包 %s" % (tm["shapes"], dm["shapes"])})
        for tag in sorted(set(dm["txStyles"]) | set(tm["txStyles"])):
            a = tm["txStyles"].get(tag, [])
            b = dm["txStyles"].get(tag, [])
            if a != b:
                findings.append({"severity": "P0", "where": "母版/%s" % tag,
                                 "msg": "母版文字样式变了：模板 %s → 本包 %s" % (a, b)})

    p0 = [f for f in findings if f["severity"] == "P0"]
    p1 = [f for f in findings if f["severity"] == "P1"]
    print("\nP0（模板被改，必须还原）= %d 条" % len(p0))
    for f in p0:
        print("  [P0] %-34s %s" % (f["where"], f["msg"]))
    print("\nP1（需人工判断，多为重存噪声）= %d 条" % len(p1))
    for f in p1[:60]:
        print("  [P1] %-34s %s" % (f["where"], f["msg"]))
    if len(p1) > 60:
        print("  ... 另有 %d 条，见 --json" % (len(p1) - 60))

    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"deck": args.deck, "template": tpl_path,
                       "P0": p0, "P1": p1}, fh, ensure_ascii=False, indent=2)
        print("\n结论已写 %s" % args.json)

    print("\n结论：%s" % ("模板已被改动，产物不是「模板格式」——先还原模板再谈其它"
                        if p0 else "版式/母版与模板一致（P1 项请人工确认）"))
    return 1 if p0 else 0


if __name__ == "__main__":
    sys.exit(main())
