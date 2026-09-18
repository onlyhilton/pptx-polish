# -*- coding: utf-8 -*-
"""配色普查与品牌色替换（Step 4）

两种模式：
  --report                普查：列出所有 srgbClr 色值、出现次数、所在部件与页面
  --map "OLD=NEW,..."     替换：把指定色值换成品牌色，输出新文件

替换范围（**默认只动页面**）：ppt/slides/*.xml + ppt/notesSlides/*.xml
  ⚠️ **版式与母版默认不改**（`ppt/slideLayouts/` `ppt/slideMasters/`）。
  它们是模板本体 —— 改了就不是"套模板"，而是"改模板"，用户一眼能看出来。
  确实需要连模板一起改时显式加 `--include-template`，并先在报告里说清改了哪几个版式。
  （2026-09-18 事故：默认扫版式/母版，导致产物里的模板与 .thmx 不再一致。）

不触碰 a:schemeClr（主题引用）—— 改主题槽位会连带影响所有引用处，风险不可控。

中性色（近黑/近白/灰）默认不在报告主表里，用 --all 一并列出。

用法：
    python recolor_deck.py "<deck.pptx>" --report
    python recolor_deck.py "<deck.pptx>" --map "0B49C1=0055CD,0EBE8E=00DCA5" --out "<new.pptx>"
    python recolor_deck.py "<deck.pptx>" --map "..." --out "..." --include-template   # 动模板，慎用
"""
from __future__ import annotations

import argparse
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
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 品牌色（用户给定，2026-09 确认）
BRAND = {
    "0055CD": "Reyee 蓝（主色）",
    "00DCA5": "Reyee 绿（强调色）",
}

SLIDE_DIRS = ("ppt/slides/", "ppt/notesSlides/")
TEMPLATE_DIRS = ("ppt/slideLayouts/", "ppt/slideMasters/")
# 默认扫描范围 = 只扫页面。模板本体（版式/母版）要显式 --include-template 才动。
SCAN_DIRS = SLIDE_DIRS


def q(tag):
    pre, local = tag.split(":")
    return "{%s}%s" % (NS[pre], local)


def hex_of(el) -> str:
    v = (el.get("val") or "").strip().upper()
    return v if re.fullmatch(r"[0-9A-F]{6}", v) else ""


def is_neutral(h: str, spread: int = 14) -> bool:
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return max(r, g, b) - min(r, g, b) < spread


TEXT_TAGS = {"rPr", "defRPr", "endParaRPr"}


def role_of(el) -> str:
    """判断这个颜色是干什么用的。

    这一步不能省：文字色和装饰填充色可能色值相近但用途完全不同 ——
    把正文的深灰蓝文字统一成品牌亮蓝，可读性会直接崩掉。
    """
    node = el.getparent()
    depth = 0
    while node is not None and depth < 8:
        tag = etree.QName(node).localname
        if tag in TEXT_TAGS:
            # 标题占位符里的文字色**单独成一类**：它的正确值是版式槽位决定的
            # （本模板内容页标题 = `043CC1`），不属于"品牌色统一"的对象 ——
            # 把它统一成品牌主色 `0055CD` 就是把合格改成不合格（2026-09-17 的实测事故）。
            holder = node.getparent()
            while holder is not None and etree.QName(holder).localname != "sp":
                holder = holder.getparent()
            if holder is not None:
                for ph in holder.iter():
                    if etree.QName(ph).localname == "ph" and \
                            (ph.get("type") or "body") in ("title", "ctrTitle"):
                        return "title"
            return "text"
        if tag == "ln":
            return "line"
        if tag == "gs":
            return "gradient"
        if tag in ("outerShdw", "innerShdw", "prstShdw"):
            return "shadow"
        if tag in ("bgPr", "bg"):
            return "bg"
        if tag in ("spPr", "grpSpPr", "txPr", "fill", "solidFill"):
            # 遇到 spPr 就说明是形状属性；但如果是 txPr 则往下不是文字色
            if tag == "spPr":
                return "fill"
        node = node.getparent()
        depth += 1
    return "other"


def scan(z, names, dirs=SCAN_DIRS):
    """返回 {色值: {'count','roles','parts','slides','ctx'}}
    另附 --include-template 时才算的"模板本体命中数"（见 scan_template_hits）。
    """
    stats = defaultdict(lambda: {"count": 0, "roles": defaultdict(int),
                                 "parts": defaultdict(int),
                                 "slides": defaultdict(int), "ctx": defaultdict(int)})
    for n in sorted(names):
        if not n.endswith(".xml"):
            continue
        if not n.startswith(dirs):
            continue
        try:
            root = etree.fromstring(z.read(n))
        except Exception:
            continue
        is_slide = bool(re.match(r"ppt/slides/slide\d+\.xml$", n))
        page = int(re.search(r"slide(\d+)", n).group(1)) if is_slide else None
        for el in root.iter(q("a:srgbClr")):
            h = hex_of(el)
            if not h:
                continue
            s = stats[h]
            s["count"] += 1
            s["roles"][role_of(el)] += 1
            s["parts"][n] += 1
            if page is not None:
                s["slides"][page] += 1
            par = el.getparent()
            s["ctx"][etree.QName(par).localname if par is not None else "?"] += 1
    return stats


def report(stats, show_all=False):
    neutral = {h: s for h, s in stats.items() if is_neutral(h)}
    color = {h: s for h, s in stats.items() if not is_neutral(h)}

    print("=" * 78)
    print("一、有彩色（%d 种，共 %d 处）" % (
        len(color), sum(s["count"] for s in color.values())))
    print("=" * 78)
    print("  %-9s %5s  %-32s %s" % ("色值", "次数", "用途分布", "出现页面 / 品牌"))
    for h, s in sorted(color.items(), key=lambda kv: -kv[1]["count"]):
        pages = sorted(s["slides"])
        ptxt = ("P" + ",P".join(str(x) for x in pages[:8]) +
                ("…" if len(pages) > 8 else "")) if pages else "（仅版式/母版）"
        roles = ",".join("%s×%d" % (k, v) for k, v in
                         sorted(s["roles"].items(), key=lambda kv: -kv[1]))
        brand = BRAND.get(h, "")
        print("  #%-8s %5d  %-32s %s%s" % (
            h, s["count"], roles, ptxt, ("   [%s]" % brand) if brand else ""))

    print()
    print("=" * 78)
    print("二、中性色（%d 种，共 %d 处）%s" % (
        len(neutral), sum(s["count"] for s in neutral.values()),
        "" if show_all else "  —— 加 --all 展开"))
    print("=" * 78)
    if show_all:
        for h, s in sorted(neutral.items(), key=lambda kv: -kv[1]["count"]):
            print("  #%-8s %5d" % (h, s["count"]))
    else:
        top = sorted(neutral.items(), key=lambda kv: -kv[1]["count"])[:6]
        for h, s in top:
            print("  #%-8s %5d" % (h, s["count"]))
        if len(neutral) > 6:
            print("  ... 另有 %d 种" % (len(neutral) - 6))
    return color, neutral


def template_hits(z, names, mapping):
    """统计"若连模板一起改会命中多少处" —— 只统计，不写文件。"""
    hit = defaultdict(int)
    for n in sorted(names):
        if not n.endswith(".xml") or not n.startswith(TEMPLATE_DIRS):
            continue
        try:
            root = etree.fromstring(z.read(n))
        except Exception:
            continue
        for el in root.iter(q("a:srgbClr")):
            if hex_of(el) in mapping:
                hit[posixpath.basename(n)] += 1
    return hit


def do_replace(src: Path, out: Path, mapping: dict, only_roles=None, dirs=SCAN_DIRS):
    """按色值替换。only_roles 给定时只替换这些角色的出现处。

    为什么需要角色过滤：同一色值可能同时用于填充和文字。把浅色背景上的
    深绿文字统一成品牌亮绿，白底对比度会从 4.07:1 掉到 1.78:1（WCAG 需 4.5:1），
    文字直接不可读。所以"色系统一"必须区分"装饰用"与"文字用"。

    dirs 默认只含页面；模板本体（版式/母版）要显式传进来才会被改。
    """
    z = zipfile.ZipFile(src)
    names = list(z.namelist())
    changed_parts = defaultdict(int)
    by_role = defaultdict(int)
    total = 0

    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".building.pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            blob = z.read(n)
            if (n.endswith(".xml") and n.startswith(dirs)
                    and n not in ("[Content_Types].xml",)):
                try:
                    root = etree.fromstring(blob)
                except Exception:
                    zo.writestr(n, blob)
                    continue
                hit = 0
                for el in root.iter(q("a:srgbClr")):
                    h = hex_of(el)
                    if h not in mapping:
                        continue
                    if only_roles:
                        r = role_of(el)
                        if r not in only_roles:
                            continue
                        by_role[r] += 1
                    el.set("val", mapping[h])
                    hit += 1
                if hit:
                    blob = etree.tostring(root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                    changed_parts[n] = hit
                    total += hit
            zo.writestr(n, blob)
    tmp.replace(out)

    print()
    print("[recolor] 替换 %d 处，涉及 %d 个部件%s" % (
        total, len(changed_parts),
        ("，角色分布 " + ", ".join("%s×%d" % kv for kv in
                                 sorted(by_role.items(), key=lambda kv: -kv[1])))
        if by_role else ""))
    for n, c in sorted(changed_parts.items(), key=lambda kv: -kv[1])[:15]:
        print("   %-46s %4d 处" % (posixpath.basename(n), c))
    if len(changed_parts) > 15:
        print("   ... 另有 %d 个部件" % (len(changed_parts) - 15))
    print("[recolor] 输出 %s（%.2f MB）" % (out.name, out.stat().st_size / 1024 / 1024))
    return {"replaced": total, "parts": dict(changed_parts),
            "mapping": mapping, "only_roles": sorted(only_roles) if only_roles else None}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--report", action="store_true")
    ap.add_argument("--all", action="store_true", help="中性色也全部展开")
    ap.add_argument("--map", dest="mapping", help="替换表 OLD=NEW,OLD2=NEW2")
    ap.add_argument("--only-role",
                    help="只替换这些角色（逗号分隔）：text,fill,line,gradient,shadow,bg。"
                         "不给则所有角色都替换")
    ap.add_argument("--out", help="替换后输出路径")
    ap.add_argument("--out-dir", help="报告输出目录")
    ap.add_argument("--include-template", action="store_true",
                    help="连模板本体（ppt/slideLayouts + ppt/slideMasters）一起改。"
                         "默认关闭 —— 改版式/母版=改模板，不再属于「用模板格式改 PPT」")
    args = ap.parse_args()

    src = Path(args.deck)
    z = zipfile.ZipFile(src)
    names = set(z.namelist())
    dirs = SLIDE_DIRS + (TEMPLATE_DIRS if args.include_template else ())
    stats = scan(z, names, dirs)

    color, neutral = report(stats, args.all)

    if args.out_dir:
        od = Path(args.out_dir)
        od.mkdir(parents=True, exist_ok=True)
        data = {
            "file": str(src),
            "brand": BRAND,
            "color": {h: {"count": s["count"], "slides": sorted(s["slides"]),
                          "roles": dict(s["roles"])} for h, s in color.items()},
            "neutral": {h: {"count": s["count"], "roles": dict(s["roles"])}
                        for h, s in neutral.items()},
        }
        (od / "color_survey.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print("\n[recolor] 报告已写:", od / "color_survey.json")

    if args.mapping:
        if not args.out:
            print("[recolor] --map 需要 --out")
            return 2
        mapping = {}
        for pair in args.mapping.split(","):
            if "=" not in pair:
                continue
            a, b = pair.split("=", 1)
            mapping[a.strip().upper()] = b.strip().upper()
        print("\n[recolor] 替换表: %s" % mapping)
        roles = None
        if args.only_role:
            roles = {x.strip() for x in args.only_role.split(",") if x.strip()}
            print("[recolor] 仅限角色: %s" % sorted(roles))
        hit_tpl = template_hits(z, names, mapping)
        if hit_tpl and not args.include_template:
            print("[recolor] ⚠ 模板本体里也有这些色值（默认**不改**，以免动到模板）：")
            for k, v in sorted(hit_tpl.items(), key=lambda kv: -kv[1])[:12]:
                print("             %-34s %4d 处" % (k, v))
            print("           → 若确实要改模板，加 --include-template（并知悉：产物将不再等于模板格式）")
        rep = do_replace(src, Path(args.out), mapping, only_roles=roles, dirs=dirs)
        rep["template_touched"] = bool(args.include_template)
        rep["template_hits_skipped"] = dict(hit_tpl) if not args.include_template else {}
        if args.out_dir:
            od = Path(args.out_dir)
            od.mkdir(parents=True, exist_ok=True)
            (od / "recolor_report.json").write_text(
                json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
