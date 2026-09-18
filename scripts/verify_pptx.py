# -*- coding: utf-8 -*-
"""
verify_pptx.py — PPTX 包完整性体检（交付前自检，不依赖 PowerPoint）

检查项：
    1. zip 结构完整性
    2. [Content_Types].xml 是否覆盖所有部件（尤其 .xml / .rels）
    3. 每条关系（.rels）的 Target 是否真实存在（悬空引用 = 文件损坏）
    4. 关键链路是否闭合：
         slide → slideLayout → slideMaster → theme
         presentation → slideMaster → slideLayout(×N)
    5. presentation.xml 的 sldMasterIdLst 的 r:id 是否在 rels 里
    6. python-pptx 能否正常打开，每页挂的版式名
    7. 每页占位符 vs 版式占位符的承接（孤儿内容预警）
    8. 标题「字体 / 字号 / 颜色」是否回归版式槽位设定（页面级覆盖检出）
    9. python-pptx 打开 + 逐页生效字号/字体清单

用法：
    python verify_pptx.py "file.pptx" [--out <dir>] [--json]
退出码：
    0 = 通过　1 = 有 P0 错误
"""

import argparse
import json
import posixpath
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as _ET

try:
    from lxml import etree as ET          # lxml 支持 getparent()，优先用
except ImportError:                        # 兜底：标准库（无 getparent）
    ET = _ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# OOXML 对以下元素的**子元素顺序**有硬性要求（XSD sequence）。
# 顺序错了 PowerPoint 会直接报"文件已损坏"，而 python-pptx 往往仍能打开 —— 必须单独检。
ELEM_ORDER = {
    "presentation": ["sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "sldIdLst",
                     "sldSz", "notesSz", "smartTags", "embeddedFontLst", "custShowLst",
                     "photoAlbum", "custDataLst", "kinsoku", "defaultTextStyle",
                     "modifyVerifier", "extLst"],
    "sldMaster": ["cSld", "clrMap", "sldLayoutIdLst", "transition", "timing", "hf",
                  "txStyles", "extLst"],
    "sldLayout": ["cSld", "clrMapOvr", "transition", "timing", "hf", "extLst"],
    "sld": ["cSld", "clrMapOvr", "transition", "timing", "extLst"],
    "cSld": ["bg", "spTree", "custDataLst", "controls", "extLst"],
}


def localname(el):
    return el.tag.rsplit("}", 1)[-1]


def check_order(z, names, res):
    targets = [("ppt/presentation.xml", "presentation")]
    for n in names:
        if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n):
            targets.append((n, "sldMaster"))
        elif re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n):
            targets.append((n, "sldLayout"))
        elif re.match(r"ppt/slides/slide\d+\.xml$", n):
            targets.append((n, "sld"))

    bad = []
    for part, kind in targets:
        try:
            root = ET.fromstring(z.read(part))
        except Exception as e:
            bad.append("%s 解析失败(%s)" % (part, e))
            continue
        order = ELEM_ORDER[kind]
        last, last_tag = -1, None
        for child in root:
            tag = localname(child)
            if tag not in order:
                continue
            i = order.index(tag)
            if i < last:
                bad.append("%s: <%s> 出现在 <%s> 之后" % (part, tag, last_tag))
            last, last_tag = i, tag
        cSld = root.find("p:cSld", NS)
        if cSld is not None:
            last, last_tag = -1, None
            for child in cSld:
                tag = localname(child)
                if tag not in ELEM_ORDER["cSld"]:
                    continue
                i = ELEM_ORDER["cSld"].index(tag)
                if i < last:
                    bad.append("%s: <cSld/%s> 出现在 <%s> 之后" % (part, tag, last_tag))
                last, last_tag = i, tag

    if bad:
        res["P0"].append("OOXML 元素顺序错误 %d 处：%s" % (len(bad), "; ".join(bad[:4])))
    else:
        res["info"]["element_order"] = "OK"


def resolve(host_part, target):
    """解析相对 Target 为包内绝对路径。host_part 为空串表示包根。"""
    if target.startswith("/"):
        return target.lstrip("/")
    if not host_part:
        return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join(posixpath.dirname(host_part), target))


def rels_host(rels_path):
    """由 .rels 路径反推宿主部件路径。根 _rels/.rels -> ''"""
    if rels_path == "_rels/.rels":
        return ""
    stem = rels_path[:-5]                       # 去掉 .rels
    host, _, fname = stem.rpartition("/_rels/")
    if not _:
        return ""
    return posixpath.join(host, fname) if host else fname


# ---------------------------------------------------------------------------
# 字体 / 字号 / 颜色 继承链
#
# 背景（2026-09-17，T3T4 ISP Residential Access Network Solution）：
#   原稿把「26pt Arial」写进了**每一页**的标题里（slide 的 <a:lstStyle>/<a:p>/<a:rPr>），
#   而模板「标题和文本」版式槽位写的是 28pt「Noto Sans bold」+ 043CC1。
#   页面级覆盖优先级高于版式，所以 24 页标题全部偏离模板 ——
#   但旧版 verify_pptx.py / audit_deck.py 只查「包结构 + 占位符能否被版式承接 +
#   几何越界」，**从不比对字体族与字号**，一个都不报。
#   本段把这一维度补上：逐页算「生效值」，与版式槽位设定值比对。
#
# OOXML 覆盖优先级（低 → 高）：
#   母版 p:txStyles/titleStyle
#     → 版式槽位 a:lstStyle/a:lvl1pPr/a:defRPr   ← 模板设定，权威值
#       → 页面 a:lstStyle/a:lvl1pPr/a:defRPr
#         → 页面 a:p/a:pPr/a:defRPr
#           → 页面 a:p/a:r/a:rPr                  ← 单行运行实际生效
# 所以「要让文字回归模板」必须**删掉页面级覆盖**，而不是在版式上再写一遍。
# ---------------------------------------------------------------------------

def _lname(el):
    return el.tag.rsplit("}", 1)[-1]


def _lvl1(host):
    """在 lstStyle / txStyles 子节点里取 lvl1pPr（兼容 lvl="0" 写法）。"""
    if host is None:
        return None
    for c in host:
        if _lname(c) == "lvl1pPr":
            return c
    return None


def _font_sig(rpr):
    """从 defRPr / rPr 抽字号·字体·颜色；缺项为 None。"""
    if rpr is None:
        return {}
    d = {"sz": rpr.get("sz"), "b": rpr.get("b"),
         "latin": None, "ea": None, "fill": None}
    for tag in ("latin", "ea"):
        el = rpr.find("a:%s" % tag, NS)
        if el is not None:
            d[tag] = el.get("typeface")
    sf = rpr.find("a:solidFill", NS)
    if sf is not None and len(sf):
        c0, n0 = sf[0], _lname(sf[0])
        d["fill"] = ("#" + (c0.get("val") or "").upper()) if n0 == "srgbClr" \
            else ("%s:%s" % (n0, c0.get("val")))
    return d


def _over(base, over):
    """over 里非 None 的项覆盖 base —— 复刻 OOXML 的逐级覆盖。"""
    out = dict(base or {})
    for k, v in (over or {}).items():
        if v is not None:
            out[k] = v
    return out


def _title_sp(root):
    """取页面 / 版式里的标题占位符 sp 元素（ph → nvPr → nvSpPr → sp）。"""
    if root is None:
        return None
    for ph in root.iter("{%s}ph" % NS["p"]):
        if (ph.get("type") or "") in ("title", "ctrTitle"):
            return ph.getparent().getparent().getparent()
    return None


def _autofit_scale(sp_el):
    """读页面上的 <a:normAutofit fontScale>，返回 (scale, 说明)。1.0 = 未缩排。"""
    if sp_el is None:
        return 1.0, None
    bp = sp_el.find("p:txBody/a:bodyPr", NS)
    if bp is None:
        return 1.0, None
    for c in bp:
        n = _lname(c)
        if n == "normAutofit":
            fs = c.get("fontScale")
            return (int(fs) / 100000.0, "normAutofit") if fs else (1.0, "normAutofit(空)")
        if n in ("noAutofit", "spAutoFit"):
            return 1.0, n
    return 1.0, None


def title_style_chain(sp_el, lroot, mroot):
    """算标题的生效字号·字体·颜色，并保留「哪一级覆盖了它」。"""
    # 母版 txStyles/titleStyle
    m_sig, m_src = {}, None
    ts = mroot.find("p:txStyles", NS) if mroot is not None else None
    if ts is not None:
        st = ts.find("p:titleStyle", NS)
        lv = _lvl1(st) if st is not None else None
        if lv is not None:
            m_sig = _font_sig(lv.find("a:defRPr", NS))
            if any(v for v in m_sig.values()):
                m_src = "master:titleStyle"
    # 版式槽位（模板设定，权威值）
    l_sig = {}
    lsp = _title_sp(lroot)
    if lsp is not None:
        tb = lsp.find("p:txBody", NS)
        lv = _lvl1(tb.find("a:lstStyle", NS) if tb is not None else None)
        if lv is not None:
            s = _font_sig(lv.find("a:defRPr", NS))
            if any(v for v in s.values()):
                l_sig = s
    # 页面级覆盖
    chain = []
    if sp_el is not None:
        tb = sp_el.find("p:txBody", NS)
        lv = _lvl1(tb.find("a:lstStyle", NS) if tb is not None else None)
        if lv is not None:
            s = _font_sig(lv.find("a:defRPr", NS))
            if any(v for v in s.values()):
                chain.append(("slide:lstStyle", s))
        p_sig, r_sig = {}, {}
        for p in tb.findall("a:p", NS):
            pp = p.find("a:pPr", NS)
            if not p_sig and pp is not None:
                p_sig = _font_sig(pp.find("a:defRPr", NS))
            for r in p.findall("a:r", NS):
                s = _font_sig(r.find("a:rPr", NS))
                if any(v for v in s.values()):
                    r_sig = _over(r_sig, s)
        if any(v for v in p_sig.values()):
            chain.append(("slide:pPr/defRPr", p_sig))
        if any(v for v in r_sig.values()):
            chain.append(("slide:run/rPr", r_sig))

    base = _over(m_sig, l_sig)          # 模板设定
    eff = base
    for _, sig in chain:
        eff = _over(eff, sig)           # 页面级逐步覆盖
    return {"base": base, "effective": eff, "chain": chain,
            "slot": l_sig, "master": m_sig,
            "source": "layout:slot" if l_sig else (m_src or "master:titleStyle")}


def _fmt_sz(v):
    try:
        return "%.1fpt" % (int(v) / 100.0)
    except (TypeError, ValueError):
        return "-"


def check(path: Path):
    res = {"file": str(path), "P0": [], "P1": [], "info": {}, "slides": []}
    z = zipfile.ZipFile(path)
    names = set(z.namelist())

    # 1. zip 完整性
    bad = z.testzip()
    if bad:
        res["P0"].append("zip 损坏：%s" % bad)

    # 2. Content_Types
    ct = ET.fromstring(z.read("[Content_Types].xml"))
    defaults = {d.get("Extension").lower(): d.get("ContentType") for d in ct.findall("ct:Default", NS)}
    overrides = {o.get("PartName"): o.get("ContentType") for o in ct.findall("ct:Override", NS)}
    uncovered = []
    for n in names:
        if n.endswith("/"):
            continue
        if "/" + n in overrides or n in overrides:
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if ext not in defaults:
            uncovered.append(n)
    if uncovered:
        res["P0"].append("Content_Types 未覆盖 %d 个部件：%s" % (len(uncovered), uncovered[:8]))

    # 2b. Content_Types 的**内容类型本身**是否正确（不只是"覆盖到了"）
    #     实测漏网（2026-09-17，T3T4 ISP）：inject_template.py 旧版把模板块的伴生
    #     .rels 判成 slideLayout+xml —— 覆盖率检查通过（有 Override 就算覆盖），
    #     python-pptx 照常打开、本函数报 P0=0，但 **PowerPoint 直接 E_FAIL 拒收**。
    #     .rels 的 content-type 恒为 relationships+xml，xml 部件不能拿泛型类型。
    RELS_CT = "application/vnd.openxmlformats-package.relationships+xml"
    mis_ct = []
    for pn, ctype in overrides.items():
        p = (pn or "").lstrip("/")
        if p.endswith(".rels") and ctype != RELS_CT:
            mis_ct.append("%s => %s（.rels 必须是 relationships+xml）" % (p, ctype))
        elif p.lower().endswith(".xml") and ctype in (
                "application/xml", "text/xml", "application/octet-stream"):
            mis_ct.append("%s => %s（泛型类型）" % (p, ctype))
    if mis_ct:
        res["P0"].append("Content_Types 类型错误 %d 处：%s" % (len(mis_ct), mis_ct[:6]))

    # 2c. `Default xml=application/xml` 是危险兜底：任何 .xml 部件只要没有 Override，
    #     就会静默拿到泛型类型。不是必然崩，但等于"靠兜底活着"，列为 P1。
    if defaults.get("xml") in ("application/xml", "text/xml"):
        generic = [n for n in sorted(names)
                   if n.endswith(".xml") and n != "[Content_Types].xml"
                   and ("/" + n) not in overrides and n not in overrides]
        if generic:
            res["P1"].append("有 %d 个 .xml 部件靠 Default xml=application/xml 兜底（建议补 Override）：%s"
                             % (len(generic), generic[:6]))

    # 3. 所有 rels 的 Target 存在性
    dangling = []
    rels_files = [n for n in names if n.endswith(".rels")]
    for rf in rels_files:
        host = rels_host(rf)
        root = ET.fromstring(z.read(rf))
        for rel in root:
            t = rel.get("Target")
            if not t or rel.get("TargetMode") == "External":
                continue
            abs_t = resolve(host, t)
            if abs_t not in names:
                dangling.append({"rels": rf, "id": rel.get("Id"),
                                 "type": rel.get("Type", "").rsplit("/", 1)[-1],
                                 "target": t, "resolved": abs_t})
    if dangling:
        res["P0"].append("悬空引用 %d 条：%s" % (
            len(dangling), "; ".join("%s→%s" % (d["id"], d["resolved"]) for d in dangling[:5])))
    res["info"]["dangling_count"] = len(dangling)

    # 4. 链路闭合
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    pres_rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rel_map = {r.get("Id"): r for r in pres_rels}

    mlst = pres.find("p:sldMasterIdLst", NS)
    masters = []
    if mlst is not None:
        for m in mlst:
            rid = m.get(R + "id")
            rel = rel_map.get(rid)
            if rel is None:
                res["P0"].append("sldMasterIdLst 引用了不存在的关系 %s" % rid)
                continue
            masters.append(resolve("ppt/presentation.xml", rel.get("Target")))
    res["info"]["masters"] = masters

    layouts_ok = 0
    layout_owner = {}                      # slideLayout -> slideMaster（查字体继承要用）
    master_roots = {}
    for m in masters:
        mrels = m.replace("/slideMasters/", "/slideMasters/_rels/") + ".rels"
        if mrels not in names:
            res["P0"].append("母版缺少 rels：%s" % m)
            continue
        mr = ET.fromstring(z.read(mrels))
        try:
            master_roots[m] = ET.fromstring(z.read(m))
        except Exception:
            master_roots[m] = None
        has_theme = False
        for rel in mr:
            if rel.get("Type", "").endswith("/theme"):
                has_theme = True
                tgt = resolve(m, rel.get("Target"))
                if tgt not in names:
                    res["P0"].append("母版主题缺失：%s" % tgt)
            if rel.get("Type", "").endswith("/slideLayout"):
                tgt = resolve(m, rel.get("Target"))
                if tgt in names:
                    layouts_ok += 1
                    layout_owner[tgt] = m
        if not has_theme:
            res["P0"].append("母版没有 theme 关系：%s" % m)
    res["info"]["layouts_reachable"] = layouts_ok

    # 5. slides 链路
    sid_lst = pres.find("p:sldIdLst", NS)
    slide_parts = []
    orphan_all = []
    if sid_lst is not None:
        for sid in sid_lst:
            rid = sid.get(R + "id")
            rel = rel_map.get(rid)
            if rel is None:
                res["P0"].append("sldIdLst 引用了不存在的关系 %s" % rid)
                continue
            slide_parts.append(resolve("ppt/presentation.xml", rel.get("Target")))

    for sp in sorted(slide_parts, key=lambda s: int(re.search(r"slide(\d+)", s).group(1))):
        if sp not in names:
            res["P0"].append("幻灯片部件缺失：%s" % sp)
            continue
        root = ET.fromstring(z.read(sp))
        rels_p = sp.replace("/slides/", "/slides/_rels/") + ".rels"
        layout = None
        if rels_p in names:
            rx = ET.fromstring(z.read(rels_p))
            for rel in rx:
                if rel.get("Type", "").endswith("/slideLayout"):
                    layout = resolve(sp, rel.get("Target"))
        rec = {"slide": sp, "layout": layout}
        if layout is None:
            res["P0"].append("%s 没有 slideLayout 关系" % sp)
        elif layout not in names:
            res["P0"].append("%s 的版式缺失：%s" % (sp, layout))
        else:
            lroot = ET.fromstring(z.read(layout))
            lname = None
            cSld = lroot.find("p:cSld", NS)
            if cSld is not None:
                lname = cSld.get("name")
            lphs = set()
            for ph in lroot.iter("{%s}ph" % NS["p"]):
                lphs.add((ph.get("type") or "body", ph.get("idx")))
            sphs = []
            for ph in root.iter("{%s}ph" % NS["p"]):
                sphs.append((ph.get("type") or "body", ph.get("idx")))
            ltypes = {t for t, _ in lphs}
            orphans = []
            for t, i in sphs:
                if (t, i) in lphs:
                    continue
                if t in ("title", "ctrTitle") and ({"title", "ctrTitle"} & ltypes):
                    continue
                orphans.append({"type": t, "idx": i})
            rec.update({"layout_name": lname, "slide_phs": len(sphs),
                        "layout_phs": len(lphs), "orphans": orphans})
            if orphans:
                for o in orphans:
                    orphan_all.append((sp, lname, o))
            # 标题字体 / 字号 / 颜色：页面生效值 vs 版式槽位设定值
            tsp = _title_sp(root)
            if tsp is not None:
                mroot = master_roots.get(layout_owner.get(layout))
                rec["title_style"] = title_style_chain(tsp, lroot, mroot)
                sc, af = _autofit_scale(tsp)
                rec["fontScale"] = sc
                rec["autofit"] = af
        # 文字（带有效字号的）
        title = None
        for ph in root.iter("{%s}ph" % NS["p"]):
            if (ph.get("type") or "") in ("title", "ctrTitle"):
                sp_el = ph.getparent().getparent().getparent()
                title = "".join(t.text or "" for t in sp_el.iter("{%s}t" % NS["a"])).strip()
                break
        rec["title"] = (title or "")[:60]
        res["slides"].append(rec)

    # 5.51 孤儿占位符归并（逐页报会刷屏，按 类型×版式 归并）
    if orphan_all:
        grp = {}
        for sp_, lname_, o in orphan_all:
            grp.setdefault((o["type"], lname_ or "-"), []).append(
                posixpath.basename(sp_))
        bits = []
        for (t, ln), v in sorted(grp.items()):
            bits.append("%s @版式「%s」×%d 页（%s）" % (
                t, ln, len(v), ",".join(v[:4]) + ("…" if len(v) > 4 else "")))
        res["P1"].append("占位符无法被版式承接（孤儿）共 %d 处：%s"
                         % (len(orphan_all), "; ".join(bits)))

    # 5.6 标题字体 / 字号 / 颜色 是否偏离版式槽位设定 —— 旧版缺口所在
    dev, shrunk = [], []
    color_pairs = Counter()
    for s in res["slides"]:
        ts = s.get("title_style")
        if not ts:
            continue
        slot, eff = (ts.get("slot") or {}), (ts.get("effective") or {})
        diff = {}
        if slot.get("sz") and eff.get("sz") and int(eff["sz"]) != int(slot["sz"]):
            diff["字号"] = "%s → %s" % (_fmt_sz(slot["sz"]), _fmt_sz(eff["sz"]))
        if slot.get("latin") and eff.get("latin") \
                and eff["latin"].strip().lower() != slot["latin"].strip().lower():
            diff["字体"] = "%s → %s" % (slot["latin"], eff["latin"])
        if diff:
            dev.append({"slide": posixpath.basename(s["slide"]),
                        "title": s.get("title") or "",
                        "diff": diff,
                        "override_chain": [{"level": n,
                                            "set": {k: v for k, v in sig.items() if v}}
                                           for n, sig in ts.get("chain", [])]})
        if slot.get("fill") and eff.get("fill"):
            color_pairs[(slot["fill"], eff["fill"])] += 1
        sc = s.get("fontScale") or 1.0
        if sc < 1.0 and eff.get("sz"):
            shrunk.append({"slide": posixpath.basename(s["slide"]),
                           "fontScale": sc,
                           "size_pt": round(int(eff["sz"]) / 100.0 * sc, 1)})

    res["info"]["title_deviation_count"] = len(dev)
    res["info"]["title_deviations"] = dev
    if dev:
        head = "; ".join("%s %s" % (d["slide"],
                                    " ".join("%s %s" % (k, v) for k, v in d["diff"].items()))
                         for d in dev[:6])
        res["P1"].append(
            "标题字体/字号偏离版式设定 %d 页（应删掉页面级覆盖让它继承版式）：%s%s"
            % (len(dev), head, " …" if len(dev) > 6 else ""))

    if shrunk:
        sizes = [x["size_pt"] for x in shrunk if x.get("size_pt")]
        slot_sizes = sorted({_fmt_sz(((s.get("title_style") or {}).get("slot") or {}).get("sz"))
                             for s in res["slides"]
                             if s.get("fontScale", 1.0) < 1.0 and s.get("title_style")})
        res["info"]["title_autofit_shrunk"] = shrunk
        res["P1"].append(
            "标题靠 normAutofit 缩排 %d 页（有效字号 %s；这些页版式原字号 %s）——"
            "属「不换行」取舍，不是版式原始值，交付说明里要交代"
            % (len(shrunk),
               "%.1f–%.1f pt" % (min(sizes), max(sizes)) if sizes else "-",
               "/".join(x for x in slot_sizes if x != "-") or "-"))

    if color_pairs:
        res["info"]["title_color_pairs"] = ["%s → %s ×%d" % (k[0], k[1], v)
                                            for k, v in color_pairs.items()]
        # 按「版式设定色」分组：同一设定下出现多种实际色 = 不统一
        by_slot = {}
        for (slot_c, eff_c), n in color_pairs.items():
            by_slot.setdefault(slot_c, Counter())[eff_c] += n
        bad = {s: c for s, c in by_slot.items() if len(c) > 1}
        if bad:
            parts = ["版式设定 %s，实际 %s" % (
                s, " / ".join("%s ×%d" % (k, v) for k, v in c.most_common()))
                for s, c in sorted(bad.items())]
            res["P1"].append("标题色与版式设定不一致且各页不统一：%s" % "；".join(parts))

    # 5.5 OOXML 元素顺序
    check_order(z, names, res)

    # 6. python-pptx 打开
    try:
        from pptx import Presentation
        p = Presentation(str(path))
        res["info"]["pptx_open"] = "OK"
        res["info"]["slide_count"] = len(p.slides)
        res["info"]["size"] = [p.slide_width, p.slide_height]
        res["info"]["pptx_layouts"] = []
        for i, s in enumerate(p.slides, 1):
            try:
                res["info"]["pptx_layouts"].append({"slide": i, "layout": s.slide_layout.name})
            except Exception as e:
                res["info"]["pptx_layouts"].append({"slide": i, "layout": "ERR:%s" % e})
    except Exception as e:
        res["P0"].append("python-pptx 打开失败：%s" % e)
        res["info"]["pptx_open"] = "FAIL"

    return res


def to_md(r):
    L = []
    L.append("# PPTX 完整性体检：%s\n" % Path(r["file"]).name)
    L.append("- 结论：%s" % ("**错误 %d 项，不可交付**" % len(r["P0"]) if r["P0"] else "**通过**"))
    L.append("")
    if r["P0"]:
        L.append("## P0 错误\n")
        for e in r["P0"]:
            L.append("- %s" % e)
        L.append("")
    if r["P1"]:
        L.append("## P1 预警\n")
        for e in r["P1"]:
            L.append("- %s" % e)
        L.append("")
    L.append("## 链路\n")
    i = r["info"]
    L.append("- 打开测试：%s　页数：%s　画布：%s" % (i.get("pptx_open"), i.get("slide_count"), i.get("size")))
    L.append("- 母版：%s" % i.get("masters"))
    L.append("- 可达版式：%s　悬空引用：%s" % (i.get("layouts_reachable"), i.get("dangling_count")))
    L.append("- OOXML 元素顺序：%s" % i.get("element_order", "未检查"))
    L.append("")

    L.append("## 标题规范（字体 / 字号 / 颜色）\n")
    L.append("- 偏离版式设定的页数：%s" % i.get("title_deviation_count", 0))
    if i.get("title_autofit_shrunk"):
        sh = i["title_autofit_shrunk"]
        L.append("- 缩排页：%s" % ", ".join(
            "%s(%.1fpt@%.0f%%)" % (x["slide"], x["size_pt"], x["fontScale"] * 100) for x in sh))
    if i.get("title_color_pairs"):
        L.append("- 标题色分布：%s" % " / ".join(i["title_color_pairs"]))
    if i.get("title_deviations"):
        L.append("")
        L.append("| 页 | 标题 | 偏离 | 覆盖来自 |")
        L.append("|---|---|---|---|")
        for d in i["title_deviations"]:
            L.append("| %s | %s | %s | %s |" % (
                d["slide"], d["title"][:34],
                " ".join("%s %s" % (k, v) for k, v in d["diff"].items()),
                "; ".join("%s{%s}" % (c["level"], ",".join(
                    "%s=%s" % (k, v) for k, v in c["set"].items()))
                    for c in d["override_chain"]) or "-"))
    L.append("")

    L.append("## 逐页\n")
    L.append("| 页 | 标题 | 所挂版式 | 生效字号 | 生效字体 | 页内占位符 | 版式占位符 | 孤儿 |")
    L.append("|---|---|---|---|---|---|---|---|")
    for s in r["slides"]:
        ts = (s.get("title_style") or {}).get("effective", {})
        sc = s.get("fontScale") or 1.0
        sz = ts.get("sz")
        szs = ("%.1fpt" % (int(sz) / 100.0)) if sz else "-"
        if sz and sc < 1.0:
            szs += "→%.1fpt" % (int(sz) / 100.0 * sc)
        L.append("| %s | %s | %s | %s | %s | %s | %s | %s |" % (
            posixpath.basename(s["slide"]), s.get("title") or "-",
            s.get("layout_name") or "-", szs, ts.get("latin") or "-",
            s.get("slide_phs", "-"), s.get("layout_phs", "-"),
            ",".join("%s#%s" % (o["type"], o["idx"]) for o in s.get("orphans", [])) or "-"))
    L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None)
    ap.add_argument("--json", action="store_true", help="把完整结果打成 JSON 到 stdout")
    args = ap.parse_args()
    p = Path(args.pptx)
    r = check(p)
    out = Path(args.out) if args.out else None
    if out:
        out.mkdir(parents=True, exist_ok=True)
        (out / "verify.json").write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding="utf-8")
        (out / "verify.md").write_text(to_md(r), encoding="utf-8")

    if args.json:
        print(json.dumps(r, ensure_ascii=False, indent=2))
        return 1 if r["P0"] else 0

    print("=" * 96)
    print("体检：%s" % p.name)
    print("=" * 96)
    if r["P0"]:
        print("结论：P0 错误 %d 项" % len(r["P0"]))
        for e in r["P0"]:
            print("  [P0] %s" % e)
    else:
        print("结论：通过（P0 = 0）")
    for e in r["P1"]:
        print("  [P1] %s" % e)
    i = r["info"]
    print("  打开=%s 页数=%s 母版=%s 可达版式=%s 悬空=%s" % (
        i.get("pptx_open"), i.get("slide_count"), i.get("masters"),
        i.get("layouts_reachable"), i.get("dangling_count")))
    print()
    print("%-4s %-38s %-10s %-18s %s" % ("页", "标题", "生效字号", "生效字体", "所挂版式"))
    print("-" * 96)
    for s in r["slides"]:
        ts = (s.get("title_style") or {}).get("effective", {})
        sc = s.get("fontScale") or 1.0
        sz = ts.get("sz")
        szs = ("%.1fpt" % (int(sz) / 100.0)) if sz else "-"
        if sz and sc < 1.0:
            szs = "%.1fpt→%.1f" % (int(sz) / 100.0, int(sz) / 100.0 * sc)
        print("%-4s %-38s %-10s %-18s %s" % (
            posixpath.basename(s["slide"]), (s.get("title") or "-")[:38],
            szs, (ts.get("latin") or "-")[:18], s.get("layout_name") or "-"))
    if out:
        print("\n报告：%s" % (out / "verify.md"))
    return 1 if r["P0"] else 0


if __name__ == "__main__":
    sys.exit(main())
