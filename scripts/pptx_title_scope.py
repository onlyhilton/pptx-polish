# -*- coding: utf-8 -*-
"""pptx_title_scope.py — 「哪些占位符算标题」这件事的唯一定义（fix / verify 共用）。

为什么要有这个模块
------------------
"标题不能换行"这条规则要同时被两个脚本使用：
  · fix_title_oneline.py    —— 找出会换行的标题，缩字号
  · verify_title_lines.py   —— 用 PowerPoint 排版引擎核对是否真的单行
两边如果各自实现一遍"什么叫标题"，就会一边修 A、一边验 B，永远对不上。
所以范围定义只写一次，放在这里。

范围（三档，逐档收窄）
--------------------
① **标题占位符**：页面上 `p:ph type="title"` / `ctrTitle`。永远纳入 —— 标题就是要单行。
② **标题槽位**（模板驱动，默认纳入）：版式里那些**提示文本很短且框只够一行**的 body 槽位。
   判定依据不是槽位下标（换模板就失效），而是**版式自己写的提示文本**：
     · 提示文本非空、无硬换行、长度 ≤ 24 字符
     · 该槽位的框「容行数」≤ 1
   在 Reyee 模板上，这一条正好命中：`Section Title` 布局的正文标题(idx 17)、
   `目录` / `Workshop index` 的分节名(idx 17–20/25) 与序号(idx 21–26)、
   封面的 `Speaker` / `Date & Location`。而内容占位符（提示文本长、框能放十几行）自动排除。
③ **副标题**（默认排除）：`subTitle` 换行是设计意图（封面那行长句本来就该折两行）。

"容行数"怎么算
--------------
    容行数 = floor((框高 - 上内边距 - 下内边距) / (字号 × 行距))
行距默认 100%（模板 lnSpc 均为 100%），单行高 ≈ 字号 × 1.2。

字体度量
--------
量文本宽度必须有**真实字体文件**，否则只能退化估算（本模块会显式标记 `approx`，
调用方应据此放大安全余量，而不是假装准确）。
"""
from __future__ import annotations

import os
import posixpath
import re

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}

EMU_PT = 12700.0
LINE_FACTOR = 1.2
# OOXML 默认内边距：lIns=rIns=91440 EMU=7.2pt；tIns=bIns=45720 EMU=3.6pt
DEF_INS = {"l": 7.2, "r": 7.2, "t": 3.6, "b": 3.6}

TITLE_TYPES = ("title", "ctrTitle")
SKIP_TYPES = ("subTitle", "sldNum", "dt", "ftr")

# 字体目录**现算**，不写死某个用户的字体目录绝对路径。
# 为什么必须现算：找不到字体文件时 `find_font()` 会退化成"按 0.55 字宽粗估"，
# 于是"这个标题放不放得下"判错 —— 该缩的没缩、不该缩的被缩。这是功能缺陷，不只是路径难看。
# 顺序：用户级字体目录（后装的 Noto 装在这里）→ 系统字体目录 → 其它平台常见位置。
def _font_dirs():
    ds = []
    la = os.environ.get("LOCALAPPDATA")
    if la:
        ds.append(os.path.join(la, "Microsoft", "Windows", "Fonts"))
    windir = os.environ.get("WINDIR") or os.environ.get("SystemRoot")
    if windir:
        ds.append(os.path.join(windir, "Fonts"))
    # 非 Windows（字体度量走 PIL，本身平台无关）
    ds += ["/usr/share/fonts", "/usr/local/share/fonts", "/Library/Fonts",
           os.path.expanduser("~/Library/Fonts"), os.path.expanduser("~/.fonts"),
           os.path.expanduser("~/.local/share/fonts")]
    return [d for d in ds if os.path.isdir(d)]


FONT_DIRS = _font_dirs()

# 字体族 → 候选**文件名**，按优先级。目录由 FONT_DIRS 现算。
# 判定顺序是「文件名优先、目录其次」：先在第一优先的文件名里遍历所有目录，
# 找不到才轮到下一个文件名 —— 这样"同族优先选常规字重"的语义与旧的扁平列表完全一致。
FONT_CANDIDATES = {
    "noto sans bold": ["noto-sans-bold.ttf", "NotoSans-Bold.ttf",
                       "NotoSansSC-VF.ttf", "NotoSansCJKsc-Bold.otf"],
    "noto sans": ["noto-sans-regular.ttf", "noto-sans-bold.ttf",
                  "NotoSans-Regular.ttf", "NotoSansSC-VF.ttf",
                  "NotoSansCJKsc-Regular.otf"],
    "arial": ["arial.ttf", "Arial.ttf", "LiberationSans-Regular.ttf",
              "DejaVuSans.ttf"],
    "arial bold": ["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf",
                   "DejaVuSans-Bold.ttf"],
}


def find_font(typeface, extra=None):
    """返回 (字体文件路径 或 None, 是否精确)。找不到精确字体时退回同族相近字重并标记不精确。"""
    key = (typeface or "").strip().lower()
    if extra and key in extra:
        return extra[key], True
    fns = FONT_CANDIDATES.get(key)
    if not fns:
        return None, False
    # 文件名优先、目录其次
    for fn in fns:
        for d in FONT_DIRS:
            p = os.path.join(d, fn)
            if os.path.exists(p):
                return p, True
    # 同族其它字重（度量近似，标记为不精确）
    if key:
        fam = key.split()[0]
        for k, v in FONT_CANDIDATES.items():
            if k.split()[0] != fam:
                continue
            for fn in v:
                for d in FONT_DIRS:
                    p = os.path.join(d, fn)
                    if os.path.exists(p):
                        return p, False
    return None, False


# ---------------------------------------------------------------- XML 小工具

def read_rels(z, part, names):
    rp = posixpath.join(posixpath.dirname(part), "_rels",
                        posixpath.basename(part) + ".rels")
    out = {}
    if rp in names:
        for rel in etree.fromstring(z.read(rp)):
            out[rel.get("Id")] = (rel.get("Type") or "", rel.get("Target"))
    return out


def resolve(part, target):
    if not target:
        return None
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def slide_number(part):
    m = re.search(r"slide(\d+)\.xml$", part)
    return int(m.group(1)) if m else 0


def layout_re():
    return re.compile(r"ppt/slideLayouts/\w*[Ss]lideLayout\d+\.xml$")


def master_re():
    return re.compile(r"ppt/slideMasters/\w*[Ss]lideMaster\d+\.xml$")


def text_of(body):
    """形状里的全部文字（按段落顺序，硬换行归一为 \\n）"""
    chunks = []
    for m in re.finditer(r"<a:t>(.*?)</a:t>|<a:br\s*/>", body, re.S):
        if m.group(0).startswith("<a:br"):
            chunks.append("\n")
        else:
            chunks.append(re.sub(r"<[^>]+>", "", m.group(1)))
    return "".join(chunks)


def xfrm_of(body):
    m = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"/>',
                  body)
    if not m:
        return None
    return tuple(int(v) / EMU_PT for v in m.groups())


def insets_of(body):
    """形状自己声明的内边距；没声明的用 OOXML 默认值。"""
    o = dict(DEF_INS)
    m = re.search(r"<a:bodyPr([^>]*?)/>", body)
    if not m:
        m = re.search(r"<a:bodyPr([^>]*)>", body)
    attrs = m.group(1) if m else ""
    for k, a in (("l", "lIns"), ("r", "rIns"), ("t", "tIns"), ("b", "bIns")):
        mm = re.search(a + r'="(-?\d+)"', attrs)
        if mm:
            o[k] = int(mm.group(1)) / EMU_PT
    return o


def autofit_of(body):
    """返回 (kind, fontScale)  kind ∈ none|noAutofit|normAutofit|spAutoFit"""
    inner = ""
    m = re.search(r"<a:bodyPr([^>]*)>(.*?)</a:bodyPr>", body, re.S)
    if not m:
        return "none", None
    inner = m.group(2)
    fs = re.search(r'<a:normAutofit[^>]*?fontScale="(\d+)"', inner)
    if "<a:normAutofit" in inner:
        return "normAutofit", (int(fs.group(1)) if fs else None)
    if "<a:noAutofit" in inner:
        return "noAutofit", None
    if "<a:spAutoFit" in inner:
        return "spAutoFit", None
    return "none", None


def def_style(xml, level=1):
    """取 lstStyle 里某一级的 <a:defRPr>，返回 (attrs, inner)。"""
    m = re.search(r"<a:lstStyle>(.*?)</a:lstStyle>", xml, re.S)
    if not m:
        return None, ""
    m2 = re.search(r'<a:lvl%dpPr[^>]*>\s*<a:defRPr([^>]*?)(/>|>(.*?)</a:defRPr>)' % level,
                   m.group(1), re.S)
    if not m2:
        return None, ""
    return m2.group(1), (m2.group(3) or "")


def font_of(attrs, inner=""):
    """(字号pt, 字体名)"""
    sz = None
    if attrs:
        m = re.search(r'sz="(\d+)"', attrs)
        sz = int(m.group(1)) / 100.0 if m else None
    fam = None
    if inner:
        m = re.search(r'<a:latin typeface="([^"]*)"', inner)
        fam = m.group(1) if m else None
    return sz, fam


def capacity(box, ins, size_pt, lnspc=1.0):
    """这个框能放几行"""
    if not box or not size_pt or size_pt <= 0:
        return 99
    usable = box[3] - ins["t"] - ins["b"]
    if usable <= 0:
        return 0
    return int(usable // (size_pt * LINE_FACTOR * lnspc))


def run_overrides(body):
    """段落/run 级别的 rPr 覆盖（会盖过版式），返回 [(sz, latin)]"""
    out = []
    for m in re.finditer(r"<a:rPr([^>]*?)(/>|>(.*?)</a:rPr>)", body, re.S):
        attrs, inner = m.group(1), (m.group(3) or "")
        sz, fam = font_of(attrs, inner)
        out.append((sz, fam))
    for m in re.finditer(r"<a:pPr[^>]*>\s*<a:defRPr([^>]*?)(/>|>(.*?)</a:defRPr>)", body, re.S):
        attrs, inner = m.group(1), (m.group(3) or "")
        out.append(font_of(attrs, inner))
    return out


def iter_shapes(xml):
    for m in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
        yield m.group(0)


def shape_ph(body):
    m = re.search(r"<p:ph([^>]*)/>", body)
    if not m:
        return None
    att = m.group(1)
    return {
        "type": (re.search(r'type="([^"]*)"', att) or [None, "body"])[1],
        "idx": (re.search(r'idx="([^"]*)"', att) or [None, "0"])[1],
    }


def shape_name(body):
    m = re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"', body)
    return m.group(2) if m else ""


def shape_id(body):
    m = re.search(r'<p:cNvPr id="(\d+)"', body)
    return m.group(1) if m else None


# ---------------------------------------------------------------- 索引构建

def build_index(z, base="ppt"):
    """扫全包，建版本式槽位与"标题槽位"集合。返回 dict。"""
    names = set(z.namelist())
    lre = layout_re()

    layouts = {}
    for n in sorted(names):
        if not lre.match(n):
            continue
        xml = z.read(n).decode("utf-8", "ignore")
        cs = re.search(r'<p:cSld name="([^"]*)"', xml)
        slots = {}
        for b in iter_shapes(xml):
            ph = shape_ph(b)
            if ph is None:
                continue
            attrs, inner = def_style(b)
            sz, fam = font_of(attrs, inner)
            box = xfrm_of(b)
            ins = insets_of(b)
            slots[(ph["type"], ph["idx"])] = {
                "name": shape_name(b), "type": ph["type"], "idx": ph["idx"],
                "size": sz, "font": fam, "box": box, "ins": ins,
                "prompt": text_of(b).strip(),
                "capacity": capacity(box, ins, sz),
            }
        layouts[n] = {"name": (cs.group(1) if cs else ""), "slots": slots, "xml": xml}

    # 标题槽位：版式说明它只放一行（提示文本短、无折行、框只够一行）
    heading_slots = set()
    for L in layouts.values():
        for key, s in L["slots"].items():
            if key[0] in TITLE_TYPES or key[0] in SKIP_TYPES:
                continue
            p = s["prompt"]
            if not p or len(p) > 24 or "\n" in p:
                continue
            if s["capacity"] <= 1:
                heading_slots.add(key)

    slides = {}
    for n in sorted(names, key=lambda s: slide_number(s)):
        if not re.match(r"ppt/slides/slide\d+\.xml$", n):
            continue
        lay = None
        for rid, (typ, tgt) in read_rels(z, n, names).items():
            if typ.endswith("/slideLayout"):
                lay = resolve(n, tgt)
        slides[n] = {
            "number": slide_number(n),
            "layout": lay,
            "layout_name": layouts.get(lay, {}).get("name", ""),
        }
    return {"layouts": layouts, "slides": slides, "heading_slots": heading_slots,
            "names": names}


def parse_slots(spec):
    """把 "body:17,title:0" 解析成 {("body","17"), ("title","0")}"""
    if not spec:
        return None
    out = set()
    for item in (spec if isinstance(spec, (list, tuple, set)) else str(spec).split(",")):
        item = str(item).strip()
        if not item:
            continue
        if ":" in item:
            t, i = item.split(":", 1)
            out.add((t.strip() or "body", i.strip()))
        else:
            out.add(("body", item))
    return out


def iter_targets(z, index, slide_part, include_heading=True, include_subtitle=False,
                 slots=None):
    """产出该页所有"需单行"的占位符信息。

    每个 dict：page / name / ph_type / ph_idx / text / size / font / box / ins /
    capacity / avail_w / hard_break / page_override / autofit / font_scale / layout
    """
    xml = z.read(slide_part).decode("utf-8", "ignore")
    meta = index["slides"].get(slide_part, {})
    L = index["layouts"].get(meta.get("layout"), {"slots": {}, "name": ""})
    want = parse_slots(slots)

    for b in iter_shapes(xml):
        ph = shape_ph(b)
        if ph is None:
            continue
        key = (ph["type"], ph["idx"])
        if ph["type"] in SKIP_TYPES and not (include_subtitle and ph["type"] == "subTitle"):
            continue
        in_scope = ph["type"] in TITLE_TYPES
        if not in_scope and include_heading and key in index["heading_slots"]:
            in_scope = True
        if want is not None:
            in_scope = key in want          # 显式指定槽位时以指定为准
        if not in_scope:
            continue
        text = text_of(b)
        if not text.strip():
            continue
        slot = L["slots"].get(key, {})
        box = xfrm_of(b) or slot.get("box")
        ins = insets_of(b)
        if not re.search(r"<a:bodyPr", b):
            ins = slot.get("ins", dict(DEF_INS))
        p_sz, p_fam = def_style(b)                  # 页面 lstStyle
        ov = run_overrides(b)                        # 段落 / run 覆盖
        ov_sz = next((s for s, _ in ov if s), None)
        ov_fam = next((f for _, f in ov if f), None)
        size = p_sz or ov_sz or slot.get("size")
        font = p_fam or ov_fam or slot.get("font")
        af, fs = autofit_of(b)
        avail_w = (box[2] - ins["l"] - ins["r"]) if box else None
        yield {
            "page": meta.get("number", 0), "part": slide_part,
            "shape": b, "shape_id": shape_id(b),
            "name": shape_name(b), "ph_type": ph["type"], "ph_idx": ph["idx"],
            "text": text, "size": size, "font": font, "box": box, "ins": ins,
            "capacity": capacity(box, ins, size),
            "avail_w": avail_w,
            "hard_break": "\n" in text,
            "page_override": bool(p_sz or ov_sz or p_fam or ov_fam),
            "autofit": af, "font_scale": fs,
            "layout_name": L.get("name", ""), "is_title": ph["type"] in TITLE_TYPES,
        }


def measure_width_pt(text, size_pt, font_file, fallback_ratio=0.55):
    """文本宽度（pt）。有字体文件用真实度量；没有就按平均字宽估算。"""
    if not size_pt:
        return 0.0
    if font_file and os.path.exists(font_file):
        try:
            from PIL import ImageFont
            return ImageFont.truetype(font_file, 512).getlength(text) / 512.0 * size_pt
        except Exception:                                     # noqa: BLE001
            pass
    return fallback_ratio * size_pt * len(text)
