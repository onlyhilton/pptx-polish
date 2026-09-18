#!/usr/bin/env python3
"""audit_deck.py — 对 .pptx 做 XML 级量化审计，输出 audit.json + audit.md。

用途：pptx-polish skill 的 Step 1（解构原件）与 Step 2（解析母版）。
依赖：python-pptx（Windows / macOS / Linux 均可；本 skill 的渲染步骤需要 Windows + PowerPoint）
不依赖 LibreOffice / poppler。

用法：
    python audit_deck.py "<file.pptx>" --out "<out_dir>" [--extract-images]

产物：
    <out_dir>/audit.json              结构化数据（重建时的事实基座）
    <out_dir>/audit.md                人读报告（含逐页文字全文）
    <out_dir>/assets/*.png|jpg        原件的图片素材（仅 --extract-images 时产出，不过滤小图）
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

# 标题判定规则与规范化脚本共用同一份实现（见 deck_rules.py 顶部说明）
from deck_rules import (  # noqa: E402
    COVER_ZONE_PX, MIN_TITLE_PT, TITLE_MAX_CHARS, TITLE_TOP_BAND, TITLE_ZONE_PX,
    image_geometry, parse_src_rect, pick_title, title_zone_for,
)

EMU_PER_PX = 9525
SAFE_MARGIN_LR = 60   # design-principle.md 的左右安全边距
SAFE_MARGIN_TB = 20   # 上下安全边距
NEAR_EDGE_PX = 20
NEUTRAL_SPREAD = 14   # RGB 极差小于此值视为中性色
NEUTRAL_SAT = 0.25    # 明度居中时，饱和度低于此值视为灰调（不计入单页颜色数）
OVERLAP_RATIO = 0.30  # 文本形状互相覆盖比例阈值
RESERVED_MIN_RATIO = 0.05   # 内容压到版式装饰保留区的最小判定比例
RESERVED_BIG_ZONE = 3000    # 保留区面积(px²)阈值：≥ 视为 LOGO/大装饰（P1），否则细装饰条（P2）
CJK_RE = re.compile(r"[\u2e80-\u9fff\uff00-\uffef\u3000-\u303f]")
MAX_PAGE_COLORS = 3   # 单页有彩色上限（黑白与中性灰不计入）
ZOOM_LOW = 0.75       # 图片过采样不足 1.33x 判为"分辨率余量不足"（见 collect_image 注释）

PLACEHOLDER_PAT = re.compile(
    r"单击此处|单击以添加|点击输入|在此键入|请在此|输入标题|输入文本"
    r"|Click to (edit|add)|Type the title|Your text here|Lorem ipsum",
    re.IGNORECASE,
)

THEME_ALIAS = {
    "ACCENT_1": "accent1", "ACCENT_2": "accent2", "ACCENT_3": "accent3",
    "ACCENT_4": "accent4", "ACCENT_5": "accent5", "ACCENT_6": "accent6",
    "TEXT_1": "dk1", "BACKGROUND_1": "lt1",
    "TEXT_2": "dk2", "BACKGROUND_2": "lt2",
    "DARK_1": "dk1", "LIGHT_1": "lt1", "DARK_2": "dk2", "LIGHT_2": "lt2",
    "HYPERLINK": "hlink", "FOLLOWED_HYPERLINK": "folHlink",
}

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def to_px(emu):
    return None if emu is None else round(emu / EMU_PER_PX, 1)


def to_pt(length):
    try:
        return None if length is None else round(length.pt, 1)
    except Exception:
        return None


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def load_theme_colors(pptx_path: Path) -> dict:
    """解析 ppt/theme/theme1.xml 的 clrScheme，返回 {dk1: '#RRGGBB', ...}。"""
    out: dict = {}
    try:
        with zipfile.ZipFile(str(pptx_path)) as z:
            names = sorted(n for n in z.namelist()
                           if n.startswith("ppt/theme/theme") and n.endswith(".xml"))
            if not names:
                return out
            root = ET.fromstring(z.read(names[0]))
        scheme = root.find(".//{%s}clrScheme" % A_NS)
        if scheme is None:
            return out
        for child in scheme:
            tag = child.tag.split("}")[-1]
            srgb = child.find("{%s}srgbClr" % A_NS)
            sysc = child.find("{%s}sysClr" % A_NS)
            if srgb is not None and srgb.get("val"):
                out[tag] = "#" + srgb.get("val").upper()
            elif sysc is not None and sysc.get("lastClr"):
                out[tag] = "#" + sysc.get("lastClr").upper()
    except Exception:
        pass
    return out


def color_of(color_format, theme_map: dict):
    """把 ColorFormat 归一化为 #RRGGBB；主题色经 theme_map 解析。"""
    if color_format is None:
        return None
    rgb = safe(lambda: color_format.rgb)
    if rgb is not None:
        return "#" + str(rgb).upper()
    tc = safe(lambda: color_format.theme_color)
    if tc is not None:
        name = getattr(tc, "name", "") or ""
        tag = THEME_ALIAS.get(name)
        if tag:
            return theme_map.get(tag) or ("theme:%s" % name)
    return None


def is_neutral(hex_color: str) -> bool:
    """判断这个色是否应算作"中性"（不计入单页颜色数）。

    只看 RGB 极差会把两类色误判成有彩色：
      · 很暗的近黑深色（如 #18253A，正文文字色）—— 视觉上就是黑
      · 很亮的近白浅色（如 #F0F6FF，浅底块）—— 视觉上就是白
      · 低饱和的灰调（如 #63758A，次要文字）—— 视觉上就是灰
    这三类在"单页不超过三个颜色（黑白除外）"里都不该占名额。
    """
    if not hex_color or not hex_color.startswith("#") or len(hex_color) != 7:
        return False
    try:
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
    except ValueError:
        return False
    mx, mn = max(r, g, b), min(r, g, b)
    spread = mx - mn
    if spread < NEUTRAL_SPREAD:          # 极差小 → 灰
        return True
    lum = (mx + mn) / 510.0
    if lum <= 0.30:                       # 很暗 → 视觉近黑
        return True
    if lum >= 0.90:                       # 很亮 → 视觉近白（浅底）
        return True
    sat = spread / (mx + mn) if lum <= 0.5 else spread / (510 - mx - mn)
    return sat <= NEUTRAL_SAT             # 低饱和 → 灰调


# ---- 组合形状（grpSp）的坐标变换 ----------------------------------------
# python-pptx **不做**组合变换：组合内子形状的 .left/.top/.width/.height 返回的
# 是 `<a:off>/<a:ext>` 里的**子坐标系**原值（见 GroupShape 文档），不是页面绝对坐标。
# 直接拿去和别的形状比，就会把相距很远的两个元素判成完全重合。
# 实测：P22 两个 PoE 徽标的绝对位置相差 107px，各自的子坐标系却完全相同
# （两个组的 chOff/chExt 一致），于是被判成"100% 完全叠印"，差点被当成
# "编辑残留"删掉一个可见徽标。
# 正确做法：绝对坐标 = off + (child - chOff) × (ext / chExt)，逐层累积。
IDENT = (1.0, 1.0, 0.0, 0.0)


def _child_map(shp):
    """组合自身的"子坐标系 → 父坐标系"映射 (sx, sy, dx, dy)；非组合返回 None。"""
    el = getattr(shp, "_element", None)
    if el is None or not el.tag.endswith("}grpSp"):
        return None
    gsp = el.find("{%s}grpSpPr" % P_NS)
    xfrm = gsp.find("{%s}xfrm" % A_NS) if gsp is not None else None
    if xfrm is None:
        return None
    o, e = xfrm.find("{%s}off" % A_NS), xfrm.find("{%s}ext" % A_NS)
    co, ce = xfrm.find("{%s}chOff" % A_NS), xfrm.find("{%s}chExt" % A_NS)
    if None in (o, e, co, ce):
        return None
    try:
        cw, ch = int(ce.get("cx")), int(ce.get("cy"))
        if cw == 0 or ch == 0:
            return None
        sx, sy = int(e.get("cx")) / cw, int(e.get("cy")) / ch
        return (sx, sy, int(o.get("x")) - int(co.get("x")) * sx,
                int(o.get("y")) - int(co.get("y")) * sy)
    except (TypeError, ValueError):
        return None


def compose(outer, inner):
    """先套 inner 再套 outer（嵌套组合时逐层累积）。"""
    if not outer:
        return inner
    if not inner:
        return outer
    ox, oy, odx, ody = outer
    ix, iy, idx_, idy = inner
    return (ox * ix, oy * iy, ox * idx_ + odx, oy * idy + ody)


def abs_geom(shp, tr=None):
    """变换到页面绝对坐标的 (left, top, width, height)，单位 EMU。"""
    l, t, w, h = shp.left, shp.top, shp.width, shp.height
    if None in (l, t, w, h):
        return None
    if not tr:
        return (l, t, w, h)
    sx, sy, dx, dy = tr
    return (l * sx + dx, t * sy + dy, w * sx, h * sy)


def rect_of(shape, tr=None):
    g = abs_geom(shape, tr)
    if g is None:
        return None
    return (g[0], g[1], g[0] + g[2], g[1] + g[3])


def overlap_ratio(a, b) -> float:
    x1, y1 = max(a[0], b[0]), max(a[1], b[1])
    x2, y2 = min(a[2], b[2]), min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    inter = (x2 - x1) * (y2 - y1)
    smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
    return inter / smaller if smaller > 0 else 0.0


def area_of(r) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def visual_rect(shape, r):
    """文本框的"实际绘制范围"。

    常见坑：文本框被拉得很宽（R 远大于文字末端），若按整个矩形判冲突会大量误报。
    这里对"左对齐的纯文本框"按字号估算文字宽度收缩右边界；居中/右对齐时整框都可能被占，不收缩。
    """
    if not safe(lambda: shape.has_text_frame, False):
        return r
    tf = shape.text_frame
    for para in tf.paragraphs:
        a = safe(lambda: para.alignment)
        if a is not None and "LEFT" not in str(a).upper():
            return r
    best = 0.0
    for para in tf.paragraphs:
        w = 0.0
        for run in para.runs:
            pt = to_pt(safe(lambda: run.font.size)) or 18.0
            for ch in (run.text or ""):
                w += pt * (1.0 if CJK_RE.match(ch) else 0.52)
        best = max(best, w)
    if best <= 0:
        return r
    est_px = best / 72.0 * 96.0
    if est_px >= to_px(r[2] - r[0]) - 4:
        return r
    return (r[0], r[1], r[0] + int(est_px * EMU_PER_PX), r[3])


def blank_image(blob) -> bool:
    """全透明图片判定（编辑残留的空占位）。

    这类图放大多少倍都不会"发虚"——它根本不产生任何像素。把它算进
    "分辨率不足/被放大"只会制造噪音。（实测：P25 一张 48×48 全透明图和
    122×122 的显示框，修正组合坐标后才暴露出来。）
    """
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(blob))
        if "A" not in im.getbands() and "transparency" not in im.info:
            return False
        return im.convert("RGBA").getchannel("A").getextrema()[1] == 0
    except Exception:
        return False


def collect_image(shp, slide_idx: int, extract_root, tr=None):
    """记录一张图片；extract_root 非空时同时落盘。不过滤任何尺寸的图片。

    tr：组合变换。图片若在组合里，显示尺寸必须换算到绝对坐标，
    否则拿子坐标系的尺寸去除原图像素 —— 会算出完全错误的"放大倍率"。
    """
    img = safe(lambda: shp.image)
    if img is None:
        return None
    ext = (safe(lambda: img.ext) or "png").lower()
    blob = safe(lambda: img.blob)
    if not blob:
        return None
    px_size = safe(lambda: img.size) or (None, None)
    pw, ph = px_size[0], px_size[1]
    g = abs_geom(shp, tr) or (shp.left, shp.top, shp.width, shp.height)
    disp = [to_px(g[2]), to_px(g[3])]
    pos = [to_px(g[0]), to_px(g[1])]
    shape_id = safe(lambda: shp.shape_id) or 0
    fname = "p%02d_s%02d.%s" % (slide_idx, shape_id, ext)
    saved = None
    if extract_root is not None:
        try:
            (extract_root).mkdir(parents=True, exist_ok=True)
            target = extract_root / fname
            target.write_bytes(blob)
            saved = str(target)
        except Exception:
            saved = None
    # 清晰度判据 = 过采样倍率，不是绝对像素数。
    # 绝对像素数是错的：333px 的图缩到 145px 显示（2.3x 过采样）是清晰的，
    # 而 401px 的图被拉到 455px 显示（1.13x）才是真的糊。
    # zoom 见下方计算（显示尺寸 / srcRect 之后的"有效原图区域"）。

    # srcRect / 有效区域 / 变形 / 过采样倍率的计算与 fix_geometry.py 共用
    # deck_rules.py（同一判定只允许一份实现）。
    src_rect = None
    try:
        sr = shp._element.find(".//" + "{%s}srcRect" % A_NS)
        if sr is not None:
            src_rect = parse_src_rect(sr)
    except Exception:
        pass
    blank = blank_image(blob)
    geo = image_geometry(src_rect, (pw, ph), disp)
    if blank:
        # 全透明空图：谈"变形 / 清晰度"没有意义 —— 会把
        # 122px 显示框 / 48px 原图的 2.54x 误报成"放大"。比例字段保留。
        geo["src_aspect"] = None
        geo["distort_pct"] = None
        geo["zoom"] = None
    distort_pct = geo["distort_pct"]
    zoom = geo["zoom"]

    low_quality = bool(zoom is not None and zoom > ZOOM_LOW)
    return {
        "file": fname,
        "saved_as": saved,
        "in_group": bool(tr),
        "blank": blank,
        "name": safe(lambda: shp.name) or "",
        "native_px": [pw, ph],
        "display_px": disp,
        "pos_px": pos,
        "src_rect": src_rect,
        "src_aspect": geo["src_aspect"],
        "display_aspect": geo["display_aspect"],
        "distort_pct": distort_pct,
        "zoom": zoom,
        "bytes": len(blob),
        "low_quality": low_quality,
    }


def walk_shapes(shapes, tr=None):
    """展平组合形状内的子形状，同时累积父级变换。

    yield (shape, transform)：transform 把该形状的名义坐标映射到页面绝对坐标
    （顶层形状为 None）。**任何几何判定都必须用它**，否则组合内形状的坐标是
    子坐标系的，和外部形状不可比。
    """
    for shp in shapes:
        yield shp, tr
        if safe(lambda: shp.shape_type) is not None and safe(lambda: len(shp.shapes)):
            sub_tr = compose(tr, _child_map(shp))
            for sub in walk_shapes(shp.shapes, sub_tr):
                yield sub


def audit(pptx_path: Path, extract_root=None) -> dict:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

    theme_map = load_theme_colors(pptx_path)
    prs = Presentation(str(pptx_path))
    sw = prs.slide_width
    sh = prs.slide_height

    all_fonts = Counter()
    all_sizes = Counter()
    all_colors = Counter()
    slides_out = []

    for idx, slide in enumerate(prs.slides, 1):
        texts, fonts, sizes, colors = [], Counter(), Counter(), Counter()
        n_shapes = n_pic = n_tbl = n_txt = 0
        leftovers, oob, near, text_rects = [], [], [], []
        images = []
        title_cand = []
        title_ph_text = None
        hardcoded = 0
        total_chars = 0

        layout_name = safe(lambda: slide.slide_layout.name) or ""
        # 该页的标题搜索策略（'full' 封面类 / 'top' 内容类 / None 无标题槽位），
        # 与规范化脚本共用同一规则，避免一边报、一边跳
        title_zone = title_zone_for(slide)

        for shp, tr in walk_shapes(slide.shapes):
            n_shapes += 1

            if safe(lambda: shp.shape_type == MSO_SHAPE_TYPE.PICTURE, False):
                n_pic += 1
                info = collect_image(shp, idx, extract_root, tr)
                if info:
                    images.append(info)
            if safe(lambda: shp.has_table, False):
                n_tbl += 1
                for row in shp.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            texts.append(cell.text.strip())
                            total_chars += len(cell.text.strip())

            fill_color = None
            if safe(lambda: shp.fill.type is not None, False):
                fill_color = color_of(safe(lambda: shp.fill.fore_color), theme_map)
                if fill_color:
                    colors[fill_color] += 1

            if safe(lambda: shp.has_text_frame, False):
                txt = shp.text_frame.text or ""
                if txt.strip():
                    n_txt += 1
                    texts.append(txt.strip())
                    total_chars += len(txt.strip())
                    r = rect_of(shp, tr)
                    if r:
                        text_rects.append((shp.shape_id, r, txt.strip()[:24], bool(tr)))
                    if PLACEHOLDER_PAT.search(txt):
                        leftovers.append(txt.strip()[:80])
                    shape_pt = 0.0
                    for para in shp.text_frame.paragraphs:
                        for run in para.runs:
                            fn = safe(lambda: run.font.name)
                            if fn:
                                fonts[fn] += 1
                                all_fonts[fn] += 1
                                hardcoded += 1
                            fs = to_pt(safe(lambda: run.font.size))
                            if fs:
                                sizes[fs] += 1
                                all_sizes[fs] += 1
                                shape_pt = max(shape_pt, fs)
                            fc = color_of(safe(lambda: run.font.color), theme_map)
                            if fc:
                                colors[fc] += 1
                                all_colors[fc] += 1
                    # 标题归位检测：占位符里的标题 vs 自由文本框冒充的标题
                    if safe(lambda: shp.is_placeholder, False):
                        ptype = safe(lambda: shp.placeholder_format.type)
                        if ptype in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                            title_ph_text = ((title_ph_text or "") + " " + txt.strip()).strip()
                    else:
                        g_top = abs_geom(shp, tr)
                        top_px = to_px(g_top[1]) if g_top else 0
                        zone_limit = (COVER_ZONE_PX if title_zone == "full"
                                      else TITLE_ZONE_PX)
                        if (title_zone is not None and top_px <= zone_limit
                                and len(txt.strip()) <= TITLE_MAX_CHARS):
                            title_cand.append((shape_pt, top_px, txt.strip()))

            r = rect_of(shp, tr)
            if r:
                if r[0] < 0 or r[1] < 0 or r[2] > sw or r[3] > sh:
                    oob.append({
                        "shape": safe(lambda: shp.name) or "?",
                        "text": (safe(lambda: shp.text_frame.text) or "")[:40],
                        "px": [to_px(r[0]), to_px(r[1]), to_px(r[2]), to_px(r[3])],
                    })
                elif (r[0] < SAFE_MARGIN_LR * EMU_PER_PX
                      or r[1] < SAFE_MARGIN_TB * EMU_PER_PX
                      or r[2] > sw - SAFE_MARGIN_LR * EMU_PER_PX
                      or r[3] > sh - SAFE_MARGIN_TB * EMU_PER_PX):
                    near.append({
                        "shape": safe(lambda: shp.name) or "?",
                        "text": (safe(lambda: shp.text_frame.text) or "")[:30],
                        "px": [to_px(r[0]), to_px(r[1]), to_px(r[2]), to_px(r[3])],
                    })

        # ---- 版式/母版上的"装饰保留区"（非占位符形状 = 换模板后仍在原位的元素）
        reserved = []
        seen_zones = set()
        for host in (safe(lambda: slide.slide_layout, None),
                     safe(lambda: slide.slide_layout.slide_master, None)):
            if host is None:
                continue
            for lshp in safe(lambda: list(host.shapes), []) or []:
                if safe(lambda: lshp.is_placeholder, False):
                    continue
                lr = rect_of(lshp)
                if not (lr and area_of(lr) > 0):
                    continue
                # 版式上常有位置完全相同的重复装饰（同名/异名），按几何去重
                key = tuple(round(to_px(v)) for v in lr)
                if key in seen_zones:
                    continue
                seen_zones.add(key)
                reserved.append((safe(lambda: lshp.name) or "?", lr))

        zone_hits = []
        grid_hits = []
        layout_for_grid = safe(lambda: slide.slide_layout, None)
        gleft = gright = None
        if layout_for_grid is not None:
            for lshp in safe(lambda: list(layout_for_grid.placeholders), []) or []:
                lr = rect_of(lshp)
                if not lr:
                    continue
                pt = safe(lambda: lshp.placeholder_format.type)
                is_title = pt in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
                if is_title:
                    gleft = to_px(lr[0]) if gleft is None else min(gleft, to_px(lr[0]))
                else:
                    gright = to_px(lr[2]) if gright is None else max(gright, to_px(lr[2]))

        for shp, tr in walk_shapes(slide.shapes):
            if safe(lambda: shp.is_placeholder, False):
                continue
            r = rect_of(shp, tr)
            if not r:
                continue
            vis = visual_rect(shp, r)
            shp_txt = ""
            if safe(lambda: shp.has_text_frame, False):
                shp_txt = (safe(lambda: shp.text_frame.text) or "").strip()
            for zname, zr in reserved:
                ov = overlap_ratio(vis, zr)
                if ov >= RESERVED_MIN_RATIO:
                    zone_hits.append({
                        "zone": zname,
                        "shape": safe(lambda: shp.name) or "?",
                        "text": shp_txt[:40],
                        "ratio": round(ov, 2),
                        "zone_area": round(area_of(zr) / (EMU_PER_PX ** 2)),
                    })
            # 栅格约束只作用于"内容"：有文字的文本框、图片、表格。
            # 纯装饰形状（无文字的色块/线条）本就不受内容栅格约束 ——
            # 版式自带的装饰条还常落在栅格左侧（如 L=29 vs 栅格 64），
            # 拿它报错是误报（实测 19 处"越栅格"里 19 处都是装饰形状）。
            is_content = bool(shp_txt) or safe(
                lambda: shp.shape_type == MSO_SHAPE_TYPE.PICTURE, False)
            if is_content and to_px(vis[2] - vis[0]) > 4:
                # 容差 5px：小于此值的偏差是手工拖拽/浮点误差级别，
                # 报出来只是噪音（实测 2px 偏差会让 10 处里 4 处变成假警报）
                if gleft is not None and to_px(vis[0]) < gleft - 5:
                    grid_hits.append({"side": "左", "px": to_px(vis[0]),
                                      "grid": gleft, "text": shp_txt[:30]})
                if gright is not None and to_px(vis[2]) > gright + 5:
                    grid_hits.append({"side": "右", "px": to_px(vis[2]),
                                      "grid": gright, "text": shp_txt[:30]})

        overlaps = []
        for i in range(len(text_rects)):
            for j in range(i + 1, len(text_rects)):
                ratio = overlap_ratio(text_rects[i][1], text_rects[j][1])
                if ratio >= OVERLAP_RATIO:
                    rec = {
                        "a": text_rects[i][2], "b": text_rects[j][2],
                        "ratio": round(ratio, 2),
                    }
                    if text_rects[i][3] or text_rects[j][3]:
                        rec["in_group"] = True
                    overlaps.append(rec)

        suspected_title = None
        if not title_ph_text and title_cand:
            # 判定规则见 deck_rules.pick_title（与规范化脚本共用）
            pick = pick_title(title_cand)
            suspected_title = pick[2] if pick else None

        page_colors = {c: n for c, n in colors.items() if not is_neutral(c)}

        slides_out.append({
            "index": idx,
            "layout": layout_name,
            "text_chars": total_chars,
            "shape_count": n_shapes,
            "text_shape_count": n_txt,
            "picture_count": n_pic,
            "table_count": n_tbl,
            "fonts": dict(fonts),
            "font_sizes": dict(sorted(sizes.items(), key=lambda kv: -kv[1])),
            "colors": dict(sorted(colors.items(), key=lambda kv: -kv[1])),
            "color_count": len(page_colors),
            "hardcoded_fonts": hardcoded,
            "title_in_placeholder": title_ph_text,
            "suspected_title_textbox": suspected_title,
            "placeholder_leftovers": leftovers,
            "out_of_canvas": oob,
            "near_edge": near,
            "overlapping_text_pairs": overlaps,
            "reserved_zone_conflicts": zone_hits,
            "grid_overflow": grid_hits,
            "images": images,
            "texts": texts,
        })

    canvas_px = [to_px(sw), to_px(sh)]
    effective_colors = {c: n for c, n in all_colors.items() if not is_neutral(c)}
    fonts_clean = Counter({f: n for f, n in all_fonts.items() if f.strip()})

    findings = []
    for s in slides_out:
        idx = s["index"]
        if s["out_of_canvas"]:
            findings.append({
                "level": "P0", "slide": idx, "item": "元素超出画布",
                "detail": "; ".join("%s → px%s" % (o["shape"], o["px"]) for o in s["out_of_canvas"][:4]),
                "basis": "checklist §1.3 元素出画布",
            })
        if s["placeholder_leftovers"]:
            findings.append({
                "level": "P0", "slide": idx, "item": "残留模板占位符",
                "detail": "; ".join(s["placeholder_leftovers"][:3]),
                "basis": "checklist §2 未填占位符",
            })
        if s["overlapping_text_pairs"]:
            findings.append({
                "level": "P1", "slide": idx, "item": "文本形状重叠",
                "detail": "; ".join('"%s" × "%s" %.0f%%' % (o["a"], o["b"], o["ratio"] * 100)
                                    for o in s["overlapping_text_pairs"][:3]),
                "basis": "checklist §1.3 元素重叠",
            })
        # 全透明空图：编辑残留。它不产生任何像素，不能算进"分辨率不足"
        blank_imgs = [im for im in s["images"] if im.get("blank")]
        if blank_imgs:
            findings.append({
                "level": "P2", "slide": idx, "item": "空图（全透明，不产生任何像素）",
                "detail": "；".join("%s %dx%d 显示 %sx%s" % (
                    im["file"], im["native_px"][0], im["native_px"][1],
                    im["display_px"][0], im["display_px"][1])
                    for im in blank_imgs[:3]) + " —— 编辑残留，删除零视觉风险",
                "basis": "checklist §2.2 图片有效内容",
            })
        # 图片：真变形（有效原图区域比例 ≠ 显示框比例）与放大发虚
        warped = [im for im in s["images"]
                  if im.get("distort_pct") is not None and im["distort_pct"] > 5]
        if warped:
            findings.append({
                "level": "P1", "slide": idx, "item": "图片被拉伸变形",
                "detail": "; ".join("%s 源比例 %.2f → 显示 %.2f（偏 %.0f%%）" % (
                    im["file"], im["src_aspect"], im["display_aspect"], im["distort_pct"])
                    for im in warped[:3]),
                "basis": "checklist §2.2 图片被拉伸（已扣除 srcRect 裁切）",
            })
        zoomed = [im for im in s["images"]
                  if im.get("zoom") and im["zoom"] > 1.05]
        if zoomed:
            findings.append({
                "level": "P2", "slide": idx, "item": "图片被放大，必然发虚",
                "detail": "; ".join("%s %dx%d → 显示 %sx%s（%.2fx）" % (
                    im["file"], im["native_px"][0], im["native_px"][1],
                    im["display_px"][0], im["display_px"][1], im["zoom"])
                    for im in sorted(zoomed, key=lambda x: -x["zoom"])[:3]),
                "basis": "checklist §2.2 分辨率不足",
            })
        small = [im for im in s["images"] if im.get("low_quality")]
        if small and not zoomed:
            findings.append({
                "level": "P2", "slide": idx, "item": "图片分辨率余量不足",
                "detail": "%d 张过采样不足 1.33x（显示尺寸 / 有效原图尺寸 > %.2f）：%s" % (
                    len(small), ZOOM_LOW,
                    "、".join("%s(%.2fx)" % (im["file"], im["zoom"]) for im in
                              sorted(small, key=lambda x: -x["zoom"])[:4])),
                "basis": "checklist §2.2 分辨率不足",
            })
        if s["text_chars"] and s["text_chars"] < 80 and idx > 1:
            findings.append({
                "level": "P1", "slide": idx, "item": "页面密度不足",
                "detail": "本页文字仅 %d 字（内容页下限 80 字；若为封面/过渡/结尾页可忽略）" % s["text_chars"],
                "basis": "checklist §2 页面密度不足",
            })
        too_small = [pt for pt in (s["font_sizes"] or {}) if float(pt) < 12]
        if too_small:
            findings.append({
                "level": "P1", "slide": idx, "item": "字号过小",
                "detail": "出现 %s pt 字号" % ", ".join(str(x) for x in sorted(too_small)),
                "basis": "checklist §1.2 最小字号",
            })
        if len(s["fonts"]) > 2:
            findings.append({
                "level": "P1", "slide": idx, "item": "单页字体过多",
                "detail": "%d 套字体：%s" % (len(s["fonts"]), ", ".join(s["fonts"])),
                "basis": "checklist §1.2 字体家族数",
            })
        if s.get("reserved_zone_conflicts"):
            zc = s["reserved_zone_conflicts"]
            big = [h for h in zc if h.get("zone_area", 0) >= RESERVED_BIG_ZONE]
            small = [h for h in zc if h.get("zone_area", 0) < RESERVED_BIG_ZONE]
            # 保留区框取自版式里装饰形状（如 logo 图片）的**声明包围盒**，
            # 而 logo 素材本身常自带透明/白色内边距 —— 内容落在那段空白里
            # 也会被算成交叠。实测 2026-09-16：10 页「压 LOGO」渲染复核后
            # **全部为误报**（脚注文字与 logo 之间有 8px 以上净空）。
            # 因此这里降级为**候选**并打 needs_render，不直接当结论。
            if big:
                findings.append({
                    "level": "P2", "slide": idx, "needs_render": True,
                    "item": "内容进入版式 LOGO / 大装饰保留区（候选，须渲染复核）",
                    "detail": "%d 个形状与 %s 的声明包围盒相交（最大 %.0f%%；例：%r）" % (
                        len(big), "、".join(sorted({h["zone"] for h in big})),
                        max(h["ratio"] for h in big) * 100,
                        big[0]["text"] or big[0]["shape"]),
                    "basis": "SKILL Step 3 留白区不可侵犯 / 陷阱「声明包围盒 ≠ 可见墨迹」",
                })
            if small:
                findings.append({
                    "level": "P2", "slide": idx, "needs_render": True,
                    "item": "内容与版式装饰条重叠（候选，须渲染复核）",
                    "detail": "%d 处与 %s 的声明包围盒相交（例：%r）" % (
                        len(small), "、".join(sorted({h["zone"] for h in small})),
                        small[0]["text"] or small[0]["shape"]),
                    "basis": "SKILL Step 3 留白区不可侵犯 / 陷阱「声明包围盒 ≠ 可见墨迹」",
                })
        if s.get("grid_overflow"):
            go = s["grid_overflow"]
            findings.append({
                "level": "P2", "slide": idx, "item": "内容与版式栅格不齐",
                "detail": "%d 处越出栅格：%s" % (len(go), "、".join(
                    "%s边界 %.0f vs 栅格 %.0f" % (h["side"], h["px"], h["grid"])
                    for h in go[:3])),
                "basis": "SKILL Step 3 对齐（换模板后不会自动跟随）",
            })
        if s.get("suspected_title_textbox") and not s.get("title_in_placeholder"):
            findings.append({
                "level": "P0", "slide": idx, "item": "标题未在标题占位符中",
                "detail": "疑似标题「%s」由自由文本框承载，换模板时会错位" % s["suspected_title_textbox"][:30],
                "basis": "SKILL Step 1.2 标题归位",
            })
        if s.get("hardcoded_fonts"):
            findings.append({
                "level": "P1", "slide": idx, "item": "硬编码字体",
                "detail": "%d 处 run 写了具体字体名，未继承主题字体" % s["hardcoded_fonts"],
                "basis": "SKILL Step 1.3 字体规范",
            })
        if s.get("color_count", 0) > MAX_PAGE_COLORS:
            non_neutral = [c for c in s["colors"] if not is_neutral(c)]
            findings.append({
                "level": "P1" if s["color_count"] <= 5 else "P0",
                "slide": idx, "item": "单页颜色超过三色",
                "detail": "有彩色 %d 种（上限 %d，黑白与中性灰不计入）：%s" % (
                    s["color_count"], MAX_PAGE_COLORS, ", ".join(non_neutral[:6])),
                "basis": "SKILL Step 4.1 单页三色规则",
            })

    if len(effective_colors) > 4:
        findings.append({
            "level": "P1", "slide": 0, "item": "全篇颜色超标",
            "detail": "有效非中性色 %d 种：%s" % (
                len(effective_colors),
                ", ".join("%s×%d" % (c, n) for c, n in sorted(
                    effective_colors.items(), key=lambda kv: -kv[1])[:12])),
            "basis": "checklist §1.1 颜色数量（≤ 4）",
        })
    if len(fonts_clean) > 2:
        findings.append({
            "level": "P1", "slide": 0, "item": "全篇字体超标",
            "detail": "%d 套字体：%s" % (len(fonts_clean), ", ".join(
                "%s×%d" % (f, n) for f, n in fonts_clean.most_common(8))),
            "basis": "checklist §1.2 字体家族数（≤ 2）",
        })
    if len(all_sizes) > 8:
        findings.append({
            "level": "P2", "slide": 0, "item": "字号层级过多",
            "detail": "%d 级字号，层级不成阶梯" % len(all_sizes),
            "basis": "checklist §1.2 字号层级",
        })

    return {
        "file": str(pptx_path),
        "slide_count": len(slides_out),
        "canvas_px": canvas_px,
        "canvas_ratio": safe(lambda: round(sw / sh, 3)),
        "theme_colors": theme_map,
        "global": {
            "fonts": dict(fonts_clean.most_common()),
            "font_sizes": dict(sorted(all_sizes.items(), key=lambda kv: -kv[1])),
            "colors_all": dict(all_colors.most_common()),
            "colors_effective": dict(sorted(effective_colors.items(), key=lambda kv: -kv[1])),
        },
        "slides": slides_out,
        "findings": findings,
    }


def render_md(data: dict) -> str:
    L = []
    L.append("# PPT 审计报告\n")
    L.append("| 项 | 值 |")
    L.append("|---|---|")
    L.append("| 文件 | `%s` |" % data["file"])
    L.append("| 页数 | %d |" % data["slide_count"])
    L.append("| 画布 | %s px（宽高比 %s） |" % (data["canvas_px"], data.get("canvas_ratio")))
    L.append("")

    L.append("## 母版色板（theme clrScheme）\n")
    if data["theme_colors"]:
        L.append("| 槽位 | 色值 |")
        L.append("|---|---|")
        for k, v in data["theme_colors"].items():
            L.append("| %s | `%s` |" % (k, v))
    else:
        L.append("_未解析到主题色板_")
    L.append("")

    g = data["global"]
    L.append("## 全局统计\n")
    L.append("**字体**（规范 ≤ 2 套）\n")
    L.append("| 字体 | 出现次数 |")
    L.append("|---|---|")
    for k, v in g["fonts"].items():
        L.append("| %s | %d |" % (k, v))
    L.append("")
    L.append("**字号**（共 %d 级）\n" % len(g["font_sizes"]))
    L.append("| pt | 出现次数 |")
    L.append("|---|---|")
    for k, v in g["font_sizes"].items():
        L.append("| %s | %d |" % (k, v))
    L.append("")
    L.append("**有效颜色**（排除中性色后 %d 种，规范 ≤ 4）\n" % len(g["colors_effective"]))
    L.append("| 色值 | 出现次数 |")
    L.append("|---|---|")
    for k, v in g["colors_effective"].items():
        L.append("| `%s` | %d |" % (k, v))
    L.append("")

    L.append("## 逐页概览\n")
    L.append("| # | 版式 | 标题状态 | 字数 | 形状 | 图 | 有彩色数 | 压保留区 | 问题 |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    for s in data["slides"]:
        issues = []
        if s["out_of_canvas"]:
            issues.append("出画布×%d" % len(s["out_of_canvas"]))
        if s["placeholder_leftovers"]:
            issues.append("占位符×%d" % len(s["placeholder_leftovers"]))
        if s["overlapping_text_pairs"]:
            issues.append("重叠×%d" % len(s["overlapping_text_pairs"]))
        if s.get("grid_overflow"):
            issues.append("越栅格×%d" % len(s["grid_overflow"]))
        if not issues:
            issues.append("-")
        if s.get("title_in_placeholder"):
            tstate = "占位符 OK"
        elif s.get("suspected_title_textbox"):
            tstate = "**文本框 ✗**"
        else:
            tstate = "无标题"
        rz = s.get("reserved_zone_conflicts") or []
        L.append("| %d | %s | %s | %d | %d | %d | %d | %s | %s |" % (
            s["index"], s["layout"] or "-", tstate, s["text_chars"], s["shape_count"],
            s["picture_count"], s.get("color_count", 0),
            ("**×%d**" % len(rz)) if rz else "-",
            " ".join(issues)))
    L.append("")

    L.append("## 自动判定问题清单\n")
    if data["findings"]:
        L.append("| 级别 | 页 | 问题 | 详情 | 判定依据 |")
        L.append("|---|---|---|---|---|")
        for f in sorted(data["findings"], key=lambda x: (x["level"], x["slide"])):
            L.append("| %s | %s | %s | %s | %s |" % (
                f["level"], f["slide"] or "全篇", f["item"], f["detail"], f["basis"]))
    else:
        L.append("_未发现可自动判定的问题_")
    L.append("")
    L.append("> 自动判定只覆盖可量化项。结构完整性、标题观点性、逻辑承接等需人工按 checklist 复核。")
    L.append("")

    L.append("## 装饰保留区冲突（版式上的 logo / 装饰，换模板后位置不变，内容压上去即冲突）\n")
    L.append("| 页 | 保留区 | 压入的形状 | 文字 | 覆盖 |")
    L.append("|---|---|---|---|---|")
    any_zone = False
    for s in data["slides"]:
        for h in (s.get("reserved_zone_conflicts") or []):
            any_zone = True
            L.append("| %d | %s | %s | %s | %.0f%% |" % (
                s["index"], h["zone"], h["shape"], (h["text"] or "-")[:36], h["ratio"] * 100))
    if not any_zone:
        L.append("| - | - | - | - | - |")
    L.append("")

    L.append("## 内容与版式栅格偏差（同一内容越出栅格的形状数）\n")
    L.append("| 页 | 左边界偏差 | 右边界偏差 |")
    L.append("|---|---|---|")
    for s in data["slides"]:
        go = s.get("grid_overflow") or []
        if not go:
            continue
        lefts = [h for h in go if h["side"] == "左"]
        rights = [h for h in go if h["side"] == "右"]
        L.append("| %d | %s | %s |" % (
            s["index"],
            ("%d 个（最左 %.0f，栅格 %.0f）" % (len(lefts), min(h["px"] for h in lefts), lefts[0]["grid"])) if lefts else "-",
            ("%d 个（最右 %.0f，栅格 %.0f）" % (len(rights), max(h["px"] for h in rights), rights[0]["grid"])) if rights else "-"))
    L.append("")

    all_images = [(s["index"], im) for s in data["slides"] for im in s.get("images", [])]
    if all_images:
        L.append("## 图片素材清单（重建时必须复用，禁止用生图替代真实数据图/产品截图）\n")
        L.append("| 页 | 文件 | 原始像素 | 显示尺寸 px | 位置 px | KB | 备注 |")
        L.append("|---|---|---|---|---|---|---|")
        for pidx, im in all_images:
            notes = []
            if im.get("blank"):
                notes.append("**空图（全透明）**")
            if im.get("distort_pct") is not None and im["distort_pct"] > 5:
                notes.append("**变形 %.0f%%**" % im["distort_pct"])
            if im.get("zoom") and im["zoom"] > 1.05:
                notes.append("**放大 %.2fx**" % im["zoom"])
            if im.get("in_group"):
                notes.append("组合内(已换算绝对)")
            if im.get("src_rect"):
                notes.append("已裁切")
            if im["low_quality"]:
                notes.append("余量不足 %.2fx" % im["zoom"])
            note = "、".join(notes)
            npx = im["native_px"]
            L.append("| %d | %s | %s×%s | %s×%s | %s,%s | %d | %s |" % (
                pidx, im["file"], npx[0], npx[1],
                im["display_px"][0], im["display_px"][1],
                im["pos_px"][0], im["pos_px"][1],
                round(im["bytes"] / 1024), note))
        L.append("")
        L.append("素材落盘位置：`audit/assets/`（仅 `--extract-images` 时产出）")
        L.append("")
    else:
        L.append("## 图片素材清单\n")
        L.append("_原件未检测到图片素材_")
        L.append("")

    L.append("## 逐页文字全文（重建时逐字沿用，禁止改写事实）\n")
    for s in data["slides"]:
        L.append("### 第 %d 页（%s）\n" % (s["index"], s["layout"] or "无版式名"))
        if s["texts"]:
            for t in s["texts"]:
                L.append("- %s" % t.replace("\n", " / "))
        else:
            L.append("_（本页无文字）_")
        L.append("")
    return "\n".join(L)


def main() -> int:
    ap = argparse.ArgumentParser(description="pptx 量化审计")
    ap.add_argument("pptx", help="待审计的 .pptx 路径")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--extract-images", action="store_true",
                    help="同时把原件内嵌图片导出到 <out>/assets/（不过滤小图）")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        print("ERROR: 文件不存在 %s" % src, file=sys.stderr)
        return 1
    if src.suffix.lower() == ".ppt":
        print("ERROR: .ppt 老格式不受支持，请先用 PowerPoint 另存为 .pptx", file=sys.stderr)
        return 1

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    data = audit(src, (out_dir / "assets") if args.extract_images else None)
    (out_dir / "audit.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "audit.md").write_text(render_md(data), encoding="utf-8")

    print("[audit] 页数 %d | 画布 %s px" % (data["slide_count"], data["canvas_px"]))
    print("[audit] 字体 %d 套 | 字号 %d 级 | 有效色 %d 种" % (
        len(data["global"]["fonts"]), len(data["global"]["font_sizes"]),
        len(data["global"]["colors_effective"])))
    print("[audit] 自动判定问题 %d 项" % len(data["findings"]))
    print("[audit] 产出 %s" % (out_dir / "audit.json"))
    print("[audit] 产出 %s" % (out_dir / "audit.md"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
