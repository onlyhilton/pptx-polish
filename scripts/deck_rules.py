# -*- coding: utf-8 -*-
"""共享判定规则（唯一事实源）—— 标题判定 + 图片几何判定

背景：审计脚本与规范化脚本原先各自实现了一套"什么算标题"的判定，
      结果一次阈值修正只改了一边 —— audit 用 60 字符、normalize 用 40，
      导致 46 字符的标题在审计里被报为"未归位"、在规范化里却被静默跳过。
      同一判定必须只有一处实现。图片几何（变形 / 清晰度）同理，
      audit_deck 与 fix_geometry 必须共用本模块。

标题判定顺序（先过滤，再排序）：
  1. 非图片、非占位符、有文本
  2. 文本长度 ≤ TITLE_MAX_CHARS
  3. 顶部位置 ≤ 区域上限（内容版式 TITLE_ZONE_PX / 封面版式 COVER_ZONE_PX）
  4. 字号不足 MIN_TITLE_PT 时：封面不接受；内容页只认紧贴顶部的（≤ TOP_STRICT_PX）
  5. 排序：先按高度带（TITLE_TOP_BAND）分组，带内比字号，带间取更靠上的

这样"标题"的定义在两个脚本里完全一致，不会再出现一边报、一边跳。
"""
from __future__ import annotations

import re

PX = 9525                      # EMU per pixel（96 DPI）

TITLE_ZONE_PX = 200            # 内容版式：标题候选的垂直上限
COVER_ZONE_PX = 660            # 封面版式：标题可落在中下部，放宽到近整页
TITLE_MAX_CHARS = 60           # 标题候选的最大字数
                               # （实测 46 字符的长标题曾被 40 的旧阈值误过滤）
MIN_TITLE_PT = 18              # 低于此字号不视为标题
TOP_STRICT_PX = 100            # 字号无从判断时，内容页只认此高度内的文本框
TITLE_TOP_BAND = 24            # 同一"高度带"内才比字号；标题必然在该带最上面


def emu_to_px(v) -> float:
    return (v or 0) / PX


def max_font_pt(text_frame) -> float:
    """文本帧内出现的最大字号（pt）。显式字号取最大；全为继承时返回 0。"""
    best = 0.0
    try:
        for para in text_frame.paragraphs:
            for run in para.runs:
                sz = run.font.size
                if sz is not None:
                    best = max(best, sz.pt)
            # 段落级默认字号
            if para.font.size is not None:
                best = max(best, para.font.size.pt)
    except Exception:
        pass
    return best


def collect_title_candidates(shapes, zone: str = "top"):
    """收集疑似标题的自由文本框。

    zone='top'  内容版式：只在上部区域找
    zone='full' 封面版式：全页找，取字号最大的短文本
    zone=None   该版式无标题槽位：不归位

    返回 [(pt, top_px, text, shape), ...]，未排序。
    """
    if zone is None:
        return []
    limit = COVER_ZONE_PX if zone == "full" else TITLE_ZONE_PX

    out = []
    for shp in shapes:
        try:
            if shp.shape_type == 13:            # PICTURE
                continue
        except Exception:
            pass
        try:
            if shp.is_placeholder:
                continue                        # 占位符不参与
        except Exception:
            pass
        if not getattr(shp, "has_text_frame", False):
            continue
        text = (shp.text_frame.text or "").strip()
        if not text or len(text) > TITLE_MAX_CHARS:
            continue
        top_px = emu_to_px(shp.top)
        if top_px > limit:
            continue
        pt = max_font_pt(shp.text_frame)
        if pt < MIN_TITLE_PT:
            # 字号无从判断：封面不接受（封面标题必有显式大字号）；
            # 内容页只认紧贴顶部的
            if zone == "full" or top_px > TOP_STRICT_PX:
                continue
        out.append((pt, top_px, text, shp))
    return out


def pick_title(candidates):
    """从候选中选标题。

    先取最靠上的高度带（top_min ~ top_min + TITLE_TOP_BAND），再在带内取字号最大者。
    标题必然在同一高度带里字号最大，而不会选中页面中部的"大号正文字"，
    也不会因为某个更靠上的小字号装饰文字而误判。
    """
    if not candidates:
        return None
    top_min = min(c[1] for c in candidates)
    band = [c for c in candidates if c[1] <= top_min + TITLE_TOP_BAND]
    band.sort(key=lambda c: (-c[0], c[1]))
    return band[0]


def find_title_textbox(shapes, zone: str = "top", has_title_ph_with_text: bool = False):
    """一步到位：返回 (pt, top_px, text, shape) 或 None。"""
    if has_title_ph_with_text:
        return None
    return pick_title(collect_title_candidates(shapes, zone))


def title_zone_for(slide):
    """依据当前版式提供的标题槽位，决定标题搜索策略。

    'full' = 封面类（ctrTitle，标题可能在中下部）
    'top'  = 内容类（title，标题在顶部）
    None   = 该版式没有标题槽位（封底 / 空白页），不做归位
    """
    from pptx.enum.shapes import PP_PLACEHOLDER
    has_ctr = has_title = False
    try:
        for ph in slide.slide_layout.placeholders:
            try:
                t = ph.placeholder_format.type
            except Exception:
                continue
            if t == PP_PLACEHOLDER.CENTER_TITLE:
                has_ctr = True
            elif t == PP_PLACEHOLDER.TITLE:
                has_title = True
    except Exception:
        pass
    if has_ctr:
        return "full"
    if has_title:
        return "top"
    return None


# --------------------------------------------------------------------------
# 图片几何判定（audit_deck 与 fix_geometry 的唯一事实源）
# --------------------------------------------------------------------------
#
# 两条判据都必须区分"设计意图"与"真问题"，否则会大量误报：
#
#   1. 变形：图片用 srcRect 裁切填充时，显示框比例 ≠ 原图比例是**设计意图**。
#      只有"有效原图区域"（扣除 srcRect 裁掉的部分）的比例与显示框比例不符，
#      才是真被拉伸。不读 srcRect 会把所有非等比显示都报成变形。
#
#   2. 清晰度：只取决于**过采样倍率** = 显示尺寸 / 有效原图区域。
#      绝对像素数说明不了任何事 —— 333px 的图缩到 145px 显示（2.3x 过采样）
#      是清晰的，401px 的图拉到 455px 显示（1.13x）才是真的糊。

SRC_RECT_UNITS = 100000.0      # srcRect 的 l/t/r/b 单位：千分比（100000 = 100%）
DISTORT_TOL_PCT = 5.0          # 变形容差：有效区域比例与显示框比例偏差 > 5% 才算
ZOOM_TIGHT = 0.75              # 过采样倍率 < 0.75（余量不足 1.33x）算"分辨率余量不足"


def parse_src_rect(el):
    """读 <a:srcRect> 的 l/t/r/b，缺省 0；无裁切时返回 None。

    el 可以是：XML 元素 / 已解析好的 dict / **XML 片段字符串**。
    支持字符串是因为 fix_geometry.py 用字符串级定位（不重排整份 XML），
    拿到的是片段而不是元素。
    """
    if el is None:
        return None
    if isinstance(el, str):
        m = re.search(r"<a:srcRect\b([^>]*?)/?>", el)
        if not m:
            return None
        attrs = dict(re.findall(r'([a-z]+)="(-?\d+)"', m.group(1)))
        if not attrs:
            return None
        return {k: int(attrs.get(k, 0)) for k in ("l", "t", "r", "b")}
    if isinstance(el, dict):
        return {k: int(el.get(k) or 0) for k in ("l", "t", "r", "b")}
    try:
        return {k: int(el.get(k) or 0) for k in ("l", "t", "r", "b")}
    except Exception:
        return None


def effective_area(src_rect, pw, ph):
    """srcRect 裁切之后真正可见的原图区域（像素）。"""
    l = (src_rect or {}).get("l", 0) / SRC_RECT_UNITS
    r = (src_rect or {}).get("r", 0) / SRC_RECT_UNITS
    t = (src_rect or {}).get("t", 0) / SRC_RECT_UNITS
    b = (src_rect or {}).get("b", 0) / SRC_RECT_UNITS
    return max(pw * (1 - l - r), 1.0), max(ph * (1 - t - b), 1.0)


def aspect_fix(src_rect, native_px, display_px, keep="height"):
    """算出"消除变形"需要的显示框尺寸。

    keep='height' 保持高度、只改宽度（默认；下方有文字时不会压到）
    keep='width'  保持宽度、只改高度

    返回 dict：target_px 目标显示尺寸、src_aspect、display_aspect、
              distort_pct 当前变形百分比、zoom 当前过采样倍率、
              zoom_after 改后过采样倍率（改大显示框会让它变紧）。
    """
    pw, ph = native_px
    dw, dh = display_px
    eff_w, eff_h = effective_area(src_rect, pw, ph)
    src_ar = eff_w / eff_h
    disp_ar = (dw / dh) if (dw and dh) else None
    distort = (abs(disp_ar - src_ar) / src_ar * 100) if disp_ar else None
    zoom = max(dw / eff_w, dh / eff_h) if (dw and dh) else None
    if keep == "width":
        tw, th = dw, dw / src_ar
    else:
        tw, th = dh * src_ar, dh
    zoom_after = max(tw / eff_w, th / eff_h) if (tw and th) else None
    return {
        "keep": keep,
        # 不做舍入 —— 调用方要拿它换回 EMU，先舍到 0.1px 会让"保持不变"的那一维
        # 少几十个 EMU（实测 43.7 vs 43.7463 px → 高度少了 439 EMU）。
        "target_px": (tw, th),
        "src_aspect": round(src_ar, 3),
        "display_aspect": round(disp_ar, 3) if disp_ar else None,
        "distort_pct": round(distort, 1) if distort is not None else None,
        "zoom": round(zoom, 2) if zoom else None,
        "zoom_after": round(zoom_after, 2) if zoom_after else None,
    }


def image_geometry(src_rect, native_px, display_px):
    """audit 用的那三个数：src_aspect / display_aspect / distort_pct / zoom。"""
    r = aspect_fix(src_rect, native_px, display_px)
    return {
        "src_aspect": r["src_aspect"],
        "display_aspect": r["display_aspect"],
        "distort_pct": r["distort_pct"],
        "zoom": r["zoom"],
    }
