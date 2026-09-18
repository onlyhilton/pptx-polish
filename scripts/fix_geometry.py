#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fix_geometry.py — 按审计结论做**确定性**的几何修复（不改素材、不动别的元素）

两个子动作，都只作用于指定的一个形状：

  --fix-aspect   把图片的显示框宽高比改成与「有效原图区域」一致 → 消除变形
                 保持左上角不动；默认保持高度、只改宽度（--keep height），
                 这样下方有文字时不会压到。典型场景：1U 交换机图被横向拉宽 27.5%。

  --drop         删除一个形状（空图 / 编辑残留），并级联清理：
                   形状 → 该形状独占的 r:embed 关系 → 无人引用的媒体部件
                 → 若祖先组合因此变空，连空组合一起删

定位一律用 --slide N --shape-id ID —— 与 audit_deck.py 输出里的
p{slide}_s{id} 命名一一对应（报告里的 p22_s58 = 第 22 页 shape_id 58）。

为什么不用 python-pptx 保存：它会重写整包、丢掉未建模的部件。这里用
**字符串精确替换**，除目标部件的目标片段外，其余字节零变化，可逐像素比对。

用法：
    python fix_geometry.py deck.pptx --slide 22 --shape-id 58 --fix-aspect -o fixed.pptx
    python fix_geometry.py deck.pptx --slide 25 --shape-id 96 --drop -o fixed.pptx
    python fix_geometry.py deck.pptx --slide 25 --shape-id 96 --drop --dry-run

修复后务必：verify_pptx.py（P0 必须为 0）+ render_deck.py 渲染复核。
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck_rules import aspect_fix, effective_area, parse_src_rect  # noqa: E402

EMU_PER_PX = 9525
SHAPE_TAGS = ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp")
LEAF_TAGS = ("p:sp", "p:pic", "p:graphicFrame", "p:cxnSp")


def emu2px(v):
    return round(v / EMU_PER_PX, 1)


def px2emu(v):
    return int(round(v * EMU_PER_PX))


def cnpv_id(seg):
    m = re.search(r'<p:cNvPr id="(\d+)"', seg)
    return int(m.group(1)) if m else None


def cnpv_name(seg):
    m = re.search(r'<p:cNvPr[^>]*name="([^"]*)"', seg)
    return m.group(1) if m else ""


# --------------------------------------------------------------------------
# 包读写
# --------------------------------------------------------------------------
def read_pkg(path):
    z = zipfile.ZipFile(path)
    order = z.namelist()
    items = {n: z.read(n) for n in order}
    z.close()
    return items, order


def write_pkg(items, order, out):
    """逐部件写出，保持原有顺序；未改动的部件字节不变（writestr 是压缩存储，
    所以压缩流会重算，但部件内容逐字节相同）。"""
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for n in order:
            if n in items:
                z.writestr(n, items[n])


def slide_part(idx):
    return "ppt/slides/slide%d.xml" % idx


def rels_part(slide_path):
    d, f = os.path.split(slide_path)
    return "%s/_rels/%s.rels" % (d, f)


# --------------------------------------------------------------------------
# 形状定位：字符串级，标签平衡扫描（能正确处理嵌套组合）
# --------------------------------------------------------------------------
def spans_of(xml, tag):
    """返回 <tag>…</tag> 的所有 (start, end)，正确跳过嵌套的同名标签。"""
    open_re = re.compile(r"<%s(?:\s[^>]*)?>" % re.escape(tag))
    close_re = re.compile(r"</%s>" % re.escape(tag))
    out = []
    for mo in open_re.finditer(xml):
        pos, depth, end = mo.end(), 1, None
        while depth:
            a = open_re.search(xml, pos)
            b = close_re.search(xml, pos)
            if b is None:
                break
            if a is not None and a.start() < b.start():
                depth += 1
                pos = a.end()
            else:
                depth -= 1
                pos = b.end()
                if depth == 0:
                    end = pos
        if end is not None:
            out.append((mo.start(), end))
    return out


def find_shape(xml, shape_id):
    """找 shape_id 对应的形状元素：返回 (tag, start, end, seg)。

    每个候选段只认**段内第一个 cNvPr 的 id** —— 那必定是元素自己的 id，
    所以嵌套组合不会被误判。
    """
    for tag in SHAPE_TAGS:
        for st, en in spans_of(xml, tag):
            seg = xml[st:en]
            if cnpv_id(seg) == shape_id:
                return tag, st, en, seg
    return None


def group_chain(xml, shape_id):
    """返回包住目标形状的所有祖先组合 [(start, end, seg)]，从外到内，附带累积缩放。"""
    outer = []
    for st, en in spans_of(xml, "p:grpSp"):
        seg = xml[st:en]
        if re.search(r'<p:cNvPr id="%d"' % shape_id, seg):
            outer.append((st, en, seg))
    outer.sort(key=lambda x: x[1] - x[0], reverse=True)     # 越大越外 → 从外到内

    sx = sy = 1.0
    for _st, _en, seg in outer:
        xf = re.search(
            r"<p:grpSpPr>\s*<a:xfrm>.*?<a:ext cx=\"(\d+)\" cy=\"(\d+)\"\s*/>"
            r".*?<a:chExt cx=\"(\d+)\" cy=\"(\d+)\"\s*/>", seg, re.S)
        if xf:
            cx, cy, chx, chy = (int(g) for g in xf.groups())
            if chx:
                sx *= cx / chx
            if chy:
                sy *= cy / chy
    return outer, sx, sy


def strip_emptied_ancestors(xml, anc_ids):
    """反复删掉"祖先链上、已不含任何形状子元素"的组合。返回 (新 xml, 删了几个)。"""
    anc_ids = [i for i in anc_ids if i is not None]
    removed = 0
    while True:
        hit = None
        for st, en in sorted(spans_of(xml, "p:grpSp"), key=lambda x: x[1] - x[0]):
            seg = xml[st:en]                     # 从最内层开始试
            if cnpv_id(seg) not in anc_ids:
                continue
            if any(re.search(r"<%s(?:\s[^>]*)?>" % re.escape(t), seg) for t in LEAF_TAGS):
                continue                          # 还有叶子形状 → 不空
            if len(re.findall(r"<p:grpSp(?:\s[^>]*)?>", seg)) > 1:
                continue                          # 还含子组合 → 等下一轮
            hit = (st, en)
            break
        if hit is None:
            return xml, removed
        xml = xml[:hit[0]] + xml[hit[1]:]
        removed += 1


# --------------------------------------------------------------------------
# 关系与媒体清理
# --------------------------------------------------------------------------
def target_refs(items):
    """全包 .rels 里的媒体部件引用计数 → {归一化路径: [引用它的 rels 部件]}"""
    ref = {}
    for n, blob in items.items():
        if not n.endswith(".rels"):
            continue
        try:
            txt = blob.decode("utf-8")
        except Exception:
            continue
        base = os.path.dirname(os.path.dirname(n))
        for m in re.finditer(r'Target="([^"]+)"', txt):
            t = m.group(1)
            if t.startswith("http"):
                continue
            full = os.path.normpath(os.path.join(base, t)).replace("\\", "/")
            ref.setdefault(full, []).append(n)
    return ref


def media_of(items, sl_path, rid):
    """rId → 归一化媒体部件路径。"""
    rp = rels_part(sl_path)
    if rp not in items:
        return None
    txt = items[rp].decode("utf-8")
    m = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % re.escape(rid), txt)
    if not m:
        return None
    return os.path.normpath(
        os.path.join(os.path.dirname(sl_path), m.group(1))).replace("\\", "/")


def drop_relationship(items, sl_path, rid):
    rp = rels_part(sl_path)
    if rp not in items:
        return False
    txt = items[rp].decode("utf-8")
    m = re.search(r'<Relationship[^>]*Id="%s"[^>]*/>' % re.escape(rid), txt)
    if not m:
        return False
    items[rp] = (txt[:m.start()] + txt[m.end():]).encode("utf-8")
    return True


# --------------------------------------------------------------------------
# 子动作
# --------------------------------------------------------------------------
def do_fix_aspect(items, sl_path, shape_id, keep, dry):
    xml = items[sl_path].decode("utf-8")
    found = find_shape(xml, shape_id)
    if not found:
        raise SystemExit("ERROR: %s 里找不到 shape_id=%d" % (sl_path, shape_id))
    tag, st, en, seg = found
    if tag != "p:pic":
        raise SystemExit("ERROR: shape_id=%d 是 <%s>，不是图片；--fix-aspect 只作用于图片"
                         % (shape_id, tag))

    emb = re.search(r'r:embed="([^"]+)"', seg)
    if not emb:
        raise SystemExit("ERROR: 该图片没有 r:embed，无法确定原生尺寸")
    rid = emb.group(1)
    media = media_of(items, sl_path, rid)
    if not media or media not in items:
        raise SystemExit("ERROR: 媒体部件不存在（%s）" % media)

    from PIL import Image
    with Image.open(io.BytesIO(items[media])) as im:
        pw, ph = im.size

    src_rect = parse_src_rect(seg)
    xf = re.search(r"<a:xfrm>\s*<a:off x=\"(-?\d+)\" y=\"(-?\d+)\"\s*/>"
                   r"\s*<a:ext cx=\"(\d+)\" cy=\"(\d+)\"\s*/>", seg)
    if not xf:
        raise SystemExit("ERROR: 该图片没有 a:xfrm/off/ext，无法修复")
    off_x, off_y, ext_cx, ext_cy = (int(g) for g in xf.groups())

    chain, kx, ky = group_chain(xml, shape_id)
    in_group = bool(chain)
    disp = (ext_cx * kx / EMU_PER_PX, ext_cy * ky / EMU_PER_PX)

    info = aspect_fix(src_rect, (pw, ph), disp, keep=keep)
    if info["distort_pct"] is None:
        raise SystemExit("ERROR: 无法计算变形（缺 srcRect 或尺寸异常）")
    eff_w, eff_h = effective_area(src_rect, pw, ph)
    tw, th = info["target_px"]
    # 保持不变的那一维**直接沿用原 EMU**，不经过 px→EMU 往返（否则会少几十个 EMU）
    if keep == "width":
        new_cx, new_cy = ext_cx, px2emu(th / ky)
    else:
        new_cx, new_cy = px2emu(tw / kx), ext_cy

    print("P%s  shape_id=%d  「%s」" % (sl_path.rsplit("slide", 1)[-1].split(".")[0],
                                        shape_id, cnpv_name(seg)))
    print("  媒体        %s  原生 %d×%d" % (media, pw, ph))
    print("  srcRect     %s → 有效区域 %.1f×%.1f px（比例 %.3f）"
          % (src_rect or "无", eff_w, eff_h, info["src_aspect"]))
    print("  显示框      当前 %.1f×%.1f px（比例 %.3f）%s"
          % (disp[0], disp[1], info["display_aspect"],
             "，在组合内（累积缩放 %.3f / %.3f）" % (kx, ky) if in_group else ""))
    print("  变形        %.1f%% → 0.0%%" % info["distort_pct"])
    print("  过采样倍率  %.2fx → %.2fx" % (info["zoom"] or 0, info["zoom_after"] or 0))
    print("  修正后      %.1f×%.1f px（保持 %s 不变）" % (tw, th, keep))
    print("  位置        左上角不动 (%.1f, %.1f)；右下角 x %.1f → %.1f"
          % (emu2px(off_x), emu2px(off_y), emu2px(off_x) + disp[0], emu2px(off_x) + tw))

    if info["distort_pct"] <= 0.05:
        print("  ⚠ 本来就没有变形，无需修改")
        return False
    if info["zoom_after"] and info["zoom_after"] > 0.75:
        print("  ⚠ 改后过采样 %.2fx 偏紧，确认是否要同时换更高清素材" % info["zoom_after"])
    if dry:
        print("  [dry-run] 未写入")
        return False

    new_xf = ('<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
              % (off_x, off_y, new_cx, new_cy))
    seg2 = seg.replace(xf.group(0), new_xf, 1)
    if seg2 == seg:
        raise SystemExit("ERROR: a:xfrm 替换失败")
    items[sl_path] = (xml[:st] + seg2 + xml[en:]).encode("utf-8")
    print("  已写入      a:ext cx %d → %d EMU（%d 字节片段外零变化）"
          % (ext_cx, new_cx, len(seg)))
    return True


def do_drop(items, order, sl_path, shape_id, dry):
    xml = items[sl_path].decode("utf-8")
    found = find_shape(xml, shape_id)
    if not found:
        raise SystemExit("ERROR: %s 里找不到 shape_id=%d" % (sl_path, shape_id))
    tag, st, en, seg = found

    emb = re.search(r'r:embed="([^"]+)"', seg)
    rid = emb.group(1) if emb else None
    media = media_of(items, sl_path, rid) if rid else None
    if media and media not in items:
        media = None

    rest = xml[:st] + xml[en:]
    rid_left = len(re.findall(r'r:(?:embed|link)="%s"' % re.escape(rid), rest)) if rid else 0
    others = []
    if media:
        ref = target_refs(items)
        others = [r for r in ref.get(media, []) if r != rels_part(sl_path)]

    chain, _kx, _ky = group_chain(xml, shape_id)
    anc_ids = [cnpv_id(s) for _a, _b, s in chain]
    new_xml, n_grp = strip_emptied_ancestors(rest, anc_ids)

    print("P%s  shape_id=%d  「%s」  <%s>"
          % (sl_path.rsplit("slide", 1)[-1].split(".")[0], shape_id, cnpv_name(seg), tag))
    print("  删除        形状本身（%d 字节 XML）" % len(seg))
    if rid:
        print("  关系        %s → %s%s" % (
            rid, media or "?",
            "（本幻灯片仅此一处引用 → 一并删）" if rid_left == 0 else "（仍有引用 → 保留）"))
    if media:
        print("  媒体部件    %s%s" % (
            media, "（无其它引用 → 一并删）" if not others
            else "（仍被引用 → 保留）：%s" % "、".join(others)))
    if n_grp:
        print("  祖先组合    删后变空 → 一并删除 %d 个" % n_grp)
    if tag != "p:pic":
        print("  注意        目标不是图片，请确认没有连带语义")
    if dry:
        print("  [dry-run] 未写入")
        return False

    items[sl_path] = new_xml.encode("utf-8")
    if rid and rid_left == 0 and drop_relationship(items, sl_path, rid):
        print("  已删关系    %s" % rid)
    if media and not others and media in items:
        del items[media]
        order.remove(media)
        print("  已删部件    %s" % media)
    print("  已写入      slide XML 缩短 %d 字节" % (len(xml) - len(new_xml)))
    return True


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="按审计结论做确定性的几何修复（消变形 / 删残留形状）")
    ap.add_argument("pptx")
    ap.add_argument("--slide", type=int, required=True, help="页码，1-based")
    ap.add_argument("--shape-id", type=int, required=True,
                    help="形状 id（cNvPr/@id，即审计报告 p{slide}_s{id} 里的数字）")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--fix-aspect", action="store_true",
                   help="把显示框宽高比改成与有效原图区域一致（消除变形）")
    g.add_argument("--drop", action="store_true",
                   help="删除该形状，并清理独占关系与孤儿媒体部件")
    ap.add_argument("--keep", choices=["height", "width"], default="height",
                    help="--fix-aspect 保持哪一边不变（默认 height：只改宽，不压下方文字）")
    ap.add_argument("-o", "--out", help="输出路径（默认原名加 .fixed）")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不写文件")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        raise SystemExit("ERROR: 文件不存在 %s" % src)
    if src.suffix.lower() != ".pptx":
        raise SystemExit("ERROR: 只处理 .pptx（.ppt 请先用 PowerPoint 另存为 .pptx）")

    items, order = read_pkg(str(src))
    sl = slide_part(args.slide)
    if sl not in items:
        raise SystemExit("ERROR: 包内没有 %s" % sl)

    if args.fix_aspect:
        changed = do_fix_aspect(items, sl, args.shape_id, args.keep, args.dry_run)
    else:
        changed = do_drop(items, order, sl, args.shape_id, args.dry_run)

    if args.dry_run or not changed:
        return

    import xml.etree.ElementTree as ET
    try:
        ET.fromstring(items[sl].decode("utf-8"))
    except Exception as e:
        raise SystemExit("ERROR: 修改后 XML 不合法：%s" % e)

    out = args.out or (str(src.with_suffix("")) + ".fixed.pptx")
    write_pkg(items, order, out)
    print("→ %s  (%.2f MB)" % (out, os.path.getsize(out) / 1048576))
    print("  下一步：verify_pptx.py 复查 + render_deck.py 渲染复核")


if __name__ == "__main__":
    main()
