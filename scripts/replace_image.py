# -*- coding: utf-8 -*-
"""
replace_image.py —— 用高清素材替换 pptx 里的低清图片，版面几何零变化。

原理
----
图片在页面上的位置/尺寸由 `<a:xfrm>` 决定，和图片本身多少像素无关。
所以"图糊了"的正确修法不是去改版面，而是**只换 ppt/media 里的图像数据**：

    图片框（off / ext / srcRect）一个字节不动  →  渲染尺寸与位置必然不变
    换上一张更大的位图                        →  清晰度提升

脚本把新素材放进一张"与旧图同构"的画布（画布比例、内容占比都照抄旧图），
因此渲染出来的**内容足迹**与旧图完全重合。唯一变量是清晰度。

内容足迹 vs 产品本体
--------------------
"内容 bbox"包含柔和阴影/光晕。素材如果是**紧贴阴影裁切**的（圆盘几乎占满画布），
直接按内容 bbox 对位会让产品本体显得大一圈（实测 +7%）。
这时用 `--scale 0.93` 左右微调到产品本体等大。

用法
----
    # 按媒体部件名定位
    python replace_image.py deck.pptx --part image89.png --new hi.webp --out new.pptx

    # 按"第几页第几张图"定位
    python replace_image.py deck.pptx --slide 23 --pic 10 --new hi.webp --out new.pptx

    # 按**型号**取素材（走图库索引，不用写路径）—— 推荐
    python replace_image.py deck.pptx --slide 23 --pic 10 --new-model RG-RAP72 --out new.pptx

    # 批量按型号换：媒体部件=型号
    python replace_image.py deck.pptx --map-model "image89.png=RG-RAP72,image90.png=RG-NBS3100-24GT4SFP-P" --out new.pptx

    # 只出预览不写包
    python replace_image.py deck.pptx --slide 23 --pic 10 --new-model RG-RAP72 --dry-run

    # 微调内容占比（紧贴裁切的产品图常用）
    python replace_image.py ... --scale 0.93 --canvas-scale 16

    # 批量：一次换多张（写路径）
    python replace_image.py deck.pptx --map "image89.png=hi_rap72.webp,image90.png=hi_sw.png" --out new.pptx

型号怎么取到素材
----------------
`--new-model` 走 `product_lib.py` 的图库索引（默认 `assets/product_index.json`）：
优先透明底 webp（体积小、PPT 首选），若相对显示尺寸过采样不足 1.5x 则自动升级到 PNG。
用 `--lib webp|png` 强制指定，`--variant front|side` 选视角。
索引不存在时会自动扫描图库重建（可用 `--root` 指定图库位置）。

输出
----
    <out>.pptx              替换后的演示文稿（未给 --out 时写到 <deck>-换图.pptx）
    <out>-换图报告.md        几何对照表 + 尺寸变化，供交付留档
"""
from __future__ import annotations

import argparse
import io
import os
import re
import shutil
import sys
import zipfile

import numpy as np
from PIL import Image
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import product_lib                                            # noqa: E402

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
EMU_PER_PX = 9525

WHITE = 248      # 近白判为空白
ALPHA_MIN = 8    # alpha 低于此值判为透明
DISC_ALPHA = 240
DISC_WHITE = 235

# 成图画布总像素上限。PIL 的 decompression-bomb 闸门在 89MP，
# 留足余量取 40MP —— 超过这个数画布只是白白占内存，对最终显示毫无收益。
MAX_CANVAS_PX = 40_000_000


# ---------------------------------------------------------------- 工具


def content_bbox(im, white=WHITE, alpha_min=ALPHA_MIN):
    """返回 (x0, y0, x1, y1)；全透明/纯白时回退到整幅。"""
    a = np.array(im.convert("RGBA")).astype(np.int16)
    m = (a[:, :, 3] > alpha_min) & (a[:, :, :3].min(axis=2) < white)
    ys, xs = np.nonzero(m)
    if len(xs) == 0:
        return 0, 0, im.size[0] - 1, im.size[1] - 1
    return int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())


def body_bbox(im):
    """"产品本体"（高不透明浓度的实体）bbox —— 用于 --match body。"""
    a = np.array(im.convert("RGBA")).astype(np.int16)
    m = (a[:, :, 3] > DISC_ALPHA) & (a[:, :, :3].min(axis=2) > DISC_WHITE)
    ys, xs = np.nonzero(m)
    if len(xs) == 0:
        return None
    return int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())


def premul_resize(im, size):
    """预乘 alpha 再缩放、再反预乘 —— 避免透明区脏 RGB 在缩放时渗成灰边。"""
    a = np.array(im.convert("RGBA")).astype(np.float32)
    alp = a[:, :, 3:4] / 255.0
    pm = np.concatenate([a[:, :, :3] * alp, a[:, :, 3:4]], axis=2)
    pi = Image.fromarray(pm.astype(np.uint8), "RGBA").resize(size, Image.LANCZOS)
    b = np.array(pi).astype(np.float32)
    al2 = b[:, :, 3:4] / 255.0
    rgb = np.where(al2 > 0, b[:, :, :3] / np.maximum(al2, 1e-6), 0)
    return Image.fromarray(
        np.concatenate([np.clip(rgb, 0, 255), b[:, :, 3:4]], axis=2).astype(np.uint8), "RGBA")


def encode(im, ext):
    """按目标部件的扩展名编码；jpg 需拍平 alpha。"""
    buf = io.BytesIO()
    if ext in ("jpg", "jpeg"):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[3])
        bg.save(buf, "JPEG", quality=92, optimize=True)
    else:
        im.save(buf, "PNG", optimize=True)
    return buf.getvalue()


# ---------------------------------------------------------------- 定位


def slide_pics(z, slide_no):
    """返回第 slide_no 页上的图片列表：[{rid, part, name, box, src_rect}]（文档顺序）。"""
    sp = "ppt/slides/slide%d.xml" % slide_no
    rels = "ppt/slides/_rels/slide%d.xml.rels" % slide_no
    if sp not in z.namelist():
        raise SystemExit("找不到 %s" % sp)
    rmap = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read(rels))}
    root = etree.fromstring(z.read(sp))
    out = []
    for pic in root.iter("{%s}pic" % P_NS):
        blip = pic.find(".//{%s}blip" % A_NS)
        if blip is None:
            continue
        rid = blip.get("{%s}embed" % R_NS)
        if not rid or rid not in rmap:
            continue
        part = os.path.normpath(os.path.join("ppt/slides", rmap[rid])).replace("\\", "/")
        xfrm = pic.find(".//{%s}xfrm" % A_NS)
        box = None
        if xfrm is not None:
            off, ext = xfrm.find("{%s}off" % A_NS), xfrm.find("{%s}ext" % A_NS)
            if off is not None and ext is not None:
                box = (int(off.get("x")) // EMU_PER_PX, int(off.get("y")) // EMU_PER_PX,
                       int(ext.get("cx")) // EMU_PER_PX, int(ext.get("cy")) // EMU_PER_PX)
        sr = pic.find(".//{%s}srcRect" % A_NS)
        src_rect = None
        if sr is not None:
            src_rect = {k: int(sr.get(k) or 0) for k in ("l", "t", "r", "b")}
        nv = pic.find(".//{%s}cNvPr" % P_NS)
        out.append({"rid": rid, "part": part, "name": nv.get("name") if nv is not None else "",
                    "box": box, "src_rect": src_rect})
    return out


def resolve(z, args):
    """把 --part / --slide+--pic 解析成 [(part, box, src_rect, label)]。"""
    jobs = []
    if args.part:
        for p in args.part.split(","):
            p = p.strip()
            part = "ppt/media/%s" % os.path.basename(p)
            if part not in z.namelist():
                raise SystemExit("包内没有 %s" % part)
            jobs.append({"part": part, "box": None, "src_rect": None, "label": os.path.basename(p)})
    if args.slide:
        pics = slide_pics(z, args.slide)
        sel = pics if not args.pic else [pics[args.pic - 1]]
        for p in sel:
            if p["part"] not in z.namelist():
                continue
            jobs.append({"part": p["part"], "box": p["box"], "src_rect": p["src_rect"],
                         "label": "P%d %s (%s)" % (args.slide, p["name"], os.path.basename(p["part"]))})
    for spec in (args.map or "").split(","):
        if "=" not in spec:
            continue
        old, new = spec.split("=", 1)
        part = "ppt/media/%s" % os.path.basename(old.strip())
        if part not in z.namelist():
            raise SystemExit("包内没有 %s" % part)
        jobs.append({"part": part, "box": None, "src_rect": None, "label": os.path.basename(part),
                     "asset": new.strip()})
    for spec in (getattr(args, "map_model", None) or "").split(","):
        if "=" not in spec:
            continue
        old, model = spec.split("=", 1)
        part = "ppt/media/%s" % os.path.basename(old.strip())
        if part not in z.namelist():
            raise SystemExit("包内没有 %s" % part)
        jobs.append({"part": part, "box": None, "src_rect": None, "label": os.path.basename(part),
                     "model": model.strip()})
    return jobs


def part_boxes(z):
    """扫全篇，返回 {媒体部件: 首个使用它的显示框 (x,y,w,h)}。"""
    out = {}
    for n in z.namelist():
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
        if not m:
            continue
        for p in slide_pics(z, int(m.group(1))):
            if p["part"] not in out and p["box"]:
                out[p["part"]] = p["box"]
    return out


def resolve_models(jobs, args, z):
    """把 job 里的 `model` 换成实际素材路径（走图库索引）。"""
    pending = [j for j in jobs if j.get("model") and not j.get("asset")]
    if not pending:
        return
    idx = product_lib.load(args.index, args.root)
    boxes = None
    for j in pending:
        if not j.get("box"):
            if boxes is None:
                boxes = part_boxes(z)
            j["box"] = boxes.get(j["part"])
        rec, note = product_lib.pick_for_box(
            j["model"], idx, j.get("box"), prefer=args.lib, variant=args.variant,
            min_oversample=args.min_oversample)
        if not rec:
            raise SystemExit("图库里找不到型号 %s（%s）" % (j["model"], note))
        if not os.path.exists(rec["path"]):
            raise SystemExit("索引里的文件不在了（图库挪动过？重跑 --build）：%s" % rec["path"])
        j["asset"] = rec["path"]
        j["model_note"] = "%s ← %s [%s]  %s" % (j["model"], rec["file"], rec["lib"], note)


# ---------------------------------------------------------------- 主流程


def build_replacement(old_img, new_src, scale, canvas_scale, match):
    """把 new_src 放进一张与 old_img 同构的画布，返回 (Image, 统计 dict)。

    对位规则：把新素材的"参照框"渲染成与旧图"参照框"等大、中心重合。
    参照框由 match 决定 —— content（含阴影的内容足迹）或 body（产品本体）。
    """
    ocw, och = old_img.size
    if match == "body":
        ob_ref = body_bbox(old_img) or content_bbox(old_img)
        nb_ref = body_bbox(new_src) or content_bbox(new_src)
    else:
        ob_ref = content_bbox(old_img)
        nb_ref = content_bbox(new_src)

    o_rw, o_rh = ob_ref[2] - ob_ref[0] + 1, ob_ref[3] - ob_ref[1] + 1
    n_rw, n_rh = nb_ref[2] - nb_ref[0] + 1, nb_ref[3] - nb_ref[1] + 1

    # canvas_scale 自适应当：目标 = 让新素材**按原生像素 1:1 落在画布上**，
    # 而不是盲目放大旧图。旧图本身已经很大时（例如这张图已经换过一次高清，
    # 旧图从 78x62 变成 1248x992），再固定 x16 会搭出 3 亿像素画布 ——
    # PIL 直接抛 DecompressionBombError；就算不崩，素材也会被上采样十几倍变糊。
    fit = min(n_rw / float(o_rw), n_rh / float(o_rh))     # 素材原生 / 旧图参照框
    cs = max(1.0, min(canvas_scale, fit))
    if ocw * cs * och * cs > MAX_CANVAS_PX:               # 仍然过大 → 按像素上限回退
        cs = max(1.0, (MAX_CANVAS_PX / float(ocw * och)) ** 0.5)
    CW = max(1, int(round(ocw * cs)))
    CH = max(1, int(round(och * cs)))
    tw = o_rw * cs * scale
    th = o_rh * cs * scale
    ratio = min(tw / n_rw, th / n_rh)          # 等比缩放，取小者，绝不拉伸
    aw = max(1, int(round(new_src.size[0] * ratio)))
    ah = max(1, int(round(new_src.size[1] * ratio)))
    asset = premul_resize(new_src, (aw, ah))

    # 参照框中心对位
    o_cx = (ob_ref[0] + ob_ref[2] + 1) / 2.0 * cs
    o_cy = (ob_ref[1] + ob_ref[3] + 1) / 2.0 * cs
    n_cx = (nb_ref[0] + nb_ref[2] + 1) / 2.0 * ratio
    n_cy = (nb_ref[1] + nb_ref[3] + 1) / 2.0 * ratio
    px = int(round(o_cx - n_cx))
    py = int(round(o_cy - n_cy))

    canvas = Image.new("RGBA", (CW, CH), (0, 0, 0, 0))
    canvas.alpha_composite(asset, (px, py))

    overflow = (px < 0 or py < 0 or px + aw > CW or py + ah > CH)
    obb, nbb = body_bbox(old_img), body_bbox(canvas)
    cb = content_bbox(canvas)
    stats = {
        "old_canvas": (ocw, och), "new_canvas": (CW, CH),
        "canvas_scale": round(cs, 3),
        "old_content": (o_rw, o_rh),
        "new_content": (cb[2] - cb[0] + 1, cb[3] - cb[1] + 1),
        "old_body": ((obb[2] - obb[0] + 1, obb[3] - obb[1] + 1) if obb else None),
        "new_body": ((nbb[2] - nbb[0] + 1, nbb[3] - nbb[1] + 1) if nbb else None),
        "asset": (aw, ah),
        "overflow": overflow,
    }
    return canvas, stats


def main():
    ap = argparse.ArgumentParser(description="用高清素材替换 pptx 里的低清图片")
    ap.add_argument("deck")
    ap.add_argument("--new", help="新高清素材（png/jpg/webp…）")
    ap.add_argument("--new-model", help="按**型号**取素材（走图库索引），如 RG-RAP72")
    ap.add_argument("--part", help="按媒体部件名定位，逗号分隔")
    ap.add_argument("--slide", type=int, help="按页码定位（配合 --pic）")
    ap.add_argument("--pic", type=int, help="该页第几张图，1 起；省略 = 该页全部")
    ap.add_argument("--map", help='批量：OLD=NEW,OLD=NEW…（与 --new/--slide 互斥）')
    ap.add_argument("--map-model", dest="map_model",
                    help='批量按型号：OLD=型号,OLD=型号…')
    ap.add_argument("--out", help="输出的 pptx 路径")
    ap.add_argument("--scale", type=float, default=1.0,
                    help="内容足迹缩放系数；紧贴裁切的产品图常用 0.93 左右（默认 1.0）")
    ap.add_argument("--canvas-scale", type=float, default=16.0,
                    help="新画布相对旧图的**放大上限**（默认 16）。实际倍率自适应到"
                         "「素材原生像素 : 旧图参照框」≈1:1 就停，不再无谓上采样")
    ap.add_argument("--match", choices=["content", "body"], default="content",
                    help="对位依据：content=含阴影的内容足迹（默认），body=产品本体")
    ap.add_argument("--lib", choices=["webp", "png"], help="图库索引里优先取哪个库")
    ap.add_argument("--variant", help="取哪个视角：front / side / back …")
    ap.add_argument("--index", default=product_lib.DEFAULT_INDEX, help="图库索引路径")
    ap.add_argument("--root", help="图库根目录（索引不存在时据此重建）")
    ap.add_argument("--min-oversample", type=float, default=1.5,
                    help="素材相对显示尺寸的最低过采样倍率（默认 1.5，不够就升级到更清晰的库）")
    ap.add_argument("--dry-run", action="store_true", help="只出预览 PNG，不写 pptx")
    args = ap.parse_args()

    deck = os.path.abspath(args.deck)
    if not os.path.exists(deck):
        raise SystemExit("找不到 %s" % deck)
    out = os.path.abspath(args.out) if args.out else re.sub(r"\.pptx$", "", deck) + "-换图.pptx"

    zin = zipfile.ZipFile(deck)
    jobs = resolve(zin, args)
    if not jobs:
        raise SystemExit("没指定要换哪张：用 --part / --slide+--pic / --map / --map-model")
    if args.new:
        for j in jobs:
            if "asset" not in j:
                j["asset"] = args.new
    if args.new_model:
        for j in jobs:
            if "asset" not in j:
                j["model"] = args.new_model
    resolve_models(jobs, args, zin)

    if args.dry_run:
        pass
    elif args.map and args.new:
        raise SystemExit("--map 与 --new 互斥")

    report = ["# 换图报告", "",
              "| 目标 | 旧素材 | 新素材 | 旧画布 | 新画布 | 旧内容 | 新内容 | 旧本体 | 新本体 | 部件体积 |",
              "|---|---|---|---|---|---|---|---|---|---|"]
    payload = {}
    for j in jobs:
        asset = j.get("asset")
        if not asset or not os.path.exists(asset):
            raise SystemExit("素材不存在：%s" % asset)
        if j.get("model_note"):
            print("  [型号索引] %s" % j["model_note"])
        src = Image.open(asset)
        src.load()
        part = j["part"]
        old_bytes = zin.read(part)
        old_img = Image.open(io.BytesIO(old_bytes))
        ext = os.path.splitext(part)[1].lower().lstrip(".")
        if ext not in ("png", "jpg", "jpeg"):
            raise SystemExit("%s 不是位图（%s），本脚本不处理矢量/EMF" % (part, ext))

        canvas, st = build_replacement(old_img, src, args.scale, args.canvas_scale, args.match)
        data = encode(canvas, ext)
        payload[part] = data

        box = j["box"]
        scalenote = ""
        if box:
            sx, sy = box[2] / st["new_canvas"][0], box[3] / st["new_canvas"][1]
            if st["new_body"]:
                scalenote = "渲染本体 %.1fx%.1f px（旧 %.1fx%.1f）" % (
                    st["new_body"][0] * sx, st["new_body"][1] * sy,
                    (st["old_body"][0] if st["old_body"] else 0) * (box[2] / old_img.size[0]),
                    (st["old_body"][1] if st["old_body"] else 0) * (box[3] / old_img.size[1]))
        report.append("| %s | %s %dx%d | %s %dx%d | %dx%d | %dx%d | %dx%d | %dx%d | %s | %s | %.1f→%.1f KB |" % (
            j["label"], os.path.basename(part), old_img.size[0], old_img.size[1],
            os.path.basename(asset), src.size[0], src.size[1],
            st["old_canvas"][0], st["old_canvas"][1], st["new_canvas"][0], st["new_canvas"][1],
            st["old_content"][0], st["old_content"][1], st["new_content"][0], st["new_content"][1],
            st["old_body"] or "-", st["new_body"] or "-",
            len(old_bytes) / 1024, len(data) / 1024))

        print("  %-28s %dx%d -> %dx%d   PNG %.1f KB" % (
            os.path.basename(part), old_img.size[0], old_img.size[1],
            canvas.size[0], canvas.size[1], len(data) / 1024))
        if scalenote:
            print("  %-28s %s" % ("", scalenote))
        if st["overflow"]:
            print("  %-28s 警告：素材超出画布被裁切，请调小 --scale 或 --canvas-scale" % "")
        if args.dry_run:
            pv = os.path.splitext(out)[0] + "-预览-%s.png" % os.path.splitext(os.path.basename(part))[0]
            canvas.save(pv)
            print("  %-28s 预览 -> %s" % ("", pv))

    if args.dry_run:
        print("\n[dry-run] 未写 pptx。")
        return 0

    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            raw = payload.get(it.filename, zin.read(it.filename))
            zo.writestr(it, raw)
    zin.close()
    rp = os.path.splitext(out)[0] + "-换图报告.md"
    with open(rp, "w", encoding="utf-8") as fh:
        fh.write("\n".join(report) + "\n")
    print("\n-> %s  (%.2f MB)" % (out, os.path.getsize(out) / 1048576))
    print("-> %s" % rp)
    print("提醒：换完必须跑 verify_pptx.py + render_deck.py 复核。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
