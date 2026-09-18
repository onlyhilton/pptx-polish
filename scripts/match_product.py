# -*- coding: utf-8 -*-
"""
match_product.py —— 认出 PPT 里的产品图是哪个型号，并判断要不要换高清版。

解决什么
--------
图库索引解决了"知道型号就能取图"，但还有一半问题没解决：**PPT 上那张小图是哪个型号？**
本脚本用感知匹配（内容裁切 → 去均值归一化 → 相关度）把 PPT 里的产品图
和 `product_lib.py` 的图库索引对上号，然后按过采样倍率给出结论：

    清晰（不用动） / 可换高清（库里有更清晰的同款） / 换不了（库里没有或图不是产品图）

这样"按型号自助修复"就闭环了：扫描 → 认型号 → 取高清 → `replace_image.py --new-model`。

用法
----
    python match_product.py deck.pptx
    python match_product.py deck.pptx --pages 22,23 --min-score 0.75
    python match_product.py deck.pptx --cache        # 重建图库特征缓存（首次自动建）

判定口径
--------
过采样倍率 `zoom = 素材有效像素 / 页面显示像素`（与 audit_deck.py 的 ZOOM_LOW 同一套）：

    zoom <= 0.60   够用（≥1.67x 过采样）
    zoom <= 0.75   偏紧
    zoom >  0.75   偏糊，若库里有同款就该换

输出
----
    <out>.md    逐图对照表：PPT 部件 / 认出的型号 / 匹配分 / 现状 / 处理建议
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
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
EMU_PER_PX = 9525

SIG = 64                      # 特征图边长
WHITE = 248
ALPHA_MIN = 8
ZOOM_OK = 0.60
ZOOM_WARN = 0.75


# ---------------------------------------------------------------- 特征

def content_bbox(rgba, white=WHITE, alpha_min=ALPHA_MIN):
    a = np.asarray(rgba).astype(np.int16)
    m = (a[:, :, 3] > alpha_min) & (a[:, :, :3].min(axis=2) < white)
    ys, xs = np.nonzero(m)
    if len(xs) == 0:
        return 0, 0, rgba.size[0] - 1, rgba.size[1] - 1
    return int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())


def signature(im, size=SIG):
    """内容裁切 → 白底合成 → 保比例居中 → 灰度归一化。返回展平向量。"""
    rgba = im.convert("RGBA")
    if max(rgba.size) > 512:
        rgba = rgba.copy()
        rgba.thumbnail((512, 512), Image.LANCZOS)
    x0, y0, x1, y1 = content_bbox(rgba)
    crop = rgba.crop((x0, y0, x1 + 1, y1 + 1))
    bg = Image.new("RGBA", crop.size, (255, 255, 255, 255))
    bg.alpha_composite(crop)
    g = bg.convert("L")
    s = max(g.size)
    canvas = Image.new("L", (s, s), 255)
    canvas.paste(g, ((s - g.width) // 2, (s - g.height) // 2))
    v = np.asarray(canvas.resize((size, size), Image.LANCZOS)).astype(np.float32).ravel()
    v -= v.mean()
    sd = float(v.std())
    return v / sd if sd > 1e-6 else v


def lib_pool(idx, libs):
    """筛出参与匹配的图库条目。

    必须过滤：`Cybrey/` 下有一批按"正面/顶/右/105.307"这类**文件夹名**当型号的条目，
    不过滤会产出 `顶`、`右`、`云办公登录界面` 这种高分假命中。
    """
    out = {}
    for k, v in idx["items"].items():
        if v["lib"] not in libs:
            continue
        key = v["key"]
        if not key.isascii() or len(key) < 4 or not any(c.isdigit() for c in key):
            continue
        out[k] = v
    return out


def build_lib_sigs(pool, cache_path):
    """图库特征矩阵（带 npz 缓存）。"""
    keys = sorted(pool.keys())
    if os.path.exists(cache_path):
        try:
            z = np.load(cache_path, allow_pickle=True)
            M0 = z["sigs"]
            if (list(z["keys"]) == keys and M0.ndim == 2
                    and M0.shape == (len(keys), SIG * SIG)):
                return M0, keys
        except Exception:
            pass
    sigs, kept = [], []
    for i, k in enumerate(keys):
        p = pool[k]["path"]
        try:
            im = Image.open(p)
            im.load()
            sigs.append(signature(im))
            kept.append(k)
        except Exception:
            pass
        if (i + 1) % 200 == 0:
            print("   图库特征 ... %d/%d" % (i + 1, len(keys)), file=sys.stderr)
    M = np.asarray(sigs, dtype=np.float32)
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    np.savez_compressed(cache_path, keys=np.asarray(kept, dtype=object), sigs=M)
    return M, kept


# ---------------------------------------------------------------- 定位

def deck_images(z):
    """返回 [{part, size, box, pages}]，按部件去重（首个出现的位置为准）。"""
    out = {}
    for n in z.namelist():
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
        if not m:
            continue
        page = int(m.group(1))
        rmap = {r.get("Id"): r.get("Target")
                for r in etree.fromstring(z.read("ppt/slides/_rels/slide%d.xml.rels" % page))}
        root = etree.fromstring(z.read(n))
        for pic in root.iter("{%s}pic" % P_NS):
            blip = pic.find(".//{%s}blip" % A_NS)
            if blip is None:
                continue
            rid = blip.get("{%s}embed" % R_NS)
            if not rid or rid not in rmap:
                continue
            part = os.path.normpath(os.path.join("ppt/slides", rmap[rid])).replace("\\", "/")
            ext = os.path.splitext(part)[1].lower()
            if ext not in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"):
                continue
            if part not in z.namelist():
                continue
            xfrm = pic.find(".//{%s}xfrm" % A_NS)
            box = None
            if xfrm is not None:
                off, ex = xfrm.find("{%s}off" % A_NS), xfrm.find("{%s}ext" % A_NS)
                if off is not None and ex is not None:
                    box = (int(off.get("x")) // EMU_PER_PX, int(off.get("y")) // EMU_PER_PX,
                           int(ex.get("cx")) // EMU_PER_PX, int(ex.get("cy")) // EMU_PER_PX)
            sr = pic.find(".//{%s}srcRect" % A_NS)
            src_rect = {k: int(sr.get(k) or 0) for k in ("l", "t", "r", "b")} if sr is not None else None
            rec = out.setdefault(part, {"part": part, "box": None, "src_rect": None, "pages": []})
            if page not in rec["pages"]:
                rec["pages"].append(page)
            if rec["box"] is None and box:
                rec["box"], rec["src_rect"] = box, src_rect
    return sorted(out.values(), key=lambda r: (r["pages"][0], r["part"]))


def effective_px(im, src_rect=None):
    """扣掉 srcRect 之后的"有效原图区域"像素。"""
    w, h = im.size
    if not src_rect:
        return w, h
    l, t, r, b = (src_rect.get(k, 0) / 100000.0 for k in ("l", "t", "r", "b"))
    return max(1.0, w * (1 - l - r)), max(1.0, h * (1 - t - b))


def content_aspect(im, src_rect=None):
    """内容裁切后的长宽比（用 srcRect 之后的有效区域）。"""
    rgba = im.convert("RGBA")
    if max(rgba.size) > 512:
        rgba = rgba.copy()
        rgba.thumbnail((512, 512), Image.LANCZOS)
    x0, y0, x1, y1 = content_bbox(rgba)
    w, h = x1 - x0 + 1, y1 - y0 + 1
    return w / float(h) if h else 1.0


def content_fill(im, white=WHITE, alpha_min=ALPHA_MIN):
    """内容像素占整幅的比例。产品抠图一般 0.3–0.9，实拍照片接近 1.0。"""
    rgba = im.convert("RGBA")
    if max(rgba.size) > 512:
        rgba = rgba.copy()
        rgba.thumbnail((512, 512), Image.LANCZOS)
    a = np.asarray(rgba).astype(np.int16)
    m = (a[:, :, 3] > alpha_min) & (a[:, :, :3].min(axis=2) < white)
    return float(m.mean())


def lib_aspect(rec):
    """图库条目内容区的长宽比。"""
    cr = rec.get("content_ratio") or rec.get("body_ratio")
    if not cr:
        return None
    w = rec["size_px"][0] * cr[0]
    h = rec["size_px"][1] * cr[1]
    return w / h if h else None


def main():
    ap = argparse.ArgumentParser(description="认出 PPT 里的产品图型号，判断要不要换高清版")
    ap.add_argument("deck")
    ap.add_argument("--index", default=product_lib.DEFAULT_INDEX)
    ap.add_argument("--root", help="图库根目录")
    ap.add_argument("--cache", action="store_true", help="强制重建图库特征缓存")
    ap.add_argument("--pages", help="只看这些页，逗号分隔")
    ap.add_argument("--min-score", type=float, default=0.93,
                    help="高置信匹配下限（默认 0.93；低于此但高于 --suspect-score 判为疑似）")
    ap.add_argument("--suspect-score", type=float, default=0.72,
                    help="疑似匹配下限（默认 0.72，以下判为认不出）")
    ap.add_argument("--aspect-tol", type=float, default=1.35,
                    help="长宽比容差（默认 1.35 = 允许 35%% 偏差，超出降为疑似）")
    ap.add_argument("--topk", type=int, default=3, help="每个图报几个候选型号")
    ap.add_argument("--libs", default="webp",
                    help="参与匹配的库（默认 webp —— 图库默认根目录就是 Webp 单库；"
                         "多库模式可写 webp,png，cybrey 有脏名不建议开）")
    ap.add_argument("--out", help="报告输出路径（默认 <deck>-型号匹配.md）")
    args = ap.parse_args()

    deck = os.path.abspath(args.deck)
    if not os.path.exists(deck):
        raise SystemExit("找不到 %s" % deck)
    out = args.out or re.sub(r"\.pptx$", "", deck) + "-型号匹配.md"

    libs = {x.strip() for x in args.libs.split(",") if x.strip()}
    idx = product_lib.load(args.index, args.root)
    pool = lib_pool(idx, libs)
    cache = os.path.splitext(args.index)[0] + "_sigs_%s.npz" % "_".join(sorted(libs))
    if args.cache and os.path.exists(cache):
        os.remove(cache)
    print("图库索引：%d 条 / 根目录 %s" % (idx["count"], idx.get("root")))
    print("参与匹配：%d 条（库=%s）" % (len(pool), ",".join(sorted(libs))))
    M, keys = build_lib_sigs(pool, cache)
    print("图库特征：%d x %d" % (M.shape[0], M.shape[1]))

    pages = None
    if args.pages:
        pages = {int(x) for x in args.pages.replace(" ", "").split(",") if x}

    z = zipfile.ZipFile(deck)
    imgs = deck_images(z)
    if pages:
        imgs = [r for r in imgs if set(r["pages"]) & pages]

    rows, todo, unmatch = [], [], []
    for r in imgs:
        data = z.read(r["part"])
        im = Image.open(io.BytesIO(data))
        im.load()
        ew, eh = effective_px(im, r["src_rect"])
        box = r["box"]
        if box:
            zoom = round(max(box[2] / ew, box[3] / eh), 2)
        else:
            zoom = None

        s = signature(im)
        sims = M @ s / s.shape[0]
        order = np.argsort(-sims)[:max(1, args.topk)]
        cand = [(pool[keys[i]]["key"], float(sims[i])) for i in order]

        # 长宽比门控：实测这是最可靠的判据（真匹配比例几乎一致），
        # 而"留白占比"不可靠 —— 图库里的 PNG 多是浅色渐变底而非纯白，
        # fill 会把渐变算成内容，导致真匹配（0.999 分）也被误杀。故只用长宽比。
        d_asp = content_aspect(im, r["src_rect"])
        d_fill = content_fill(im)
        aspect_ok, asp_note = True, ""
        if order.size:
            top = pool[keys[order[0]]]
            la = lib_aspect(top)
            if la and la > 0:
                dev = max(d_asp / la, la / d_asp)
                if dev > args.aspect_tol:
                    aspect_ok = False
                    asp_note = "长宽比不符（%.2f vs %.2f，差 %.0f%%）" % (
                        d_asp, la, (dev - 1) * 100)

        best_key, best_score = cand[0]          # best_key 是**型号名**，可直接喂 --new-model
        hi = best_score >= args.min_score and aspect_ok
        mid = best_score >= args.suspect_score and aspect_ok
        # 只有"疑似及以上"才去库里找替代素材。认不出的图也去查库，会在报告里
        # 给出"库里可换：RG-ANT90-T6.webp"这种**假信号**（分数 0.4 的图配个天线图）。
        alt, note = (product_lib.pick_for_box(best_key, idx, box) if (box and mid)
                     else (None, ""))

        if zoom is None:
            verdict = "? 无显示框"
        elif zoom <= ZOOM_OK:
            verdict = "清晰，不用动"
        elif zoom <= ZOOM_WARN:
            verdict = "偏紧，可选"
        else:
            verdict = "偏糊，建议换"

        if not mid:
            verdict = "认不出（非产品图 / 库里没有）"
            unmatch.append(r)
        elif not hi:
            verdict = "疑似 %s（需人工确认）" % best_key
            unmatch.append(r)
        elif verdict.startswith("偏") and alt:
            todo.append((r, best_key, best_score, zoom, alt, note))
        if asp_note and mid:
            verdict += "　⚠ " + asp_note

        rows.append({
            "pages": r["pages"], "part": os.path.basename(r["part"]),
            "native": im.size, "eff": (round(ew), round(eh)), "box": box, "zoom": zoom,
            "match": best_key if mid else "-", "score": round(best_score, 3),
            "cand": cand, "verdict": verdict, "aspect": d_asp, "fill": d_fill,
            "aspect_ok": aspect_ok,
            "alt": (alt["file"], alt["lib"], round(alt["bytes"] / 1024)) if alt else None,
            "note": note,
        })

    lines = ["# PPT 产品图 × 图库 型号匹配", "",
             "图库根目录：`%s`　索引：%d 条" % (idx.get("root"), idx["count"]),
             "匹配口径：内容裁切 → 白底合成 → 保比例归一化相关度；"
             "**高置信 ≥ %.2f**，%.2f–%.2f 判为疑似，且都要过**长宽比门控**（容差 %.2f）"
             % (args.min_score, args.suspect_score, args.min_score, args.aspect_tol),
             "过采样倍率 zoom = 有效原图像素 / 页面显示像素（zoom ≤ %.2f 够用，> %.2f 偏糊）"
             % (ZOOM_OK, ZOOM_WARN), "",
             "| 页 | 部件 | 原生 | 有效 | 显示 | zoom | 内容长宽比 | 认出型号 | 匹配分 | 结论 | 库里可换 |",
             "|---|---|---|---|---|---|---|---|---|---|---|"]
    for r in rows:
        lines.append("| %s | %s | %dx%d | %d×%d | %s | %s | %.2f | %s | %.3f | %s | %s |" % (
            ",".join("P%d" % p for p in r["pages"]), r["part"],
            r["native"][0], r["native"][1], r["eff"][0], r["eff"][1],
            ("%.0f×%.0f" % (r["box"][2], r["box"][3])) if r["box"] else "-",
            r["zoom"] if r["zoom"] is not None else "-",
            r["aspect"], r["match"], r["score"], r["verdict"],
            ("`%s` [%s] %d KB" % r["alt"]) if r["alt"] else "-"))

    lines += ["", "## 可自动修复（本轮）", ""]
    if todo:
        lines.append("```bash")
        for r, k, sc, zoom, alt, note in todo:
            lines.append("python replace_image.py deck.pptx --part %s --new-model %s --match body --out fixed.pptx"
                         % (r["part"], k))
        lines.append("```")
        lines.append("")
        lines.append("| 部件 | 型号 | 匹配分 | zoom | 库内可用 |")
        lines.append("|---|---|---|---|---|")
        for r, k, sc, zoom, alt, note in todo:
            lines.append("| %s | %s | %.3f | %.2f | `%s` [%s] |" % (
                os.path.basename(r["part"]), k, sc, zoom, alt["file"], alt["lib"]))
    else:
        lines.append("无（所有认得出的产品图过采样都达标）。")

    lines += ["", "## 认不出 / 疑似 的图（需人工判断）", ""]
    if unmatch:
        lines.append("| 页 | 部件 | 原生 | zoom | 结论 | 候选 top3（型号 / 分） |")
        lines.append("|---|---|---|---|---|---|")
        for r in unmatch:
            row = next(x for x in rows if x["part"] == os.path.basename(r["part"]))
            cands = "、".join("%s %.2f" % (k, s) for k, s in row["cand"])
            lines.append("| %s | %s | %dx%d | %s | %s | %s |" % (
                ",".join("P%d" % p for p in r["pages"]), row["part"],
                row["native"][0], row["native"][1],
                row["zoom"] if row["zoom"] is not None else "-",
                row["verdict"], cands))
    else:
        lines.append("无。")

    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")

    print("\n%-16s %-10s %-8s %-6s %-14s %-6s %s" % (
        "部件", "原生", "显示", "zoom", "认出型号", "分", "结论"))
    for r in rows:
        print("%-16s %-10s %-8s %-6s %-14s %-6s %s" % (
            r["part"][:16],
            "%dx%d" % r["native"],
            ("%.0fx%.0f" % (r["box"][2], r["box"][3])) if r["box"] else "-",
            r["zoom"] if r["zoom"] is not None else "-",
            r["match"][:14], "%.3f" % r["score"], r["verdict"]))
    print("\n可自动修复 %d 张；认不出 %d 张" % (len(todo), len(unmatch)))
    print("-> %s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
