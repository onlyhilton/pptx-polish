# -*- coding: utf-8 -*-
"""
product_lib.py —— 产品高清图库索引：把"一堆按型号命名的图片"变成"按型号可取"。

为什么需要它
------------
修 PPT 里的糊图时，每次手工去找素材路径既慢又容易找错版本（同型号常有
透明底 webp / 高分辨率 png / 多视角 front-side 多个文件）。这个脚本扫描
图库目录，按**型号**建索引，之后 `replace_image.py --new-model RG-RAP72`
就能直接命中，不用写路径。

图库命名约定（实测）
--------------------
    <型号>.webp / <型号>.png                 主图
    <型号>-Front.png / <型号>-Side.png       视角变体
    <型号>-1.png                             同型号第 2 张
    <型号>-Front-175mm.png                   带尺寸的工程图
    <型号> for US.png / <型号> V2 for Canada.png   区域版
    <名称>-removebg-preview.png              抠图残留后缀

型号 key 的规范化只剥"视角/序号/尺寸/区域"这些**附属标记**，
不动 V1/V2/V3 这类真实硬件版本（RG-NBS3100-48GT4SFP-P-V2 与 -P 是两个产品）。

用法
----
    # 建索引（扫全库，首次约 1 分钟）
    python product_lib.py --build

    # 按型号查
    python product_lib.py --find RG-RAP72
    python product_lib.py --find "nbs3100"          # 模糊匹配

    # 列出某系列
    python product_lib.py --list RG-ES

    # 库里有什么型号和某个词沾边
    python product_lib.py --grep "RAP62"

图库根目录
----------
**图库随 skill 包分发**，解析顺序（`find_root()`）：

    1. 显式 --root
    2. 环境变量 RUIJIE_PIC_ROOT
    3. **包内： <SKILL_DIR>/assets/product_images/**   ← 默认命中这个
    4. <OneDrive>/PIC/产品图片/Webp        （外部单库）
    5. <OneDrive>/PIC/产品图片             （外部多库兜底）

`<OneDrive>` 由环境变量 `OneDrive` / `OneDriveConsumer` / `OneDriveCommercial`
现算，**不写死盘符**（D 盘、C 盘、换用户都能命中）。目录结构不同就改 `PIC_SUBPATH`。

包内优先，是为了让不同机器上取到**同一套素材**。想跟随外部最新库
（如 OneDrive 里新加了型号），用 `--root` 或 `RUIJIE_PIC_ROOT` 显式指定。

### 跨机器怎么还能找到图（重要）

索引里每条记录存两个路径字段：

    rel  = 相对图库根目录，如 "RG-ES/RG-ES106D-P V2.webp"   → 跨机器可移植
    path = 建库时的绝对路径                                  → 仅本机有效

`load()` 返回前一律过 `rebase()`：只要索引带 `root_rel`（图库在包内就会带），
就用 `<SKILL_DIR>/<root_rel>/<rel>` 把 `path` 重算成本机真实位置。
**消费方（replace_image / match_product）因此一行都不用改。**

同理，`load()` 判断"要不要重建索引"用的是 `root_key`
（包内 → `skill:assets/product_images`；外部 → `abs:<绝对路径>`），
而不是绝对路径 —— 否则同一份包解压到别人机器上会被误判成"图库挪位"而重建。

Webp 是各库里覆盖最全的一支（309 型号，比 PNG 多 30 个），且透明底、
体积只有同分辨率 PNG 的 1/35 —— 所以它是**唯一默认素材源**，PNG 不参与常规取图。

两种扫描模式（`_libraries()` 自动判定，不用手选）：
  * **单库模式** —— root 下没有已知子库名（如 root 直接指向 `Webp/` 或
    `product_images/`）：把 root 本身当一个库扫。
  * **多库模式** —— root 下存在 Webp/PNG/云桌面/Cybrey 等已知子目录
    （如 root 指向 `产品图片/`）：按 SUBLIBS 逐库扫，rank 决定优先级。

⚠ 包内目录名固定 `product_images`，靠 SUBLIBS 显式钉住 `lib="webp"`。
改名会让 lib 变化 → `match_product.py` 的特征缓存文件名（`_sigs_webp.npz`）
失配 → 310 张图全部重算特征。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_INDEX = os.path.join(SKILL_DIR, "assets", "product_index.json")

# ---- 随包分发的图库（对外发布的主要素材来源）--------------------------
# 图库与索引一起进 zip，解压到任何位置都能直接取图。
# ⚠ 目录名固定为 product_images，**不要改成别的名字**：
#    单库模式下 `lib` 名默认取目录名，改名会让 lib 从 webp 变成别的，
#    `match_product.py` 的特征缓存文件名（product_index_sigs_webp.npz）
#    随之失效，310 张图要全部重算特征。SUBLIBS 里已显式钉住 lib=webp。
BUNDLED_DIRNAME = "product_images"
BUNDLED_ROOT = os.path.join(SKILL_DIR, "assets", BUNDLED_DIRNAME)

# 图库子目录 → 用途标签。webp 是透明底小文件（PPT 首选素材），png 是高分辨率大文件
SUBLIBS = {
    "Webp": {"lib": "webp", "rank": 0, "note": "透明底 webp，文件小，PPT 首选"},
    # 包内图库目录：显式映射到 lib=webp，保持与外部 Webp/ 库完全同构
    BUNDLED_DIRNAME: {"lib": "webp", "rank": 0, "note": "随包分发的透明底 webp"},
    "PNG": {"lib": "png", "rank": 1, "note": "高分辨率 png，文件大"},
    "云桌面": {"lib": "cloud", "rank": 2, "note": "云桌面终端 CT 系列"},
    "Cybrey": {"lib": "cybrey", "rank": 3, "note": "Cybrey"},
}

# 外部图库子路径（OneDrive 下的相对位置）。可改这一个常量适配自己的目录结构。
PIC_SUBPATH = ("PIC", "产品图片")


def _onedrive_roots():
    """OneDrive 根目录，从环境变量取 —— **不写死盘符**。

    写死某个盘符下的 OneDrive 路径的问题：对别人是死路径，而且把"本机 OneDrive
    放在哪个盘"这种跟 skill 无关的信息一起公开了。环境变量 OneDrive /
    OneDriveConsumer / OneDriveCommercial 由 OneDrive 客户端自己维护，
    D 盘、C 盘、换用户都命中。
    """
    roots, seen = [], set()
    for k in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial"):
        v = os.environ.get(k)
        if v and os.path.isdir(v):
            a = os.path.abspath(v)
            if a.lower() not in seen:
                seen.add(a.lower())
                roots.append(a)
    home_od = os.path.expanduser("~/OneDrive")
    if os.path.isdir(home_od):
        a = os.path.abspath(home_od)
        if a.lower() not in seen:
            seen.add(a.lower())
            roots.append(a)
    return roots


def _external_candidates():
    """外部图库候选，按 OneDrive 根逐个展开：先单库（只看 Webp/），再多库兜底。"""
    out = []
    for r in _onedrive_roots():
        base = os.path.join(r, *PIC_SUBPATH)
        out.append(os.path.join(base, "Webp"))    # 单库模式：直接指向 Webp/
        out.append(base)                          # 多库模式：整个产品图片目录
    return out


# 解析顺序：包内图库 → 本机外部图库 → 整个产品图片目录（多库模式）。
# 包内优先是为了让"随包分发"的副本直接生效，保证不同机器上取到的是同一套素材。
# 想跟随外部最新库（如 OneDrive 里新加了型号），用 --root 或 RUIJIE_PIC_ROOT 显式指定。
ROOT_CANDIDATES = [BUNDLED_ROOT] + _external_candidates()

IMG_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff"}

# 附属标记：视角 / 序号 / 尺寸 / 区域 —— 剥掉后归到同一型号
#
# ⚠ RE_SEQ 只能匹配**单个数字**（-1 / -2），不能写成 \d{1,2}：
#   实测 RG-CS88-08、RG-ANT20S-90、RG-NIS-PA120-48 都是**真实型号**，
#   一旦允许两位数字，这些型号会被腰斩成 RG-CS88 / RG-ANT20S / RG-NIS-PA120。
RE_VIEW = re.compile(r"(?:-|\s)(FRONT|SIDE|BACK|BOTTOM|COVER|DETAIL|MAIN)(?=-|\s|$)", re.I)
RE_SEQ = re.compile(r"-\d$")
RE_DIM = re.compile(r"(?:-|\s)\d+(\.\d+)?MM(?=-|\s|$)", re.I)
RE_REGION = re.compile(r"\s+FOR\s+([A-Z]{2,3})(?:\s|$)", re.I)
RE_REMOVEBG = re.compile(r"-?REMOVEBG(-PREVIEW)?$", re.I)
RE_EXTRA_SPACE = re.compile(r"\s+")

# 停产目录：图库里 `已停产/`（RG-ES、RG-RAP 下各有一份）放的是停产型号，
# 型号名与在售目录**完全相同**，不区分就会把停产件当常规件取出来。
RE_EOL_DIR = re.compile(r"停产|停售|EOL|DISCONTINU", re.I)


def find_root(explicit=None):
    for c in ([explicit] if explicit else []) + [os.environ.get("RUIJIE_PIC_ROOT")] + ROOT_CANDIDATES:
        if c and os.path.isdir(c):
            return os.path.abspath(c)
    return None


def rel_to_skill(path):
    """path 位于 skill 包内时返回相对 SKILL_DIR 的 posix 路径，否则 None。

    "包内相对路径"是跨机器唯一稳定的标识：包解压到谁的机器上，相对结构都一样；
    而绝对路径换台机器必然失效。
    """
    if not path:
        return None
    try:
        r = os.path.relpath(os.path.abspath(path), SKILL_DIR)
    except ValueError:                     # Windows 跨盘符无法求相对路径
        return None
    if r.startswith("..") or os.path.isabs(r):
        return None
    return r.replace("\\", "/")


def root_key(root):
    """把图库根目录归一化成**可跨机器比较**的标识。

    包内目录 → `skill:assets/product_images`（换机器仍相同，不触发重建）
    外部目录 → `abs:<normcase 绝对路径>`（挪位即失效，触发重建）
    """
    if not root:
        return None
    rel = rel_to_skill(root)
    if rel:
        return "skill:" + rel
    return "abs:" + os.path.normcase(os.path.abspath(root))


def rebase(idx):
    """把索引里的 `path` 校正到**当前机器**上的真实位置（就地改写）。

    索引里的 path 是建库时的绝对路径，换台机器必然失效。只要索引带 `root_rel`
    （图库随包分发就会带），就能用 `<SKILL_DIR>/<root_rel>/<rel>` 重算。

    在 load() 里就地改写 path，`replace_image.py` / `match_product.py`
    等消费方一行都不用改。
    """
    rr = idx.get("root_rel")
    if not rr:
        return idx
    root = os.path.join(SKILL_DIR, *rr.split("/"))
    if not os.path.isdir(root):
        return idx
    n = 0
    for rec in idx["items"].values():
        rel = rec.get("rel")
        if not rel:
            continue
        p = os.path.join(root, *rel.split("/")).replace("\\", "/")
        if os.path.normcase(p) != os.path.normcase(rec.get("path", "")):
            rec["path"] = p
            n += 1
    idx["rebase_fixed"] = n
    return idx


def parse_name(filename):
    """文件名 → (型号 key, 附属标记列表)。"""
    stem = os.path.splitext(os.path.basename(filename))[0]
    s = RE_EXTRA_SPACE.sub(" ", stem).strip()
    tags = []

    m = RE_REGION.search(s)
    if m:
        tags.append("region:" + m.group(1).upper())
        s = RE_REGION.sub(" ", s).strip()

    s = RE_REMOVEBG.sub("", s).strip()
    if s != stem:
        tags.append("removebg")

    changed = True
    while changed:
        changed = False
        for rx, tag in ((RE_VIEW, "view"), (RE_DIM, "dim"), (RE_SEQ, "seq")):
            m = rx.search(s)
            if m:
                g = m.group(0).lstrip("-")
                tags.append("%s:%s" % (tag, g.lower()))
                s = (s[:m.start()] + s[m.end():]).strip()
                changed = True

    key = RE_EXTRA_SPACE.sub(" ", s).strip().upper()
    return key, tags


def _probe(path):
    """读图片规格：原生尺寸、是否有 alpha、内容/本体占比（用缩小代理算，快）。"""
    try:
        import numpy as np
        from PIL import Image
    except Exception as e:            # pragma: no cover
        return {"error": "缺少 numpy/Pillow: %s" % e}
    try:
        im = Image.open(path)
        w, h = im.size
        has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
        rgba = im.convert("RGBA")
        if max(w, h) > 400:
            rgba.thumbnail((400, 400), Image.LANCZOS)
        a = np.asarray(rgba).astype(np.int16)
        al = a[:, :, 3]
        rgbmin = a[:, :, :3].min(axis=2)
        content = (al > 8) & (rgbmin < 248)
        body = (al > 240) & (rgbmin > 235)
        pw, ph = rgba.size

        def ratio(mask):
            ys, xs = np.nonzero(mask)
            if len(xs) == 0:
                return None
            return [round((xs.max() - xs.min() + 1) / pw, 4),
                    round((ys.max() - ys.min() + 1) / ph, 4)]

        # has_alpha 只说"格式支持透明"（RGBA 但 alpha 全 255 也报 True）；
        # alpha_cov 才是**真的透明占比** —— 判断素材能不能直接压在彩色底上，看它。
        alpha_cov = float((al < 250).mean())
        return {
            "size_px": [w, h],
            "alpha": bool(has_alpha),
            "alpha_cov": round(alpha_cov, 4),
            "transparent": bool(alpha_cov > 0.02),
            "content_ratio": ratio(content),
            "body_ratio": ratio(body),
            "fill": round(float(content.mean()), 4),
        }
    except Exception as e:
        return {"error": str(e)}


def _libraries(root):
    """判定扫描模式，返回 [(子目录名 or None, meta)]。

    root 下有已知子库目录 → 多库模式；否则把 root 本身当**单库**扫
    （lib 名取目录名小写，如 `Webp/` → `webp`）。
    """
    subs = [(s, m) for s, m in SUBLIBS.items()
            if os.path.isdir(os.path.join(root, s))]
    if subs:
        return subs, "multi"
    name = os.path.basename(os.path.normpath(root))
    known = SUBLIBS.get(name, {})
    return [(None, {"lib": known.get("lib") or name.strip().lower(),
                    "rank": known.get("rank", 0),
                    "note": "单库模式：%s" % name})], "single"


def build(root=None, out=DEFAULT_INDEX, quiet=False):
    root = find_root(root)
    if not root:
        raise SystemExit("找不到图库根目录，请用 --root 指定（或设 RUIJIE_PIC_ROOT）")
    libs, mode = _libraries(root)
    items = {}
    n_file = n_err = 0
    for sub, meta in libs:
        base = root if sub is None else os.path.join(root, sub)
        if not os.path.isdir(base):
            continue
        for dp, dn, fn in os.walk(base):
            rel = os.path.relpath(dp, base)
            serie = "" if rel == "." else rel.split(os.sep)[0]
            for f in sorted(fn):
                if os.path.splitext(f)[1].lower() not in IMG_EXTS:
                    continue
                p = os.path.join(dp, f)
                key, tags = parse_name(f)
                if not key:
                    continue
                active = not RE_EOL_DIR.search(rel)
                rec = {
                    "key": key, "lib": meta["lib"], "rank": meta["rank"],
                    "series": serie, "active": active,
                    "file": f,
                    # rel  = 相对图库根目录 → 跨机器可移植（包内图库靠它定位）
                    # path = 建库时的绝对路径 → 仅本机有效，load() 会用 rel 重算
                    "rel": os.path.relpath(p, root).replace("\\", "/"),
                    "path": p.replace("\\", "/"),
                    "tags": tags,
                    "bytes": os.path.getsize(p),
                    "mtime": int(os.path.getmtime(p)),
                }
                rec.update(_probe(p))
                if "error" in rec:
                    n_err += 1
                slot = "%s|%s|%s" % (key, meta["lib"], ",".join(tags))
                old = items.get(slot)
                # 同键冲突（同一型号在 `已停产/` 与在售目录各放一份）：
                # 在售的赢，谁先谁后都不影响结果。
                if old is not None and not (rec["active"] and not old.get("active", True)):
                    continue
                items[slot] = rec
                n_file += 1
                if not quiet and n_file % 50 == 0:
                    print("   ... %d" % n_file, file=sys.stderr)
    idx = {
        "built": time.strftime("%Y-%m-%d %H:%M:%S"),
        "root": root,
        # root_key 用"包内相对路径 / 外部绝对路径"两种口径之一，让 load() 能
        # 判断"换了台机器"与"图库真的挪了位置"——前者不该重建索引。
        "root_key": root_key(root),
        # root_rel 让索引知道"图库就在 skill 包里"，可跨机器重算 path
        "root_rel": rel_to_skill(root),
        "mode": mode,
        "libs": [m["lib"] for _, m in libs],
        "count": len(items),
        "errors": n_err,
        "items": items,
    }
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(idx, fh, ensure_ascii=False, indent=1)
    keys = {v["key"] for v in items.values()}
    if not quiet:
        print("模式 %s（%s）/ 索引 %d 条 / %d 个型号 / 读图失败 %d   -> %s" % (
            mode, "+".join(idx["libs"]), len(items), len(keys), n_err, out))
    return idx


def load(index=DEFAULT_INDEX, root=None):
    """读索引；不存在、或**根目录与当前候选不一致**就现建。

    ⚠ 这里必须拿 `find_root(root)`（而不是只在 root 显式传入时）去比对：
    否则把图库从「产品图片/」改成「产品图片/Webp/」之后，旧索引会被静默
    沿用，`--new-model` 取到的还是旧库里的图。实测踩过。

    比对用 `root_key` 而不是绝对路径：图库随包分发时，索引标注的是"包内相对
    路径"，同一份索引解压到谁机器上都算匹配，不必重建（重建要求对方装了
    numpy/Pillow，还要跑几十秒）。

    返回前一律过 `rebase()`，把 path 校正成本机的真实位置。
    """
    if os.path.exists(index):
        with open(index, encoding="utf-8") as fh:
            idx = json.load(fh)
        r = find_root(root)
        # 旧索引没有 root_key → 拿它的 root 现算一个，保持向后兼容
        old = idx.get("root_key") or root_key(idx.get("root"))
        if r and root_key(r) != old:
            return build(r, index, quiet=True)
        return rebase(idx)
    return build(root, index)


def candidates(key, idx):
    """返回某型号的全部候选，按 在售 → 库优先 → 无附属标记 → 文件大 排序。"""
    k = key.strip().upper()
    out = [v for v in idx["items"].values() if v["key"] == k]
    if not out:                                  # 型号带空格差异时放宽
        nk = RE_EXTRA_SPACE.sub("", k)
        out = [v for v in idx["items"].values() if RE_EXTRA_SPACE.sub("", v["key"]) == nk]
    out.sort(key=lambda v: (not v.get("active", True), v["rank"], len(v["tags"]), -v["bytes"]))
    return out


def pick(key, idx, prefer=None, variant=None, need_alpha=False):
    """按型号挑最合适的一张。返回记录或 None。

    prefer   : 'webp' / 'png' / ... 指定库
    variant  : 'front' / 'side' / ... 指定视角（tags 里 view:xxx）
    need_alpha: True 时只挑有透明通道的
    """
    cands = candidates(key, idx)
    if prefer:
        sub = [c for c in cands if c["lib"] == prefer]
        if sub:
            cands = sub
    if variant:
        v = variant.lower()
        sub = [c for c in cands if any(t in ("view:" + v, "dim:" + v) for t in c["tags"])]
        if sub:
            cands = sub
    if need_alpha:
        sub = [c for c in cands if c.get("transparent")]
        if sub:
            cands = sub
    return cands[0] if cands else None


def pick_for_box(key, idx, box, prefer=None, variant=None, min_oversample=1.5):
    """按"要铺多大"挑素材：先按库优先序取最小的够用素材，不够再升级。

    判据用过采样倍率 —— 素材有效像素 / 页面显示像素 ≥ min_oversample 才算够清晰
    （和 audit_deck.py 的 ZOOM_LOW 同一套口径）。

    返回 (记录, 说明字符串) 或 (None, 原因)。
    """
    cands = candidates(key, idx)
    if variant:
        v = variant.lower()
        sub = [c for c in cands if any(t in ("view:" + v, "dim:" + v) for t in c["tags"])]
        if sub:
            cands = sub
    if prefer:
        sub = [c for c in cands if c["lib"] == prefer]
        if sub:
            cands = sub
    if not cands:
        return None, "图库里没有 %s" % key

    if not box:                                   # 不知道多大 → 直接按库优先序
        cands.sort(key=lambda c: (not c.get("active", True), c["rank"], len(c["tags"]), -c["bytes"]))
        return cands[0], "（未知显示尺寸，取库优先序）"

    dw, dh = box[2], box[3]
    need = min_oversample

    def oversample(c):
        cr, sr = c.get("content_ratio"), c.get("body_ratio") or c.get("content_ratio")
        if not sr:
            return 0.0
        eff_w, eff_h = c["size_px"][0] * sr[0], c["size_px"][1] * sr[1]
        if eff_w <= 0 or eff_h <= 0:
            return 0.0
        return min(eff_w / dw, eff_h / dh)

    ok = [c for c in cands if oversample(c) >= need]
    pool = ok or cands
    pool = sorted(pool, key=lambda c: (not c.get("active", True), c["rank"], c["bytes"]))
    best = pool[0]
    note = "显示 %.0fx%.0f px，过采样 %.2fx %s%s" % (
        dw, dh, oversample(best),
        "(达标)" if best in ok else "(仍不足，已是库里最清晰)",
        "" if best.get("active", True) else " [已停产]")
    return best, note


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(description="产品高清图库索引")
    ap.add_argument("--build", action="store_true", help="重建索引")
    ap.add_argument("--index", default=DEFAULT_INDEX, help="索引文件路径")
    ap.add_argument("--root", help="图库根目录")
    ap.add_argument("--find", help="按型号查（精确优先，失败则模糊）")
    ap.add_argument("--grep", help="模糊搜型号名")
    ap.add_argument("--list", dest="series", help="列出某系列（子目录名，如 RG-ES）")
    ap.add_argument("--variant", help="指定视角：front / side …")
    ap.add_argument("--prefer", choices=["webp", "png"], help="指定库")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args()

    if args.build:
        build(args.root, args.index)
        return 0

    idx = load(args.index, args.root)
    print("索引来自 %s（%s 建，%d 条，根目录 %s%s）" % (
        args.index, idx.get("built"), idx["count"], idx.get("root"),
        " ｜ 包内相对路径 %s" % idx["root_rel"] if idx.get("root_rel") else ""))

    def show(recs):
        if args.json:
            print(json.dumps(recs, ensure_ascii=False, indent=1))
            return
        for r in recs:
            cr = r.get("content_ratio")
            br = r.get("body_ratio")
            print("  [%-5s] %-40s %5dx%-5d %s %s %6.1f KB  %s" % (
                r["lib"], r["file"][:40], r["size_px"][0], r["size_px"][1],
                "A" if r.get("alpha") else " ",
                ("c=%s" % cr) if cr else "c=-",
                r["bytes"] / 1024,
                ("body=%s" % br) if br else ""))
            print("           %s%s" % (r.get("rel") or r["path"],
                                       "" if r.get("active", True) else "   ← 已停产"))
            if r["tags"]:
                print("           tags=%s" % ",".join(r["tags"]))

    if args.find:
        recs = candidates(args.find, idx)
        if not recs:
            print("没有精确命中 %r，模糊结果：" % args.find)
            recs = [v for v in idx["items"].values()
                    if args.find.upper() in v["key"]][:40]
        if args.variant or args.prefer:
            best = pick(args.find, idx, prefer=args.prefer, variant=args.variant)
            if best:
                print("→ 选中：%s" % best["path"])
        show(recs)
    elif args.grep:
        recs = sorted({v["key"]: v for v in idx["items"].values()
                       if args.grep.upper().replace(" ", "") in v["key"].replace(" ", "")}.values(),
                      key=lambda v: v["key"])
        print("命中 %d 个型号：" % len(recs))
        show(recs)
    elif args.series:
        recs = [v for v in idx["items"].values() if v["series"] == args.series]
        recs.sort(key=lambda v: (v["rank"], v["key"]))
        print("系列 %s：%d 条" % (args.series, len(recs)))
        show(recs)
    else:
        keys = sorted({v["key"] for v in idx["items"].values()})
        print("共 %d 个型号" % len(keys))
        show([idx["items"][k] for k in list(idx["items"])[:0]])
    return 0


if __name__ == "__main__":
    sys.exit(main())
