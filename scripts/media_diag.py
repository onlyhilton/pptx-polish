# -*- coding: utf-8 -*-
"""媒体与版式占用深度诊断

回答三个问题：
  1) 21 个版式里，哪些是"结构等价、仅背景图不同"？（版式层面的重复成因）
  2) 每个 media 文件被谁引用？哪些是孤立媒体（没人引用）？
  3) 删掉未引用版式后，能连带回收多少字节？

用法：python media_diag.py "<deck.pptx>"
"""
from __future__ import annotations

import hashlib
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


def q(t):
    pre, local = t.split(":")
    return "{%s}%s" % (NS[pre], local)


def resolve(host, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(host), target))


def main():
    src = Path(sys.argv[1])
    z = zipfile.ZipFile(src)
    names = set(z.namelist())

    def rels_of(part):
        rp = posixpath.join(posixpath.dirname(part), "_rels",
                            posixpath.basename(part) + ".rels")
        if rp not in names:
            return {}
        out = {}
        for rel in etree.fromstring(z.read(rp)):
            if rel.get("TargetMode") == "External":
                continue
            out[rel.get("Id")] = (resolve(part, rel.get("Target")),
                                  rel.get("Type", "").rsplit("/", 1)[-1])
        return out

    # ---- 谁引用谁
    media_refs = defaultdict(list)          # media -> [引用它的部件]
    part_media = defaultdict(set)           # 部件 -> {media}

    for n in names:
        if not n.endswith(".xml"):
            continue
        if n in ("[Content_Types].xml",):
            continue
        if "/_rels/" in n:
            continue
        try:
            root = etree.fromstring(z.read(n))
        except Exception:
            continue
        rm = rels_of(n)
        for blip in root.iter(q("a:blip")):
            rid = blip.get("{%s}embed" % REL_NS) or blip.get("{%s}link" % REL_NS)
            if rid and rid in rm:
                tgt, _ = rm[rid]
                if tgt.startswith("ppt/media/"):
                    media_refs[tgt].append(n)
                    part_media[n].add(tgt)

    # ---- 版式引用情况
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    pres_rels = rels_of("ppt/presentation.xml")
    masters = []
    for e in pres.iter(q("p:sldMasterId")):
        rid = e.get("{%s}id" % REL_NS)
        if rid in pres_rels:
            masters.append(pres_rels[rid][0])

    layouts = []
    for mp in masters:
        mrels = rels_of(mp)
        for e in etree.fromstring(z.read(mp)).iter(q("p:sldLayoutId")):
            rid = e.get("{%s}id" % REL_NS)
            if rid in mrels:
                layouts.append(mrels[rid][0])

    # 页面 -> 版式
    layout_used = defaultdict(list)
    for n in sorted(x for x in names if re.match(r"ppt/slides/slide\d+\.xml$", x)):
        srels = rels_of(n)
        used = None
        for rid, (tgt, typ) in srels.items():
            if typ == "slideLayout":
                used = tgt
        if used:
            layout_used[used].append(int(re.search(r"slide(\d+)", n).group(1)))

    # ---- 版式结构签名（不含 media 名，改比 media 内容 hash）
    def sig(part):
        root = etree.fromstring(z.read(part))
        rm = rels_of(part)
        phs = []
        for sp in root.iter(q("p:sp")):
            ph = sp.find(".//" + q("p:ph"))
            if ph is None:
                continue
            xf = sp.find(".//" + q("a:xfrm"))
            geo = None
            if xf is not None:
                off, ext = xf.find(q("a:off")), xf.find(q("a:ext"))
                if off is not None and ext is not None:
                    geo = (int(off.get("x")), int(off.get("y")),
                           int(ext.get("cx")), int(ext.get("cy")))
            phs.append((ph.get("type") or "body", geo))
        shps = []
        for el in root.iter(q("p:sp")):
            if el.find(".//" + q("p:ph")) is not None:
                continue
            xf = el.find(".//" + q("a:xfrm"))
            geo = None
            if xf is not None:
                off, ext = xf.find(q("a:off")), xf.find(q("a:ext"))
                if off is not None and ext is not None:
                    geo = (int(off.get("x")), int(off.get("y")),
                           int(ext.get("cx")), int(ext.get("cy")))
            shps.append((q("p:sp"), geo))
        bg = []
        for blip in root.iter(q("a:blip")):
            rid = blip.get("{%s}embed" % REL_NS)
            if rid and rid in rm:
                tgt, _ = rm[rid]
                if tgt.startswith("ppt/media/"):
                    bg.append(hashlib.sha1(z.read(tgt)).hexdigest()[:10])
        csld = root.find("." + q("p:cSld"))
        return {
            "name": (csld.get("name") if csld is not None else None),
            "ph": sorted(phs, key=str),
            "shapes": sorted(shps, key=str),
            "media_hash": tuple(sorted(bg)),
        }

    info = {}
    for lp in layouts:
        try:
            info[lp] = sig(lp)
        except Exception as e:
            info[lp] = {"name": "ERR %s" % e, "ph": [], "shapes": [], "media_hash": ()}

    # ---- 结构分组（忽略媒体名，比内容 hash）
    groups = defaultdict(list)
    for lp, s in info.items():
        key = (tuple(s["ph"]), tuple(s["shapes"]))
        groups[key].append(lp)

    print("=" * 78)
    print("一、版式结构分组（不含背景图，仅占位符+形状几何）")
    print("=" * 78)
    for gi, (key, members) in enumerate(
            sorted(groups.items(), key=lambda kv: -len(kv[1])), 1):
        if len(members) < 1:
            continue
        refs = sum(len(layout_used.get(m, [])) for m in members)
        mhs = {info[m]["media_hash"] for m in members}
        print("\n组 %d：%d 个版式，共被引用 %d 页，背景图 %d 种" % (
            gi, len(members), refs, len(mhs)))
        for m in sorted(members, key=lambda x: int(re.search(r"(\d+)", x).group(1))):
            idx = int(re.search(r"slideLayout(\d+)", m).group(1))
            used = layout_used.get(m, [])
            print("   idx=%-3d %-22s 引用=%-2d %s  背景hash=%s" % (
                idx, (info[m]["name"] or "-")[:22], len(used),
                (used if used else "（未用）"),
                ",".join(info[m]["media_hash"]) or "无"))

    # ---- media 引用与体积
    print()
    print("=" * 78)
    print("二、媒体引用与体积")
    print("=" * 78)
    unused_layouts = {m for m in layouts if not layout_used.get(m)}
    total_media = 0
    orphan = []
    only_unused = []
    in_use = []
    for n in sorted(names):
        if not n.startswith("ppt/media/"):
            continue
        sz = z.getinfo(n).file_size
        total_media += sz
        refs = media_refs.get(n, [])
        if not refs:
            orphan.append((n, sz))
        elif all(r in unused_layouts for r in refs):
            only_unused.append((n, sz, refs))
        else:
            in_use.append((n, sz, refs))

    print("ppt/media 总数=%d  合计 %.2f MB" % (
        len([n for n in names if n.startswith("ppt/media/")]), total_media / 1024 / 1024))
    print("  在用          %3d 个  %.2f MB" % (len(in_use), sum(s for _, s, _ in in_use) / 1024 / 1024))
    print("  仅未用版式引用 %3d 个  %.2f MB   <- 删版式可连带回收" % (
        len(only_unused), sum(s for _, s, _ in only_unused) / 1024 / 1024))
    print("  完全孤立      %3d 个  %.2f MB   <- 直接可回收" % (
        len(orphan), sum(s for _, s in orphan) / 1024 / 1024))

    if orphan:
        print("\n孤立媒体（前 20）：")
        for n, s in sorted(orphan, key=lambda x: -x[1])[:20]:
            print("   %-42s %8.1f KB" % (posixpath.basename(n), s / 1024))
    if only_unused:
        print("\n仅未用版式引用的媒体（前 20）：")
        for n, s, refs in sorted(only_unused, key=lambda x: -x[1])[:20]:
            print("   %-42s %8.1f KB  <- %s" % (
                posixpath.basename(n), s / 1024,
                ", ".join(posixpath.basename(r) for r in refs[:3])))

    # ---- 未使用版式清单 + 可回收
    print()
    print("=" * 78)
    print("三、未被引用的版式（%d 个）" % len(unused_layouts))
    print("=" * 78)
    xsz = 0
    for m in sorted(unused_layouts, key=lambda x: int(re.search(r"slideLayout(\d+)", x).group(1))):
        idx = int(re.search(r"slideLayout(\d+)", m).group(1))
        s = z.getinfo(m).file_size
        xsz += s
        med = [n for n in part_media.get(m, set())]
        msz = sum(z.getinfo(n).file_size for n in med if n in names)
        print("   idx=%-3d %-24s XML=%6.1f KB  专属媒体=%d 个/%.1f KB" % (
            idx, (info[m]["name"] or "-")[:24], s / 1024, len(med), msz / 1024))
    print("\n版式 XML 合计可回收 %.1f KB" % (xsz / 1024))


if __name__ == "__main__":
    main()
