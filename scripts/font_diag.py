# -*- coding: utf-8 -*-
"""嵌入字体深度诊断：重复、孤儿、字符集

PowerPoint 嵌入字体时，每个 typeface 会存 regular/bold/italic/boldItalic 变体，
每个变体是一份完整或子集的字体数据。长期迭代后常见两种浪费：
  1) 同一份 fntdata 被重复存成多个部件（内容完全一致）
  2) embeddedFontLst 里的关系指向重复/多余，或部件存在但没被声明（孤儿）

用法：python font_diag.py "<deck.pptx>"
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


def main():
    src = Path(sys.argv[1])
    z = zipfile.ZipFile(src)
    names = set(z.namelist())

    # rels
    prels = {}
    for rel in etree.fromstring(z.read("ppt/_rels/presentation.xml.rels")):
        rid = rel.get("Id")
        tgt = rel.get("Target")
        if tgt.startswith("/"):
            tgt = tgt.lstrip("/")
        else:
            tgt = posixpath.normpath(posixpath.join("ppt", tgt))
        prels[rid] = (tgt, rel.get("Type", "").rsplit("/", 1)[-1])

    root = etree.fromstring(z.read("ppt/presentation.xml"))
    efl = root.find("{%s}embeddedFontLst" % NS["p"])

    print("=" * 78)
    print("embeddedFontLst 声明")
    print("=" * 78)
    declared = []          # (typeface, 变体, rid, 实际部件)
    if efl is None:
        print("  无 embeddedFontLst")
    else:
        for ef in efl:
            rid = ef.get("{%s}id" % REL_NS)
            tgt = prels.get(rid, ("?", "?"))[0]
            faces = []
            for f in ef:
                faces.append((f.get("typeface"), f.get("panose"), f.get("pitchFamily"),
                              f.get("charset"), f.get("{http://schemas.openxmlformats.org/presentationml/2006/main}type")))
            print("  rId=%-8s -> %-42s" % (rid, tgt))
            for face in faces:
                print("       typeface=%-22s %s" % (face[0], face[4] or ""))
                declared.append((face[0], face[4], rid, tgt))

    # 字体部件
    parts = sorted(n for n in names if n.startswith("ppt/fonts/"))
    print()
    print("=" * 78)
    print("字体部件内容指纹")
    print("=" * 78)
    byhash = defaultdict(list)
    for p in parts:
        blob = z.read(p)
        h = hashlib.sha1(blob).hexdigest()[:12]
        byhash[h].append((p, z.getinfo(p).file_size))
    dup_groups = []
    for h, members in sorted(byhash.items(), key=lambda kv: -len(kv[1])):
        tag = "重复×%d" % len(members) if len(members) > 1 else "唯一"
        print("  [%s] %s" % (h, tag))
        for p, s in members:
            print("       %-34s %9.1f KB" % (p, s / 1024))
        if len(members) > 1:
            dup_groups.append((h, members))

    if dup_groups:
        waste = sum(sum(s for _, s in m) - max(s for _, s in m) for _, m in dup_groups)
        print()
        print("  >> 内容完全相同的重复字体部件：%d 组，冗余 %.2f MB" % (
            len(dup_groups), waste / 1024 / 1024))

    # 部件是否被声明引用
    declared_parts = {tgt for _, _, _, tgt in declared if tgt and tgt != "?"}
    orphan_parts = [p for p in parts if p not in declared_parts]
    print()
    print("=" * 78)
    print("孤儿字体部件（存在于包内，但 embeddedFontLst 未声明）")
    print("=" * 78)
    if not orphan_parts:
        print("  无")
    else:
        tot = 0
        for p in orphan_parts:
            s = z.getinfo(p).file_size
            tot += s
            print("   %-34s %9.1f KB" % (p, s / 1024))
        print("   合计 %.2f MB 可直接回收" % (tot / 1024 / 1024))

    # 字体总量
    tot = sum(z.getinfo(p).file_size for p in parts)
    print()
    print("字体部件合计 %.2f MB / 文件总 %.2f MB = %.1f%%" % (
        tot / 1024 / 1024, src.stat().st_size / 1024 / 1024,
        tot * 100.0 / src.stat().st_size))

    # 实际用到的字体（从 slide XML 扫）
    print()
    print("=" * 78)
    print("页面/版式里实际出现的 typeface")
    print("=" * 78)
    used = defaultdict(int)
    for n in names:
        if not n.endswith(".xml"):
            continue
        if not (n.startswith("ppt/slides/slide") or n.startswith("ppt/slideLayouts/slideLayout")
                or n.startswith("ppt/slideMasters/slideMaster")
                or n.startswith("ppt/theme/theme") or n.startswith("ppt/notesSlides/")):
            continue
        try:
            txt = z.read(n).decode("utf-8", "ignore")
        except Exception:
            continue
        for m in re.finditer(r'typeface="([^"]+)"', txt):
            used[m.group(1)] += 1
    for f, c in sorted(used.items(), key=lambda kv: -kv[1]):
        print("   %-28s %5d 次" % (f, c))


if __name__ == "__main__":
    main()
