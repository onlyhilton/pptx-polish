# -*- coding: utf-8 -*-
"""PPTX 包内体积分布 / 嵌入字体 / 孤立部件诊断"""
from __future__ import annotations

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
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def resolve(host, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(host), target))


def main():
    src = Path(sys.argv[1])
    z = zipfile.ZipFile(src)
    names = set(z.namelist())

    total_raw = sum(z.getinfo(n).file_size for n in names)
    print("=" * 78)
    print("文件：%s" % src.name)
    print("  压缩后  %.2f MB" % (src.stat().st_size / 1024 / 1024))
    print("  解压后  %.2f MB" % (total_raw / 1024 / 1024))
    print("  部件数  %d" % len(names))

    # ---- 按目录归类
    buckets = defaultdict(lambda: [0, 0])
    for n in names:
        if n.endswith("/"):
            continue
        if n.startswith("ppt/media/"):
            key = "ppt/media"
        elif n.startswith("ppt/slides/"):
            key = "ppt/slides"
        elif n.startswith("ppt/slideLayouts/"):
            key = "ppt/slideLayouts"
        elif n.startswith("ppt/slideMasters/"):
            key = "ppt/slideMasters"
        elif n.startswith("ppt/theme/"):
            key = "ppt/theme"
        elif n.startswith("ppt/embeddings/"):
            key = "ppt/embeddings"
        elif n.startswith("ppt/fonts/"):
            key = "ppt/fonts"
        elif n.startswith("ppt/charts/") or n.startswith("ppt/embeddings/Microsoft_Excel"):
            key = "ppt/charts"
        elif n.startswith("ppt/notesSlides/"):
            key = "ppt/notesSlides"
        elif n.startswith("ppt/handoutMasters/") or n.startswith("ppt/notesMasters/"):
            key = "ppt/notesMaster"
        elif n.startswith("ppt/"):
            key = "ppt/其他"
        elif n.startswith("docProps/"):
            key = "docProps"
        else:
            key = "包级（rels/ContentTypes）"
        b = buckets[key]
        b[0] += 1
        b[1] += z.getinfo(n).file_size

    print()
    print("=" * 78)
    print("一、体积分布")
    print("=" * 78)
    for k, (cnt, sz) in sorted(buckets.items(), key=lambda kv: -kv[1][1]):
        print("  %-26s %4d 个  %9.2f MB  %5.1f%%" % (
            k, cnt, sz / 1024 / 1024, sz * 100.0 / total_raw))

    # ---- 最大的 20 个部件
    print()
    print("=" * 78)
    print("二、最大的 20 个部件")
    print("=" * 78)
    big = sorted(((n, z.getinfo(n).file_size) for n in names if not n.endswith("/")),
                 key=lambda x: -x[1])[:20]
    for n, s in big:
        print("  %-58s %9.1f KB" % (n[:58], s / 1024))

    # ---- 媒体明细
    print()
    print("=" * 78)
    print("三、媒体明细（按体积降序前 25）")
    print("=" * 78)
    meds = [(n, z.getinfo(n).file_size) for n in names if n.startswith("ppt/media/")]
    for n, s in sorted(meds, key=lambda x: -x[1])[:25]:
        print("  %-38s %9.1f KB" % (posixpath.basename(n), s / 1024))

    byext = defaultdict(lambda: [0, 0])
    for n, s in meds:
        e = n.rsplit(".", 1)[-1].lower()
        byext[e][0] += 1
        byext[e][1] += s
    print()
    print("  按类型：")
    for e, (c, s) in sorted(byext.items(), key=lambda kv: -kv[1][1]):
        print("     %-8s %4d 个  %8.2f MB" % (e, c, s / 1024 / 1024))

    # ---- 嵌入字体
    print()
    print("=" * 78)
    print("四、嵌入字体与主题")
    print("=" * 78)
    fonts = [n for n in names if n.endswith(".fntdata")]
    for n in fonts:
        print("   %-50s %8.1f KB" % (n, z.getinfo(n).file_size / 1024))
    if not fonts:
        print("   无 .fntdata 嵌入字体部件")
    ffl = []
    for n in names:
        if n == "ppt/presentation.xml":
            root = etree.fromstring(z.read(n))
            for e in root.iter("{%s}embeddedFont" % NS["p"]):
                for f in e.iter("{%s}font" % NS["p"]):
                    ffl.append(f.get("typeface"))
    print("   presentation.xml 声明的嵌入字体：%s" % (ffl or "无"))
    th = sorted(n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n))
    print("   主题文件：%d 个 %s" % (len(th), th))
    ms = sorted(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
    ls = sorted(n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n))
    print("   母版：%d 个 %s" % (len(ms), ms))
    print("   版式：%d 个" % len(ls))

    # ---- 引用可达性：媒体
    print()
    print("=" * 78)
    print("五、引用可达性")
    print("=" * 78)
    referenced = set()
    for n in names:
        if not n.endswith(".rels"):
            continue
        host = n[:-5].replace("/_rels/", "/")
        try:
            root = etree.fromstring(z.read(n))
        except Exception:
            continue
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            t = resolve(host, rel.get("Target"))
            referenced.add(t)
    # 根 rels 的宿主特殊处理
    for rel in etree.fromstring(z.read("_rels/.rels")):
        if rel.get("TargetMode") != "External":
            referenced.add(rel.get("Target").lstrip("/"))

    unreferenced = sorted(n for n in names
                          if not n.endswith("/") and n not in referenced
                          and n not in ("[Content_Types].xml",))
    print("  未被任何关系引用的部件：%d 个" % len(unreferenced))
    grp = defaultdict(lambda: [0, 0])
    for n in unreferenced:
        key = n.rsplit("/", 1)[0] if "/" in n else "(根)"
        grp[key][0] += 1
        grp[key][1] += z.getinfo(n).file_size
    for k, (c, s) in sorted(grp.items(), key=lambda kv: -kv[1][1]):
        print("     %-36s %3d 个  %8.2f MB" % (k, c, s / 1024 / 1024))


if __name__ == "__main__":
    main()
