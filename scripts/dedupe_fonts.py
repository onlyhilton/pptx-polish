#!/usr/bin/env python3
"""PPTX 嵌入字体去重 —— 多份字节完全相同的 .fntdata 只存一份。

背景
----
PowerPoint 嵌入字体时，**同一个字体文件会被重复写进包里**。实测
`BYOD Office Network v10`：`ppt/fonts/` 12 个部件合计 9.79 MB，其中
只有 5 份不同的字节：

    sha1 2040b9f937e3  font5 font6 font8 font9 font12   1.15 MB × 5
    sha1 a2b2ba9a85f5  font3 font4 font7 font11         0.97 MB × 4
    sha1 15301f8339aa  font1   sha1 4f2598fd3244 font2
    sha1 e79a38a566e7  font10

8 个部件是纯冗余，占 7.52 MB（整包 40.68 MB 的 18.5%）。

为什么安全
----------
去重的**前提是字节完全相同**（含 OOXML 的 32 字节混淆 GUID 头）。字节相同
⇒ 解出来的字体完全相同 ⇒ 把多个 `<p:bold r:id>` 指向同一个关系，PowerPoint
拿到的数据和现在一模一样。**不做任何字体子集化/裁剪**，因此不存在"少了某个
字形"的风险。

注意：本机实测 `Noto Sans` 的 regular 与 bold 槽位本来就是同一份数据，
`Noto Sans Black`/`Noto Sans Medium` 也与它共用 —— 这是 PowerPoint 把
同一家族按不同 face 名重复嵌入的结果，不是我们造成的。

改什么
------
- `ppt/_rels/presentation.xml.rels`：被合并的 rId 的 Target 指向保留的那份
- 删除被合并的 `ppt/fonts/*.fntdata` 部件
- **其余所有部件字节原样复制**；`[Content_Types].xml` 用 `Default Extension`
  声明 fntdata，删部件无需改动

用法
----
    python dedupe_fonts.py deck.pptx --dry-run
    python dedupe_fonts.py deck.pptx -o deduped.pptx
"""

from __future__ import annotations

import argparse
import hashlib
import sys
import zipfile
from collections import OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
FONT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
PRES_RELS = "ppt/_rels/presentation.xml.rels"


def target_path(target: str) -> str:
    """关系 Target（相对 ppt/）→ 包内路径。"""
    t = target.lstrip("/")
    return t if t.startswith("ppt/") else "ppt/" + t


def plan(z: zipfile.ZipFile):
    """返回 (groups, keep_target, drop, root, rels)。

    groups      sha1 → [(rId, path, bytes, rel_el)]
    keep_target 被合并的 rId → 保留部件的 Target 字符串
    drop        要删的包内路径
    """
    root = ET.fromstring(z.read(PRES_RELS))
    rels = list(root.findall("{%s}Relationship" % RELS_NS))
    font_rels = [r for r in rels if r.get("Type") == FONT_REL]

    groups = OrderedDict()          # sha1 → [(rId, path, bytes, rel_el)]
    for r in font_rels:
        p = target_path(r.get("Target") or "")
        try:
            blob = z.read(p)
        except KeyError:
            continue
        h = hashlib.sha1(blob).hexdigest()
        groups.setdefault(h, []).append((r.get("Id"), p, len(blob), r))

    keep_target = {}
    drop = set()
    for items in groups.values():
        tgt = items[0][3].get("Target")
        for rid, p, _, _rel in items[1:]:
            keep_target[rid] = tgt
            drop.add(p)
    return groups, keep_target, drop, root, rels


def dedupe(src: Path, dst: Path, dry: bool = False) -> dict:
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        groups, keep_target, drop, root, rels = plan(z)
        before = sum(len(z.read(n)) for n in names if n.startswith("ppt/fonts/"))

        if not drop:
            return {"fonts_before": before, "fonts_after": before, "saved": 0,
                    "dropped": [], "merged": 0}

        # 改写 rels：被合并的 rId 的 Target 指向保留部件
        by_id = {r.get("Id"): r for r in rels}
        for rid, tgt in keep_target.items():
            by_id[rid].set("Target", tgt)
        ET.register_namespace("", RELS_NS)
        new_rels = ET.tostring(root, encoding="UTF-8", xml_declaration=True)

        after = before - sum(z.getinfo(p).file_size for p in drop)
        result = {
            "fonts_before": before, "fonts_after": after,
            "saved": before - after, "dropped": sorted(drop),
            "merged": len(drop), "kept": len(groups),
        }

        if dry:
            return result

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as out:
            for n in names:
                if n in drop:
                    continue
                if n == PRES_RELS:
                    out.writestr(n, new_rels)
                else:
                    out.writestr(n, z.read(n))
        return result


def main() -> int:
    ap = argparse.ArgumentParser(description="PPTX 嵌入字体去重（只合并字节完全相同的副本）")
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", help="输出文件（默认 <原名>-fonts-deduped.pptx）")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不写文件")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        print("ERROR: 文件不存在 %s" % src, file=sys.stderr)
        return 1
    dst = Path(args.out) if args.out else src.with_name(src.stem + "-fonts-deduped.pptx")

    r = dedupe(src, dst, args.dry_run)
    print("[fonts] ppt/fonts 去重前 %.2f MB → 去重后 %.2f MB，省 %.2f MB" % (
        r["fonts_before"] / 1048576, r["fonts_after"] / 1048576, r["saved"] / 1048576))
    if not r["dropped"]:
        print("[fonts] 没有字节重复的部件，无需处理")
        return 0
    print("[fonts] 保留 %d 份不同字节，删除 %d 个冗余部件：" % (r["kept"], r["merged"]))
    for p in r["dropped"]:
        print("        - %s" % p)
    if args.dry_run:
        print("[fonts] dry-run，未写文件")
    else:
        print("[fonts] 文件整体 %.2f MB → %.2f MB" % (
            src.stat().st_size / 1048576, dst.stat().st_size / 1048576))
        print("[fonts] 产出 %s" % dst)
        print("[fonts] 下一步必须跑 verify_pptx.py 并用 PowerPoint 打开验证")
    return 0


if __name__ == "__main__":
    sys.exit(main())
