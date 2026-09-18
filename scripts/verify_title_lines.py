# -*- coding: utf-8 -*-
"""verify_title_lines.py — 用 PowerPoint 自己的排版引擎数"标题占了几行"。

为什么要它
----------
"标题不能换行"是个**排版结果**问题，XML 里读不出来：
  · 字号、框宽、内边距都在 XML 里，但断行位置由字体度量 + PowerPoint 排班引擎决定；
  · 自己用 PIL 量字宽只能**预测**，与 PowerPoint 的实际断行有亚像素差，边缘案例会猜错；
  · 渲染成 PNG 再数墨迹行，容易被装饰线/图标干扰，且**溢出到框外的第二行会被裁掉**。

本脚本走 PowerPoint COM，直接读 `TextRange.Lines().Count` —— **PowerPoint 自己数出来的行数**，
是"标题是否换行"的唯一事实源。同时输出 `BoundHeight/BoundWidth` 供交叉核对。

判定范围与 fix_title_oneline.py **共用同一份定义**（`pptx_title_scope.py`）：
标题占位符（title/ctrTitle）+ 模板驱动的标题槽位（版式说明它只放一行）。
两边范围一致才不会"一边修 A、一边验 B"。

用法
----
    python verify_title_lines.py "<deck.pptx>" [--json <out.json>] [--verbose]
    python verify_title_lines.py "<deck.pptx>" --all-boxes      # 额外审"所有单行框"
退出码：0 = 无违规；1 = 有违规（或本机 PowerPoint COM 不可用）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_title_scope as TS                                    # noqa: E402

LINE_FACTOR = TS.LINE_FACTOR
INSET_V = TS.DEF_INS["t"] + TS.DEF_INS["b"]

# PpPlaceholderType：ppPlaceholderTitle=1 / ppPlaceholderCenterTitle=3 /
# ppPlaceholderVerticalTitle=5
PH_TITLE, PH_CTR_TITLE, PH_VERT_TITLE = 1, 3, 5

PS_BODY = r"""
$ErrorActionPreference = 'Stop'
$src = '__SRC__'
$out = '__OUT__'
$rows = New-Object System.Collections.ArrayList
$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($src, $true, $false, $false)
  $si = 0
  foreach ($s in $pres.Slides) {
    $si++
    $shIdx = 0
    foreach ($sh in $s.Shapes) {
      $shIdx++
      $ok = $false
      try { if ($sh.HasTextFrame -eq -1) { if ($sh.TextFrame.HasText -eq -1) { $ok = $true } } } catch { }
      if (-not $ok) { continue }
      $tr = $sh.TextFrame.TextRange
      $txt = $tr.Text
      $txt = $txt -replace "`r", ' ' -replace "`n", ' ' -replace "`t", ' '
      $phType = -1
      try { $phType = [int]$sh.PlaceholderFormat.Type } catch { }
      $fs = 0.0
      try { $fs = [double]$tr.Font.Size } catch { }
      $nLines = 0
      try { $nLines = [int]$tr.Lines().Count } catch { }
      $nParas = 0
      try { $nParas = [int]$tr.Paragraphs().Count } catch { }
      $bw = 0.0; $bh = 0.0
      try { $bw = [double]$tr.BoundWidth; $bh = [double]$tr.BoundHeight } catch { }
      $row = @($si, $shIdx, $sh.Name, $phType, $fs, $sh.Left, $sh.Top, $sh.Width, $sh.Height, $nLines, $nParas, $bw, $bh, $txt) -join "`t"
      [void]$rows.Add($row)
    }
  }
  $pres.Close()
} finally {
  $app.Quit()
}
$rows | Out-File -Encoding utf8 -LiteralPath $out
"""


def have_powerpoint():
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "try { $x = New-Object -ComObject PowerPoint.Application; $v = $x.Version; "
             "$x.Quit(); Write-Output ('VER=' + $v) } catch { Write-Output 'NOPP' }"],
            capture_output=True, text=True, errors="replace", timeout=180)
        out = (r.stdout or "") + (r.stderr or "")
        return ("VER=" in out), out.strip()
    except Exception as e:                                       # noqa: BLE001
        return False, str(e)


def dump_shapes(deck):
    work = tempfile.mkdtemp(prefix="ppttl_")
    try:
        src = os.path.join(work, "src.pptx")
        shutil.copy2(deck, src)                    # ASCII 路径，规避 .ps1 编码问题
        out = os.path.join(work, "shapes.tsv")
        ps = PS_BODY.replace("__SRC__", src).replace("__OUT__", out)
        ps_path = os.path.join(work, "dump.ps1")
        with open(ps_path, "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(ps)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", ps_path],
            capture_output=True, text=True, errors="replace", timeout=3600)
        if not os.path.isfile(out):
            raise RuntimeError("PowerPoint 未产出结果：%s"
                               % ((r.stdout or "") + (r.stderr or ""))[:400])
        rows = []
        with open(out, encoding="utf-8-sig", errors="replace") as fh:
            for line in fh:
                p = line.rstrip("\n").split("\t")
                if len(p) < 14:
                    continue
                try:
                    rows.append(dict(
                        slide=int(p[0]), idx=int(p[1]), name=p[2], phType=int(p[3]),
                        fontSize=float(p[4]), left=float(p[5]), top=float(p[6]),
                        width=float(p[7]), height=float(p[8]), lines=int(p[9]),
                        paras=int(p[10]), boundW=float(p[11]), boundH=float(p[12]),
                        text=p[13]))
                except ValueError:
                    continue
        return rows
    finally:
        shutil.rmtree(work, ignore_errors=True)


def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def one_line_capacity(height_pt, font_pt):
    if font_pt <= 0:
        return 99
    usable = height_pt - INSET_V
    return 0 if usable <= 0 else int(usable // (font_pt * LINE_FACTOR))


def main():
    ap = argparse.ArgumentParser(
        description="用 PowerPoint 排版引擎数标题行数（标题是否换行的事实源）")
    ap.add_argument("deck")
    ap.add_argument("--json", default=None)
    ap.add_argument("--verbose", action="store_true", help="连合规项一起打印")
    ap.add_argument("--all-boxes", action="store_true",
                    help="额外审「所有单行框」（会命中正文卡片，噪声大，仅排查用）")
    ap.add_argument("--only-title", action="store_true",
                    help="只审 title/ctrTitle 占位符")
    ap.add_argument("--slots", default=None, help='只审指定槽位，如 "body:17,body:18"')
    args = ap.parse_args()

    deck = os.path.abspath(args.deck)
    if not os.path.isfile(deck):
        print("ERROR: 找不到文件 %s" % deck, file=sys.stderr)
        return 2

    # ---- 期望集合：与 fix 共用的范围定义 ----
    z = zipfile.ZipFile(deck)
    index = TS.build_index(z)
    expect = {}          # (slide_no, norm_text) -> 槽位信息
    for part, meta in index["slides"].items():
        for t in TS.iter_targets(z, index, part,
                                 include_heading=not args.only_title,
                                 slots=args.slots):
            expect[(meta["number"], norm(t["text"]))] = t
    z.close()

    ok_pp, ver = have_powerpoint()
    if not ok_pp:
        print("!! 本机 PowerPoint COM 不可用，无法数行：%s" % ver)
        print("   （此时不要声称'标题已确认单行'—— XML 度量只是预测）")
        return 1
    print("PowerPoint COM 可用（%s）" % ver)

    rows = dump_shapes(deck)
    print("=" * 116)
    print("标题行数（PowerPoint 排版引擎实测）   %s" % os.path.basename(deck))
    print("判定范围：标题占位符 + 模板驱动的标题槽位 %s" % (
        "" if not index["heading_slots"] else
        "(" + ", ".join("%s:%s" % k for k in sorted(index["heading_slots"])) + ")"))
    print("=" * 116)
    print("%-5s %-22s %-7s %-9s %-6s %-6s %-8s %-9s %s" % (
        "页", "形状", "字号pt", "框高pt", "容行", "实测", "判定", "fontScale", "标题"))
    print("-" * 116)

    viol, matched, n_scope = [], set(), 0
    for r in sorted(rows, key=lambda x: (x["slide"], x["idx"])):
        key = (r["slide"], norm(r["text"]))
        t = expect.get(key)
        broad = args.all_boxes and one_line_capacity(r["height"], r["fontSize"]) <= 1
        if t is None and not broad:
            continue
        n_scope += 1
        if t is not None:
            matched.add(key)
        cap = one_line_capacity(r["height"], r["fontSize"])
        bad = r["lines"] >= 2
        fs = (t or {}).get("font_scale")
        rec = dict(r, cap=cap, scope=("标题" if t and t.get("is_title") else
                                      ("标题槽位" if t else "单行框")),
                   fontScale=fs, expect_size=(t or {}).get("size"),
                   expect_avail=(t or {}).get("avail_w"))
        if bad:
            viol.append(rec)
        if bad or args.verbose:
            print("%-5d %-22s %-7.1f %-9.1f %-6d %-6d %-8s %-9s %s" % (
                r["slide"], r["name"][:22], r["fontSize"], r["height"], cap,
                r["lines"], "换行!" if bad else "单行 OK",
                ("%d%%" % (fs // 1000)) if fs else "—", r["text"][:40]))

    missing = [k for k in expect if k not in matched]
    print("-" * 116)
    print("纳入判定 %d 个文本框（期望 %d 个）；违规 %d 个"
          % (n_scope, len(expect), len(viol)))
    if missing:
        print("!! 有 %d 个在正文里找到、但在 PowerPoint 里没匹配到（按文字匹配）—— 需人工看："
              % len(missing))
        for s, txt in missing[:8]:
            print("   P%-3d 「%s」" % (s, txt[:56]))
    if viol:
        print()
        print("需处理（单行槽位里出现了多行）：")
        for r in viol:
            print("  P%-3d %-22s %d 行  %5.1fpt  框高 %.1fpt  %s  「%s」" % (
                r["slide"], r["name"][:22], r["lines"], r["fontSize"],
                r["height"], r["scope"], r["text"][:50]))
    print("=" * 116)
    print("结论：%s" % ("标题全部单行 —— 通过" if not viol
                        else "有 %d 处标题换行 —— 未通过" % len(viol)))

    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"file": deck, "violations": viol, "missing": missing},
                      fh, ensure_ascii=False, indent=2)
        print("JSON -> %s" % args.json)
    return 0 if not viol else 1


if __name__ == "__main__":
    sys.exit(main())
