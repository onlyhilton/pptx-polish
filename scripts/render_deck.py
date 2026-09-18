#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render_deck.py — 把 .pptx 渲染成逐页 PNG（Windows / PowerPoint COM）。

为什么要它：Step 2 配图与 Step 3 排版是**视觉驱动**的，XML 层只能读到坐标和色值，
读不出"挤不挤""配不配""糊不糊"。审计报告里的"压装饰保留区""元素重叠"等判定，
经常是**声明包围盒**相交而**可见墨迹**并不相交 —— 必须用渲染图复核。

依赖：本机装有 Microsoft Office（POWERPNT.EXE + 注册表 InstallRoot 有效）。
      没有 Office 时会明确报错，不会静默产出空目录。

用法：
    python render_deck.py <deck.pptx> [--out DIR] [--width 2560] [--height 1440]
                                     [--pages 1,5,22] [--pdf] [--quiet]

输出：
    <out>/slide-01.png ... slide-NN.png     逐页图（默认 2560x1440，即 2 倍于 1280x720 版心）
    <out>/deck.pdf                          可选（--pdf）
    <out>/render.log                        运行日志

设计要点（踩过的坑）：
  * PowerPoint 的 Application.Visible **不允许**设成 False（会抛
    "Hiding the application window is not allowed"）。要隐藏必须用
    Presentations.Open(..., WithWindow:=msoFalse)。
  * 逐页 Slide.Export 而不是 Presentation.Export —— 前者能自己命名，
    避免中文版 Office 生成"幻灯片1.PNG"这种带非 ASCII 的文件名。
  * PowerShell 的 .ps1 若含非 ASCII 路径会被编码破坏（Windows 已知问题），
    因此先把源文件复制到纯 ASCII 的临时目录再渲染，脚本内容保持全 ASCII。
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

POWERPOINT_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office*\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office*\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\Office*\POWERPNT.EXE",
]

PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$src   = '{src}'
$out   = '{out}'
$pdf   = '{pdf}'
$pages = @({pages})
$log   = @()
$app = $null
$pres = $null
try {{
  $app = New-Object -ComObject PowerPoint.Application
  $log += 'COM ok, version=' + $app.Version
  $app.DisplayAlerts = 1
  $pres = $app.Presentations.Open($src, $true, $false, $false)
  $n = $pres.Slides.Count
  $log += 'opened, slides=' + $n
  if ($pages.Count -eq 0) {{ $todo = 1..$n }} else {{ $todo = $pages }}
  foreach ($i in $todo) {{
    if ($i -lt 1 -or $i -gt $n) {{ $log += ('skip out-of-range ' + $i); continue }}
    $name = 'p' + $i.ToString('00') + '.png'
    $pres.Slides.Item($i).Export((Join-Path $out $name), 'PNG', {w}, {h})
    $log += ('exported ' + $name)
  }}
  $log += 'PNG export done'
  if ($pdf -ne '') {{
    try {{ $pres.SaveCopyAs($pdf, 32); $log += 'PDF done' }}
    catch {{ $log += 'PDF failed: ' + $_.Exception.Message }}
  }}
  $pres.Close()
  $app.Quit()
  $log += 'closed+quit'
}}
catch {{ $log += 'ERROR: ' + $_.Exception.ToString() }}
finally {{
  if ($pres) {{ try {{ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }} catch {{}} }}
  if ($app)  {{ try {{ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)  | Out-Null }} catch {{}} }}
  [GC]::Collect()
  $log += 'released'
  $log | Out-File -Encoding utf8 '{logfile}'
}}
"""


def find_powerpoint():
    import glob
    for pat in POWERPOINT_CANDIDATES:
        hits = glob.glob(pat)
        if hits:
            return hits[0]
    return None


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("deck")
    ap.add_argument("--out", default=None, help="输出目录（默认 <deck 同目录>/render）")
    ap.add_argument("--width", type=int, default=2560)
    ap.add_argument("--height", type=int, default=1440)
    ap.add_argument("--pages", default="", help="只渲染这些页，如 1,5,22；默认全部")
    ap.add_argument("--pdf", action="store_true", help="同时导出 PDF")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    deck = os.path.abspath(args.deck)
    if not os.path.isfile(deck):
        print("ERROR: 找不到文件 " + deck, file=sys.stderr)
        return 2

    ppt = find_powerpoint()
    if not ppt:
        print("ERROR: 未找到 POWERPNT.EXE —— 本机可能没装 Microsoft Office。", file=sys.stderr)
        print("       替代方案：LibreOffice headless（soffice --headless --convert-to pdf）", file=sys.stderr)
        print("       注意 POWERPNT.EXE 存在不等于安装有效，需查注册表 InstallRoot。", file=sys.stderr)
        return 3

    out = os.path.abspath(args.out) if args.out else os.path.join(os.path.dirname(deck), "render")
    os.makedirs(out, exist_ok=True)
    # 清掉上次的 p*.png，避免页数变化时残留
    for f in os.listdir(out):
        if re.match(r"^p\d+\.png$", f) or f == "deck.pdf":
            try:
                os.remove(os.path.join(out, f))
            except OSError:
                pass

    pages = []
    if args.pages.strip():
        pages = [int(x) for x in re.split(r"[,\s]+", args.pages.strip()) if x]

    work = tempfile.mkdtemp(prefix="pptrender_")
    pngdir = os.path.join(work, "png")
    os.makedirs(pngdir, exist_ok=True)
    try:
        src = os.path.join(work, "src.pptx")
        shutil.copy2(deck, src)          # ASCII 路径，规避 .ps1 编码问题
        pdf = os.path.join(work, "deck.pdf") if args.pdf else ""
        logfile = os.path.join(work, "render.log")
        ps = PS_TEMPLATE.format(
            src=src, out=pngdir, pdf=pdf, logfile=logfile,
            w=args.width, h=args.height,
            pages=",".join(str(p) for p in pages),
        )
        ps_path = os.path.join(work, "render.ps1")
        with open(ps_path, "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(ps)

        cmd = ["powershell.exe", "-NoProfile", "-NonInteractive",
               "-ExecutionPolicy", "Bypass", "-File", ps_path]
        if not args.quiet:
            print("-- PowerPoint COM 渲染中（不可见）……")
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)

        if os.path.isfile(logfile):
            log = open(logfile, encoding="utf-8-sig", errors="replace").read()
            with open(os.path.join(out, "render.log"), "w", encoding="utf-8") as fh:
                fh.write(log)
            if not args.quiet:
                print(log.strip())
            if "ERROR:" in log:
                print("!! 渲染过程中报错，详见 " + os.path.join(out, "render.log"), file=sys.stderr)
        else:
            print("!! 没能拿到 PowerPoint 日志；stdout/stderr：", file=sys.stderr)
            print((r.stdout or "") + (r.stderr or ""), file=sys.stderr)

        got = sorted(f for f in os.listdir(pngdir) if f.endswith(".png"))
        if not got:
            print("ERROR: 一页都没导出。", file=sys.stderr)
            return 4
        for f in got:
            shutil.move(os.path.join(pngdir, f), os.path.join(out, "slide-" + f[1:]))
        if args.pdf and os.path.isfile(pdf):
            shutil.move(pdf, os.path.join(out, "deck.pdf"))

        if not args.quiet:
            print("完成：%d 页 -> %s" % (len(got), out))
            print("  尺寸 %dx%d，文件名 slide-01.png …" % (args.width, args.height))
        return 0
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
