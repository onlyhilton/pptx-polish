# -*- coding: utf-8 -*-
"""
open_test.py — 用 PowerPoint COM 真机打开测试（交付前最后一道闸）

为什么必须有这道闸：
    `verify_pptx.py` 检查的是"结构自洽"（覆盖、悬空、元素顺序），
    `python-pptx` 检查的是"这个包我能解析"。**两者都不能证明 PowerPoint 肯收。**
    实测（2026-09-17，T3T4 ISP）：产物 `[Content_Types].xml` 把模板块的伴生 .rels
    声明成 slideLayout+xml —— verify_pptx.py 报 P0=0、python-pptx 正常打开 34 页，
    而 PowerPoint 直接 `HRESULT E_FAIL` 拒收。所以**凡是要交给用户/客户的文件，
    必须真机开一次**。

用法：
    python open_test.py "a.pptx" ["b.pptx" ...] [--json <out.json>]
退出码：
    0 = 全部能打开　1 = 有打不开的（或本机没有可用 PowerPoint）
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$pp = New-Object -ComObject PowerPoint.Application
Write-Output "PP_OK"
$p = $pp.Presentations.Open("__FILE__", $false, $false, $false)
Write-Output ("OPENED slides=" + $p.Slides.Count)
$p.Close()
$pp.Quit()
"""


def have_powerpoint():
    """探测是否存在有效安装的 PowerPoint（COM 能起来就算有效）。"""
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "try { $x = New-Object -ComObject PowerPoint.Application; "
             "$v = $x.Version; $x.Quit(); Write-Output ('VER=' + $v) } "
             "catch { Write-Output 'NOPP' }"],
            capture_output=True, text=True, errors="replace", timeout=180)
        out = (r.stdout or "") + (r.stderr or "")
        return ("VER=" in out), out.strip()
    except Exception as e:                                    # noqa: BLE001
        return False, str(e)


def open_test(path):
    """把文件复制到 ASCII 临时目录再让 PowerPoint 打开，返回 (ok, detail)。"""
    # PowerPoint COM 对非 ASCII 路径会编码破坏 —— 必须先搬到 ASCII 临时目录
    d = tempfile.mkdtemp(prefix="pptopen_")
    tgt = os.path.join(d, "test.pptx")
    try:
        shutil.copy2(path, tgt)

        bat = os.path.join(d, "t.ps1")
        with open(bat, "w", encoding="utf-8") as f:
            f.write(PS_TEMPLATE.replace("__FILE__", tgt))

        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-File", bat],
            capture_output=True, text=True, errors="replace", timeout=1800)
        out = ((r.stdout or "") + (r.stderr or "")).strip()
        ok = "OPENED" in out
        line = next((l for l in out.splitlines() if l.startswith("OPENED")), None)
        if ok:
            detail = line
        else:
            err = next((l for l in out.splitlines()
                        if "E_FAIL" in l or "Error" in l or "Exception" in l), None)
            detail = err or (out.splitlines()[1] if len(out.splitlines()) > 1 else out[:200])
        return ok, detail
    except subprocess.TimeoutExpired:
        return False, "打开超时（>30min）—— 文件过大或 PowerPoint 卡住"
    finally:
        shutil.rmtree(d, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="PowerPoint 真机打开测试（交付前必跑）")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--json", default=None, help="把结果写进 JSON")
    args = ap.parse_args()

    ok_pp, ver = have_powerpoint()
    if not ok_pp:
        print("!! 本机 PowerPoint COM 不可用，无法做真机打开测试：%s" % ver)
        print("   （此时不要对外声称'文件已验证可打开'——改用渲染证实的口径）")
        return 1

    print("PowerPoint COM 可用（%s）" % ver)
    print("=" * 78)

    results, all_ok = [], True
    for p in args.files:
        if not os.path.exists(p):
            print("%-52s  SKIP 文件不存在" % os.path.basename(p))
            results.append({"file": p, "ok": None, "detail": "not found"})
            all_ok = False
            continue
        ok, detail = open_test(p)
        results.append({"file": p, "ok": ok, "detail": detail})
        print("%-52s  %s  %s" % (os.path.basename(p)[:52],
                                 "OK   " if ok else "FAIL ", detail))
        all_ok = all_ok and ok

    print("=" * 78)
    print("结论：%s" % ("全部可被 PowerPoint 打开" if all_ok else "有文件被 PowerPoint 拒收 —— 必须修"))
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
