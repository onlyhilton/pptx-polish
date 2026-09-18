# -*- coding: utf-8 -*-
"""pack_skill.py — 把本 skill 打成可分发的 zip，并做「解压到全新路径」的可移植性自检。

为什么需要它：SKILL.md 规定「分发前跑一次 check_links.py」，判据不是"文件在
压缩包里"，而是「**解压到一个全新路径后，文档里引用的每个相对路径都还在**」。
但"打包"这个动作本身以前是手工的，于是出现过**包比源目录旧**的事故 ——
实测 2026-09-17 10:05 那版包只有 18 个 .py，而源目录已有 21 个，
`dedupe_masters.py` / `fix_title_style.py` / `open_test.py` 三个脚本整轮没发出去，
而 check_links.py 对这情况**报 PASS**（它只查"文档引用的东西在不在"，不查"有没有漏发"）。
所以"打包 + 验证"必须合成一个动作，且必须**从产物而不是从源目录**去自检。

用法：
    python pack_skill.py                       # → ./dist/pptx-polish.zip
    python pack_skill.py -o /path/out.zip
    python pack_skill.py --verify              # 打包后解压到临时目录做功能自检
    python pack_skill.py --verify --keep       # 保留解压目录（排查用）

自检五项（全部从**产物**上做，不从源目录做）：
    1) check_links  文档相对引用无断链
    2) 图库索引      path 必须 rebase 到解压目录，不能是建库机的绝对路径
    3) 模板解析      别名与裸文件名都要在包内命中
    4) 脚本语法      **扫全量**（不写死清单），逐个 py_compile
    5) 清单对账      包内相对路径集合必须与源目录**逐一相等**（防漏发）

退出码：0 = 成功（含自检通过）；1 = 自检失败
"""

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL_NAME = os.path.basename(SKILL_DIR)
SKIP_DIRS = {'__pycache__', '.git', '.pytest_cache', 'node_modules', '.venv',
             'dist', '.idea', '.vscode'}
SKIP_FILES = {'.DS_Store', 'Thumbs.db'}
SKIP_EXT = {'.pyc', '.pyo'}
FIXED_DATE = (2026, 9, 17, 0, 0, 0)          # 时间戳归一 → 产物可复现


def collect(root):
    out = []
    for dp, ds, fs in os.walk(root):
        ds[:] = sorted(d for d in ds if d not in SKIP_DIRS)
        for f in sorted(fs):
            if f in SKIP_FILES or os.path.splitext(f)[1].lower() in SKIP_EXT:
                continue
            full = os.path.join(dp, f)
            out.append((os.path.relpath(full, root).replace('\\', '/'), full))
    return sorted(out)


def pack(dst, src=SKILL_DIR):
    os.makedirs(os.path.dirname(os.path.abspath(dst)) or '.', exist_ok=True)
    if os.path.exists(dst):
        os.remove(dst)
    files = collect(src)
    total = sum(os.path.getsize(f) for _, f in files)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for rel, full in files:
            zi = zipfile.ZipInfo('%s/%s' % (SKILL_NAME, rel), date_time=FIXED_DATE)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            with open(full, 'rb') as fh:
                z.writestr(zi, fh.read())
    md5 = hashlib.md5(open(dst, 'rb').read()).hexdigest()
    print('打包完成 %s' % dst)
    print('  条目 %d ｜ 原始 %.1f MB ｜ 压缩后 %.1f MB ｜ md5 %s'
          % (len(files), total / 1048576.0, os.path.getsize(dst) / 1048576.0, md5))
    return files, md5


def verify(zip_path, keep=False, root=None):
    """解压到全新临时路径，逐项功能自检。返回 (全部通过?, 详情行列表)。"""
    root = root or SKILL_DIR
    fresh = tempfile.mkdtemp(prefix='skillpack_')
    lines, ok = [], True
    try:
        with zipfile.ZipFile(zip_path) as z:
            if z.testzip() is not None:
                return False, ['zip 完整性：损坏']
            z.extractall(fresh)
        pkg = os.path.join(fresh, SKILL_NAME)
        s = os.path.join(pkg, 'scripts')
        lines.append('解压到 %s' % fresh)
        lines.append('  包内：%s' % sorted(os.listdir(pkg)))

        # 1) 文档相对引用无断链
        p = subprocess.run([sys.executable, os.path.join(s, 'check_links.py'),
                            '--quiet', pkg], cwd=fresh, capture_output=True,
                           text=True, encoding='utf-8', errors='replace')
        lines.append('1) check_links: exit=%d  %s' % (
            p.returncode, [l for l in (p.stdout or '').splitlines() if '结论' in l] or ''))
        ok &= p.returncode == 0

        # 2) 图库索引：path 必须 rebase 到解压目录，不能是建库机的绝对路径
        code = ('import sys,os;sys.path.insert(0,r"@S@");import product_lib as pl;'
                'idx=pl.load();rec=pl.pick("RG-NBS3100-24GT4SFP-P",idx);'
                'p=(rec or {}).get("path") or "";'
                'print("root_key="+str(idx.get("root_key")));'
                'print("条目="+str(len(idx.get("items",{}))));'
                'print("落在包内="+str(os.path.abspath(p).startswith(r"@F@")))'
                ).replace('@S@', s).replace('@F@', fresh)
        p = subprocess.run([sys.executable, '-c', code], cwd=fresh,
                           capture_output=True, text=True, encoding='utf-8',
                           errors='replace')
        o = p.stdout or ''
        lines.append('2) 图库索引: %s' % o.replace('\n', ' | ').strip())
        ok &= '落在包内=True' in o

        # 3) 模板解析：别名与裸文件名都要在包内命中
        code = ('import sys,os;sys.path.insert(0,r"@S@");import inject_template as it;'
                'ns=["default","dark","2025 Ruijie Reyee PPT Template-20250530.thmx"];'
                'bad=[t for t in ns if not os.path.abspath(it.resolve_theme(t))'
                '.startswith(r"@F@")];print("未命中包内:"+str(bad))'
                ).replace('@S@', s).replace('@F@', fresh)
        p = subprocess.run([sys.executable, '-c', code], cwd=fresh,
                           capture_output=True, text=True, encoding='utf-8',
                           errors='replace')
        o = (p.stdout or '').strip()
        lines.append('3) 模板解析: %s' % o)
        ok &= '未命中包内:[]' in o

        # 4) 脚本语法：**扫全量**，不写死清单
        #    写死清单的代价：新增脚本不进自检，语法错了也照样"通过"。
        #    （2026-09-18 修：Step 5 新增的 pptx_title_scope / fix_title_oneline /
        #      verify_title_lines 三个脚本当时就不在旧清单里，等于没有自检。）
        scripts = sorted(f for f in os.listdir(s) if f.endswith('.py'))
        bad = []
        for name in scripts:
            p = subprocess.run([sys.executable, '-m', 'py_compile',
                                os.path.join(s, name)],
                               capture_output=True, text=True, encoding='utf-8',
                               errors='replace')
            if p.returncode != 0:
                tail = [l for l in (p.stderr or '').strip().splitlines() if l.strip()]
                bad.append('%s: %s' % (name, tail[-1][:120] if tail else '?'))
        for b in bad:
            lines.append('4) 语法检查 FAIL %s' % b)
        lines.append('4) 脚本语法检查: %d/%d OK' % (len(scripts) - len(bad), len(scripts)))
        ok &= not bad

        # 5) 清单对账：源目录有的，包里必须都有（防"包比源目录旧/漏发"）
        #    注意顺序：必须在本步之前不产生新文件 —— py_compile 会写 __pycache__，
        #    靠 SKIP_DIRS 排除；这也正是第 4 步放在第 5 步之前的原因。
        src_files = {rel for rel, _ in collect(root)}
        pkg_files = set()
        for dp, ds, fs in os.walk(pkg):
            ds[:] = [d for d in ds if d not in SKIP_DIRS]
            for f in fs:
                if f in SKIP_FILES or os.path.splitext(f)[1].lower() in SKIP_EXT:
                    continue
                pkg_files.add(os.path.relpath(os.path.join(dp, f), pkg).replace('\\', '/'))
        missing = sorted(src_files - pkg_files)
        extra = sorted(pkg_files - src_files)
        lines.append('5) 清单对账: 源 %d ｜ 包 %d ｜ 漏发 %d ｜ 多余 %d'
                     % (len(src_files), len(pkg_files), len(missing), len(extra)))
        if missing:
            lines.append('   漏发: %s' % missing[:10])
        if extra:
            lines.append('   多余: %s' % extra[:10])
        ok &= not missing

        lines.append('结论：%s' % ('可移植，可分发' if ok else '自检未通过'))
        return ok, lines
    finally:
        if keep:
            print('（解压目录保留在 %s）' % fresh)
        else:
            shutil.rmtree(fresh, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description='打包本 skill 为 zip，可选可移植性自检')
    ap.add_argument('-o', '--out', default=os.path.join(os.getcwd(), 'dist',
                                                       '%s.zip' % SKILL_NAME))
    ap.add_argument('--root', default=SKILL_DIR, help='源目录（默认本 skill 包）')
    ap.add_argument('--verify', action='store_true', help='打包后解压到临时目录做功能自检')
    ap.add_argument('--keep', action='store_true', help='保留自检用的解压目录')
    args = ap.parse_args()

    pack(args.out, args.root)
    if not args.verify:
        return 0
    print()
    ok, lines = verify(args.out, args.keep, args.root)
    print('\n'.join(lines))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
