# -*- coding: utf-8 -*-
"""check_links.py — 分发前的断链自检

这个 skill 是**整包分发**的：两个 .thmx 模板、311 张产品图、品牌色规范都随包走。
只要有人只拷了部分文件、或文档里的相对路径写歪，读者在别人机器上就会点空。

本脚本扫包内所有 `.md`，抽出两类引用：
    1. markdown 链接    [文字](references/xxx.md)
    2. 反引号里的路径   `assets/master/xxx.thmx`

逐个按「相对当前文件目录」与「相对包根」两种基准解析，报缺失。

用法：
    python check_links.py                  # 检查本 skill 包（默认根 = 包根）
    python check_links.py "<包根目录>"      # 检查解压出来的副本
    python check_links.py --quiet          # 只输出结论

退出码：0 = 无断链；1 = 有断链（可直接串进发布脚本）

为什么单独有这么一个脚本：
    "文件确实在压缩包里"证明不了分发可用。真实口径是
    「**解压到一个全新路径，文档里引用的每个相对路径都还在**」——
    前两次修分发问题（图库索引绝对路径、.thmx 引用写死绝对路径）都是靠这条检查抓出来的。
"""

import argparse
import os
import re
import sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 视为"包内资源"的顶层目录
KNOWN_PREFIX = ("assets/", "scripts/", "references/")
KNOWN_EXT = (".md", ".py", ".json", ".npz", ".thmx", ".potx", ".pptx",
             ".csv", ".txt", ".png", ".webp", ".jpg")

MD_LINK = re.compile(r"\]\(([^)\s]+)\)")
CODE_SPAN = re.compile(r"`([^`\n]+)`")
PATHY = re.compile(r"^(?:\.\./)*(?:assets|scripts|references)/[^\s`]+$")

SKIP_PREFIX = ("http://", "https://", "mailto:", "ftp://")


def candidates(text):
    """从一段 markdown 里抽出所有"看起来是包内相对路径"的字符串。"""
    out = list(MD_LINK.findall(text))
    for s in CODE_SPAN.findall(text):
        s = s.strip()
        if PATHY.match(s) or (s.startswith(KNOWN_PREFIX) and s.endswith(KNOWN_EXT)):
            out.append(s)
    return out


def resolve(spec, cur_dir, root):
    """按两种基准解析；返回 (路径, 状态)。状态 ∈ {ok, MISSING, skip}"""
    spec = spec.split("#")[0].strip()
    if not spec or spec.startswith(SKIP_PREFIX):
        return None, "skip"
    # 通配符 / shell 变量 / 模板占位符（如 <原件>、<工作目录>）不当成真实路径
    if any(ch in spec for ch in "*?<>") or "$" in spec or spec.endswith("/"):
        return None, "skip"
    for base in (cur_dir, root):
        p = os.path.normpath(os.path.join(base, spec))
        if os.path.exists(p):
            return p, "ok"
    return os.path.normpath(os.path.join(root, spec)), "MISSING"


def collect_md(root):
    mds = []
    for dp, dn, fn in os.walk(root):
        dn[:] = [d for d in dn if d not in ("__pycache__", ".git")]
        for f in sorted(fn):
            if f.lower().endswith(".md"):
                mds.append(os.path.join(dp, f))
    return mds


def main():
    ap = argparse.ArgumentParser(description="扫描包内文档的相对路径引用，报告断链")
    ap.add_argument("root", nargs="?", default=SKILL_DIR, help="包根目录（默认本 skill 包）")
    ap.add_argument("--quiet", action="store_true", help="只输出汇总与缺失清单")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        raise SystemExit("目录不存在：%s" % root)

    mds = collect_md(root)
    total = miss = skip = 0
    bad = []
    for p in mds:
        rel = os.path.relpath(p, root).replace("\\", "/")
        text = open(p, encoding="utf-8").read()
        n_ref = 0                      # 本文件里"可解析的引用"条数
        for spec in candidates(text):
            _, st = resolve(spec, os.path.dirname(p), root)
            if st == "skip":
                skip += 1
                continue
            total += 1
            n_ref += 1
            if st == "MISSING":
                miss += 1
                bad.append((rel, spec))
        if not args.quiet:
            print("%-36s 引用 %d" % (rel, n_ref))

    print()
    print("=" * 74)
    print("包根：%s" % root)
    print("扫描 %d 个 .md｜可解析引用 %d 处｜跳过 %d｜缺失 %d"
          % (len(mds), total, skip, miss))
    if bad:
        print("缺失清单：")
        for src, spec in bad:
            print("  %s  ->  %s" % (src, spec))
    print("结论：", "PASS —— 包内无断链" if not miss
          else "FAIL —— 有 %d 处断链，分发出去读者会点空" % miss)
    return 1 if miss else 0


if __name__ == "__main__":
    sys.exit(main())
