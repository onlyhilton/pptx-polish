# -*- coding: utf-8 -*-
"""update_skill.py — 把 pptx-polish 更新到最新版（git 安装 / zip 安装都支持）。

为什么需要它
------------
skill 是"拷贝进来的一个文件夹"，没有安装器、没有版本号，于是：

  · 别人拿到的是**某个时间点的快照**，上游改了脚本他不知道；
  · 想更新时只能重新下载覆盖 —— 但覆盖会**连带抹掉他自己改过的东西**，且覆盖了什么他也看不见；
  · 想知道"我这份是不是最新的"更是无从下手。

这个脚本补上三件事：**能问版本**（`VERSION` + 远端比对）、**能安全更新**（git 走快进合并，
zip 安装走逐文件同步并列出变化）、**改不动就明说**（有未提交改动时拒绝更新，不静默覆盖）。

判定"新不新"的第一依据是仓库根目录的 `VERSION` 文件（不是文件时间戳 —— 时间戳在拷贝/解压后就失真）。
git 安装还会额外用 commit 对比复核一次，因为"版本号没升但内容改了"是真实存在的。

用法
----
    # 只查：本地版本 vs 远端版本，一个字都不改（退出码 2 = 有新版）
    python update_skill.py --check

    # 更新到最新
    python update_skill.py

    # 先看会写哪些文件
    python update_skill.py --dry-run

    # 网络要走代理（Python **不读** Windows 的系统代理设置，必须显式给）
    python update_skill.py --proxy http://127.0.0.1:7897

    # 机器可读
    python update_skill.py --check --json

退出码
------
    0 = 已是最新，或更新成功
    1 = 出错（网络/校验/本地有改动/合并冲突）
    2 = 发现远端有新版（`--check` 与 `--dry-run` 都返回它，方便脚本判断）

两种安装形态的处理
------------------
    git 安装（目录里有 .git）  → `git fetch <URL> main` + `git merge --ff-only`。
                                工作区有未提交改动 → **拒绝**，并打印处置命令。
                                本地有自己提交的新 commit（分叉）→ **拒绝**，不自动 merge/rebase。
    zip 安装（目录里没有 .git）→ 下载 main 的 zipball，逐文件比对（sha256）后同步；
                                跳过 `.git`/`__pycache__`，**不删除**本地多出的文件（只报告）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

REPO = "onlyhilton/pptx-polish"
BRANCH = "main"
HTML_URL = "https://github.com/%s" % REPO
CLONE_URL = "%s.git" % HTML_URL
ZIP_URL = "https://codeload.github.com/%s/zip/refs/heads/%s" % (REPO, BRANCH)
# 两个 URL 给的是同一份 zipball；有的网络只通其中一个，所以留回退
ZIP_URLS = (
    ZIP_URL,
    "https://github.com/%s/archive/refs/heads/%s.zip" % (REPO, BRANCH),
)
VERSION_URL = "https://raw.githubusercontent.com/%s/%s/VERSION" % (REPO, BRANCH)
VERSION_API = "https://api.github.com/repos/%s/contents/VERSION?ref=%s" % (REPO, BRANCH)
COMMIT_API = "https://api.github.com/repos/%s/commits/%s" % (REPO, BRANCH)

UA = "pptx-polish-updater/1.1 (+%s)" % HTML_URL
HTTP_TIMEOUT = 25

# 用 abspath 而不是 resolve()：resolve() 会解开软链（`~/.workbuddy` 在 Windows 上
# 常被链到别的盘），于是提示里打印的"安装目录"跟用户实际用的路径对不上。
SKILL_ROOT = Path(os.path.abspath(__file__)).parent.parent
SKILL_REAL = Path(__file__).resolve().parent.parent


# --------------------------------------------------------------------------- #
# 小工具
# --------------------------------------------------------------------------- #
def say(msg: str = "") -> None:
    """**不改 stdout 编码** —— 与包内其余脚本保持一致：Windows 真控制台走控制台 API 本来就正确，
    管道/重定向时按 locale（中文 Windows 是 cp936）走，读者也读得对。
    只在极端情况下（目标编码装不下某个字符）退化为可编码版本，绝不因打印而崩。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        print(msg.encode(enc, "replace").decode(enc, "replace"))


def emit_json(obj) -> None:
    data = json.dumps(obj, ensure_ascii=False, indent=2)
    try:
        sys.stdout.write(data + "\n")
    except UnicodeEncodeError:
        sys.stdout.write(json.dumps(obj, ensure_ascii=True, indent=2) + "\n")


def parse_version(text):
    """'1.2.3' -> (1,2,3)；认不出来返回 None。"""
    if not text:
        return None
    core = text.strip().split("-")[0].split("+")[0]
    parts = core.split(".")
    if not parts or len(parts) > 4:
        return None
    out = []
    for p in parts:
        if not p.isdigit():
            return None
        out.append(int(p))
    while len(out) < 3:
        out.append(0)
    return tuple(out)


def fetch_text(url: str, proxy=None, timeout=HTTP_TIMEOUT, accept=None):
    """取一个文本资源。返回 (text|None, status|None, err|None)。

    status == -1 表示连接层就失败了（DNS / 网络 / 代理不通）；有 HTTP 码就是服务端的真实回答。
    **别把 404 说成"网络不通"** —— 它只是"远端还没有这个文件"，两者的处置完全不同。
    """
    opener = urllib.request.build_opener()
    if proxy:
        opener = urllib.request.build_opener(
            urllib.request.ProxyHandler({"http": proxy, "https": proxy})
        )
    req = urllib.request.Request(url, headers={"User-Agent": UA,
                                               "Accept": accept or "*/*"})
    try:
        with opener.open(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", errors="replace"), resp.status, None
    except urllib.error.HTTPError as exc:
        return None, exc.code, "HTTP %s" % exc.code
    except Exception as exc:                                     # noqa: BLE001
        return None, -1, str(exc)[:120]


def fetch_json(url: str, proxy=None, timeout=HTTP_TIMEOUT):
    txt, status, _ = fetch_text(url, proxy=proxy, timeout=timeout)
    if not txt or not status or status < 200 or status >= 300:
        return None
    try:
        return json.loads(txt)
    except Exception:                                            # noqa: BLE001
        return None


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------- #
# git 定位与调用
# --------------------------------------------------------------------------- #
def _probe_git(exe):
    """探测一个 git 是否可用；返回 (ok, GIT_EXEC_PATH 覆盖值或 None, 备注)。

    关键坑：PortableGit 的 `cmd\\git.exe` 能跑 `--version`，但默认 `--exec-path` 指向一个
    **不存在**的 `mingw64/libexec/git-core`，于是 `git-remote-https` 找不到 —— 现象是
    `git fetch` 报 "git: 'remote-https' is not a git command"，而 `status`/`log` 全正常。
    所以不能只测 `--version`，必须确认 remote helper 真在。
    """
    exe = str(exe)
    try:
        r = subprocess.run([exe, "--exec-path"], capture_output=True, text=True, timeout=20)
    except Exception as exc:                                     # noqa: BLE001
        return False, None, "无法执行：%s" % exc
    if r.returncode != 0:
        return False, None, "`git --exec-path` 返回 %s" % r.returncode
    exec_path = Path((r.stdout or "").strip())
    needle = "git-remote-https.exe" if os.name == "nt" else "git-remote-https"
    if (exec_path / needle).is_file():
        return True, None, ""
    for cand in (Path(exe).parent, exec_path.parent,
                 Path(exe).parent.parent / "mingw64" / "bin"):
        if (cand / needle).is_file():
            return True, str(cand), "默认 exec-path 缺 remote helper，已指到 %s" % cand
    return True, None, "找不到 git-remote-https，远端操作可能失败"


def find_git():
    """返回 (git_exe, exec_path_override, note)；找不到返回 (None, None, 原因)。"""
    seen, notes, cands = [], [], []
    which = shutil.which("git")
    if which:
        cands.append(Path(which))
    base = Path(os.path.expanduser("~")) / ".workbuddy" / "binaries" / "PortableGit" / "versions"
    if base.is_dir():
        # 优先 mingw64/bin（布局最完整），再 bin/cmd；版本号倒序
        for d in sorted([p for p in base.iterdir() if p.is_dir()],
                        key=lambda p: p.name, reverse=True):
            cands += [d / "mingw64" / "bin" / "git.exe", d / "bin" / "git.exe",
                      d / "cmd" / "git.exe"]
    for root in (r"C:\Program Files\Git", r"C:\Program Files (x86)\Git"):
        cands += [Path(root) / "mingw64" / "bin" / "git.exe", Path(root) / "bin" / "git.exe",
                  Path(root) / "cmd" / "git.exe"]
    for c in cands:                              # 候选顺序已按优先级排好，第一个能用的就用它
        if not c.is_file() or c in seen:
            continue
        seen.append(c)
        ok, override, note = _probe_git(c)
        if ok:
            return str(c), override, note
        notes.append("%s：%s" % (c, note))
    if notes:
        return None, None, "；".join(notes)
    return None, None, "未找到 git（可直接改用 zip 方式更新，无需 git）"


def git_env(exec_override=None, proxy=None):
    env = dict(os.environ)
    if exec_override:
        env["GIT_EXEC_PATH"] = exec_override
    if proxy:
        env["HTTPS_PROXY"] = proxy
        env["HTTP_PROXY"] = proxy
    # 公开仓库读取不需要交互；避免在无终端环境里挂在密码提示上
    env["GIT_TERMINAL_PROMPT"] = "0"
    return env


def run_git(git, args, env, cwd=None, raw=False):
    """raw=True 时**不 strip** stdout —— `git status --porcelain` 的第一行首字符是状态位前的空格，
    一 strip 就会把后面的解析切错位（实测把 `tracked.txt` 读成 `racked.txt`）。"""
    cmd = [git] + list(args)
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, errors="replace",
                           timeout=300, env=env, cwd=cwd)
    except Exception as exc:                                     # noqa: BLE001
        return 1, "", "无法执行 git：%s" % exc
    stdout = r.stdout or ""
    return r.returncode, (stdout if raw else stdout.strip()), (r.stderr or "").strip()


# --------------------------------------------------------------------------- #
# 本地 / 远端版本信息
# --------------------------------------------------------------------------- #
def local_version():
    p = SKILL_ROOT / "VERSION"
    if p.is_file():
        try:
            return p.read_text(encoding="utf-8").strip()
        except Exception:                                        # noqa: BLE001
            return None
    return None


def remote_version(proxy=None):
    """返回 (version|None, note|None)。两个源都试，因为有的网络只通其中一个。

    远端**没有** VERSION 文件时返回 (None, '远端还没有 VERSION…') —— 那是"远端版本太老"，
    不是"网络坏了"，两者的处置完全不同。
    """
    last = None
    for url in (VERSION_URL, VERSION_API):
        txt, status, err = fetch_text(url, proxy=proxy,
                                      accept="application/vnd.github.raw")
        if status == -1:
            last = "连不上（%s）" % err
            continue
        if status == 404:
            return None, "远端仓库还没有 VERSION 文件，即 1.1.0 之前的老版本"
        if status and 200 <= status < 300 and txt and txt.strip():
            return txt.strip(), None
        last = err or ("HTTP %s" % status)
    return None, last or "未知原因"


def remote_commit(proxy=None):
    """拿得到就返回 {'sha','date','subject'}；拿不到返回 None（API 限流不该阻断判断）。"""
    data = fetch_json(COMMIT_API, proxy=proxy)
    if not isinstance(data, dict):
        return None
    commit = data.get("commit") or {}
    msg = (commit.get("message") or "").strip().splitlines()
    return {
        "sha": (data.get("sha") or "")[:7],
        "date": ((commit.get("committer") or {}).get("date") or "")[:19].replace("T", " "),
        "subject": msg[0] if msg else "",
    }


def version_cmp(local, remote):
    """返回 'behind' | 'current' | 'ahead' | 'unknown'。"""
    if not remote:
        return "unknown"
    if not local:
        return "behind"                       # 本地没有 VERSION = 老版本（1.0 就是这样发的）
    lv, rv = parse_version(local), parse_version(remote)
    if lv and rv:
        if rv > lv:
            return "behind"
        if rv < lv:
            return "ahead"
        return "current"
    return "current" if local.strip() == remote.strip() else "behind"


# --------------------------------------------------------------------------- #
# 形态 A：git 安装
# --------------------------------------------------------------------------- #
def is_git_install():
    return (SKILL_ROOT / ".git").exists()


def git_update(git, env, dry_run):
    """快进更新。返回 (status, detail, changed)。

    status ∈ current | behind（dry_run 的"有新版"）| updated | blocked | error
    """
    code, head, err = run_git(git, ["-C", str(SKILL_ROOT), "rev-parse", "HEAD"], env)
    if code != 0:
        return "error", "读不到本地 HEAD：%s" % (err or head), []
    local_head = head.strip()

    # 有未提交改动就不动 —— 快进合并会带着这些改动走，结果变成"更新了但说不清改了什么"
    code, dirty, err = run_git(
        git, ["-C", str(SKILL_ROOT), "status", "--porcelain", "--untracked-files=no"],
        env, raw=True)
    if code == 0 and dirty.strip():
        files = [ln[3:].strip() for ln in dirty.splitlines() if ln.strip()]
        return "blocked", (
            "工作区有 %d 个未提交改动，已停止更新（不静默覆盖你的改动）：\n" % len(files)
            + "".join("    %s\n" % f for f in files[:15])
            + ("    …\n" if len(files) > 15 else "")
            + "  处置：先提交，或 `git -C \"%s\" stash` 之后重跑本脚本。" % SKILL_ROOT
        ), []

    # 直接按 URL 取，不依赖 origin 是否配好、是不是 fork
    code, out, err = run_git(git, ["-C", str(SKILL_ROOT), "fetch", "--quiet", CLONE_URL, BRANCH], env)
    if code != 0:
        return "error", "fetch 失败：%s\n  提示：需要代理时加 --proxy" % (err or out), []
    code, remote_head, err = run_git(git, ["-C", str(SKILL_ROOT), "rev-parse", "FETCH_HEAD"], env)
    if code != 0:
        return "error", "读不到远端 HEAD：%s" % (err or remote_head), []
    remote_head = remote_head.strip()

    if remote_head == local_head:
        return "current", "已是 %s" % local_head[:7], []

    changed = []
    code, stat, _ = run_git(git, ["-C", str(SKILL_ROOT), "diff", "--name-only",
                                  "%s..%s" % (local_head, remote_head)], env)
    if code == 0 and stat:
        changed = [ln.strip() for ln in stat.splitlines() if ln.strip()]

    # 本地是否分叉（有远端没有的提交）
    code, mb, _ = run_git(git, ["-C", str(SKILL_ROOT), "merge-base", local_head, remote_head], env)
    if code == 0 and mb.strip() != local_head:
        return "blocked", (
            "本地与远端已分叉（本地有自己提交的 %s…%s），不自动 merge/rebase。\n"
            "  处置：自己决定 merge / rebase / 丢弃后再更新。"
            % (local_head[:7], mb.strip()[:7])
        ), changed

    if dry_run:
        return "behind", "%s → %s（%d 个文件会变）" % (local_head[:7], remote_head[:7],
                                                   len(changed)), changed

    code, out, err = run_git(git, ["-C", str(SKILL_ROOT), "merge", "--ff-only", remote_head], env)
    if code != 0:
        return "error", "快进合并失败：%s" % (err or out), changed
    return "updated", "%s → %s（%d 个文件变化）" % (local_head[:7], remote_head[:7],
                                                len(changed)), changed


# --------------------------------------------------------------------------- #
# 形态 B：zip 安装
# --------------------------------------------------------------------------- #
SKIP_NAMES = {".git", "__pycache__", ".DS_Store", "Thumbs.db"}


def _iter_files(root: Path):
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if any(part in SKIP_NAMES for part in p.relative_to(root).parts):
            continue
        yield p.relative_to(root)


def download_and_extract(proxy=None):
    """下载 main 的 zipball 并解压，返回顶层目录。失败抛 RuntimeError。"""
    tmp = Path(tempfile.mkdtemp(prefix="pptx-polish-upd-"))
    zpath = tmp / "main.zip"
    errs = []
    for url in ZIP_URLS:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        opener = urllib.request.build_opener()
        if proxy:
            opener = urllib.request.build_opener(
                urllib.request.ProxyHandler({"http": proxy, "https": proxy}))
        try:
            with opener.open(req, timeout=180) as resp, zpath.open("wb") as fh:
                shutil.copyfileobj(resp, fh)
            if zpath.stat().st_size > 1024:
                break
            errs.append("%s：下到的文件太小（%d 字节）" % (url, zpath.stat().st_size))
        except Exception as exc:                                 # noqa: BLE001
            errs.append("%s：%s" % (url, str(exc)[:80]))
    else:
        raise RuntimeError("下载失败（试了 %d 个地址）：\n    %s\n  需要代理时加 --proxy"
                           % (len(ZIP_URLS), "\n    ".join(errs)))
    try:
        with zipfile.ZipFile(zpath) as zf:
            zf.extractall(tmp / "x")
    except Exception as exc:                                     # noqa: BLE001
        raise RuntimeError("解压失败：%s" % exc)
    tops = [p for p in (tmp / "x").iterdir() if p.is_dir()]
    if len(tops) != 1:
        raise RuntimeError("zip 里顶层不止一个目录，结构不符合预期：%s" % tops)
    return tops[0]


def zip_sync(src: Path, dry_run=False):
    """把 src 的内容同步到 SKILL_ROOT。返回 (new[], changed[], extra[], same_n)。"""
    new, changed, same = [], [], 0
    src_files = set()
    for rel in _iter_files(src):
        src_files.add(rel)
        dst = SKILL_ROOT / rel
        if not dst.is_file():
            new.append(rel)
        elif sha256_file(dst) != sha256_file(src / rel):
            changed.append(rel)
        else:
            same += 1
    extra = sorted(str(rel) for rel in _iter_files(SKILL_ROOT) if rel not in src_files)
    if not dry_run:
        for rel in new + changed:
            dst = SKILL_ROOT / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src / rel, dst)
    return sorted(str(r) for r in new), sorted(str(r) for r in changed), extra, same


# --------------------------------------------------------------------------- #
# 输出
# --------------------------------------------------------------------------- #
def print_status(lv, rv, rv_note, rel, rc, git_note, extra_note=None):
    say("pptx-polish 更新检查")
    say("  仓库      : %s" % HTML_URL)
    say("  安装目录  : %s" % SKILL_ROOT)
    if SKILL_REAL != SKILL_ROOT:
        say("              （软链真实路径：%s）" % SKILL_REAL)
    say("  安装形态  : %s" % ("git 安装" if is_git_install() else "zip / 拷贝安装"))
    say("  本地版本  : %s" % (lv or "（无 VERSION —— 早于 1.1.0 的旧版）"))
    say("  远端版本  : %s" % (rv or "（%s）" % (rv_note or "取不到")))
    if rc is None:
        say("  远端最新  : 取不到（API 限流或网络不通，不影响判断）")
    else:
        say("  远端最新  : %s  %s  %s" % (rc["sha"], rc["date"], rc["subject"]))
    if git_note:
        say("  git 备注  : %s" % git_note)
    if extra_note:
        say("  注意      : %s" % extra_note)
    say()
    say("  → %s" % rel)


RELABEL = {
    "current": "已是最新",
    "behind": "有新版，可以更新",
    "ahead": "本地版本高于远端（自己改过 VERSION？）",
    "unknown": "无法判断（远端版本信息取不到；git 安装会用 commit 对比兜底）",
    "updated": "已更新",
    "dry-run": "会更新（dry-run，未写入）",
    "blocked": "已停止更新",
    "error": "出错",
}


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="把 pptx-polish 更新到最新版（git / zip 两种安装都支持）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="退出码：0=最新或成功，1=出错，2=发现远端有新版")
    ap.add_argument("--check", action="store_true", help="只查版本，不改任何文件")
    ap.add_argument("--dry-run", action="store_true", help="列出会变化的文件，不写入")
    ap.add_argument("--json", action="store_true", help="机器可读输出")
    ap.add_argument("--proxy", metavar="URL", help="HTTP(S) 代理，如 http://127.0.0.1:7897")
    ap.add_argument("--no-git", action="store_true",
                    help="即使目录里有 .git 也走 zip 同步（git 不可用时）")
    args = ap.parse_args(argv)

    out = {"skill_root": str(SKILL_ROOT), "method": None, "status": None,
           "local_version": None, "remote_version": None, "changed_files": []}

    lv = local_version()
    rv, rv_note = remote_version(args.proxy)
    rc = remote_commit(args.proxy)
    rel = version_cmp(lv, rv)
    out.update(local_version=lv, remote_version=rv)

    git_exe, exec_override, git_note = find_git()
    want_git = is_git_install() and not args.no_git
    # zip / 拷贝安装的人不关心本机 git 状态，别在报告里塞无关噪音
    shown_git_note = git_note if want_git else ""
    out["method"] = "git" if want_git else "zip"
    extra_note, git_changed = None, []

    # 版本号给不出结论（远端还没 VERSION），或版本号相同但内容可能改过 → 用 commit 复核
    if want_git and git_exe and rel in ("unknown", "current") and not args.dry_run:
        st, detail, changed = git_update(git_exe, git_env(exec_override, args.proxy),
                                         dry_run=True)
        if st == "behind":
            rel, git_changed = "behind", changed
        elif st == "current":
            rel = "current"
        elif st == "blocked":
            extra_note = detail.splitlines()[0]
        elif st == "error":
            extra_note = "commit 复核失败：%s" % detail.splitlines()[0]

    # ---------------- 只查 ----------------
    if args.check:
        out.update(status=rel, changed_files=git_changed)
        if args.json:
            emit_json(out)
        else:
            print_status(lv, rv, rv_note, RELABEL.get(rel, rel), rc, shown_git_note, extra_note)
            if rel == "behind":
                say("  跑 `python scripts/update_skill.py` 更新。")
        return 2 if rel == "behind" else (1 if rel == "error" else 0)

    if rel == "current" and not args.dry_run:
        out["status"] = "current"
        if args.json:
            emit_json(out)
        else:
            print_status(lv, rv, rv_note, RELABEL["current"], rc, shown_git_note, extra_note)
        return 0

    if rel == "ahead" and not args.dry_run:
        out["status"] = "ahead"
        if args.json:
            emit_json(out)
        else:
            print_status(lv, rv, rv_note, RELABEL["ahead"], rc, shown_git_note, extra_note)
            say("\n不做任何事（本地比远端新）。")
        return 0

    # ---------------- dry-run ----------------
    if args.dry_run:
        try:
            src = download_and_extract(args.proxy)
            new, changed, extra, same = zip_sync(src, dry_run=True)
        except RuntimeError as exc:
            say(str(exc))
            return 1
        # 内容逐字节相同就别喊"会更新" —— 版本号判断不出来不代表内容有差
        status = "dry-run" if (new or changed) else "current"
        out.update(status=status, changed_files=(new + changed)[:200])
        if args.json:
            emit_json(out)
        else:
            print_status(lv, rv, rv_note, RELABEL[status], rc, shown_git_note, extra_note)
            say("  新增 %d ｜ 修改 %d ｜ 未变 %d ｜ 本地多出 %d（不会删）"
                % (len(new), len(changed), same, len(extra)))
            for f in (new + changed)[:40]:
                say("    %s" % f)
            if len(new) + len(changed) > 40:
                say("    …（共 %d 项）" % (len(new) + len(changed)))
            if not (new or changed):
                say("  远端内容与本机逐字节相同，无需更新。")
        return 2 if status == "dry-run" else 0

    # ---------------- 真更新：git ----------------
    if want_git:
        if git_exe is None:
            say("检测到 git 安装，但本机没有可用的 git：%s" % git_note)
            say("  → 加 --no-git 改用 zip 同步，或装好 git 再跑。")
            return 1
        st, detail, changed = git_update(git_exe, git_env(exec_override, args.proxy), False)
        out.update(status=st, changed_files=changed)
        if args.json:
            emit_json(out)
        else:
            print_status(lv, rv, rv_note, RELABEL.get(st, st), rc, shown_git_note, extra_note)
            say("  %s" % detail)
            for f in changed[:40]:
                say("    %s" % f)
            if len(changed) > 40:
                say("    …（共 %d 项）" % len(changed))
            if st == "updated":
                say("\n已更新到 %s。**重启 WorkBuddy（或新开对话）后新版本才会被加载。**" % (rv or ""))
        return 0 if st in ("updated", "current") else 1

    # ---------------- 真更新：zip ----------------
    try:
        src = download_and_extract(args.proxy)
        new, changed, extra, same = zip_sync(src, dry_run=False)
    except RuntimeError as exc:
        say(str(exc))
        return 1
    out.update(status="updated", changed_files=(new + changed)[:200])
    if args.json:
        say(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        print_status(lv, rv, rv_note, RELABEL["updated"], rc, shown_git_note, extra_note)
        say("  新增 %d ｜ 覆盖 %d ｜ 未变 %d ｜ 本地多出 %d（未删除）"
            % (len(new), len(changed), same, len(extra)))
        for f in (new + changed)[:40]:
            say("    %s" % f)
        if len(new) + len(changed) > 40:
            say("    …（共 %d 项）" % (len(new) + len(changed)))
        if extra:
            say("\n以下文件是本地的、上游没有（**没有动它们**）：")
            for f in extra[:15]:
                say("    %s" % f)
            if len(extra) > 15:
                say("    …（共 %d 项）" % len(extra))
        say("\n已更新到 %s。**重启 WorkBuddy（或新开对话）后新版本才会被加载。**" % (rv or ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
