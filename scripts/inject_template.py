# -*- coding: utf-8 -*-
"""
inject_template.py — 把 .thmx / .pptx 模板的母版体系注入到目标 pptx（原位改造）

这一步解决的核心问题：
    「凡是有标题的页面，标题都应写在标题框中」——但光把标题塞进占位符还不够，
    如果页面挂的还是旧的（默认 Office）版式，换了模板照样错位、黑底变白底。
    真正让文件"规整"的动作是：把页面重新挂到模板自己的 slideLayout 上。

本脚本只做结构性动作，不改任何文字内容：
    1. 注入模板的 slideMaster / slideLayouts / theme / media / 嵌入字体 / tags
    2. 重建关系链（.rels）与 [Content_Types].xml
    3. 按映射把每张幻灯片重新挂到模板版式上
    4. 报告占位符承接情况（孤儿内容预警）——这是最容易出事的地方

模板从哪来：
    模板**随包内嵌**在 assets/master/，`--theme` 默认取浅色模板，不需要写路径：
        default / light → 2025 Ruijie Reyee PPT Template-20250530.thmx       （浅色底）
        dark            → 2025 Ruijie Reyee PPT Template-Dark-20250620.thmx （深色底）
    也接受包内裸文件名、唯一子串，或外部 .thmx/.pptx 的完整路径。
    解析全程在包内进行 —— 抄命令到别人机器上照样能跑。

用法：
    # ① 先看模板有哪些版式 + 自动映射建议（不写文件）
    python inject_template.py --source in.pptx --outline

    # ② 预演
    python inject_template.py --source in.pptx --out out.pptx \
        --map "1:2,2:10,3:12,4:12,5:16" --dry-run

    # ③ 实跑
    python inject_template.py --source in.pptx --out out.pptx \
        --map "1:2,2:10,3:12,4:12,5:16"

    # ④ 不指定 --map，走自动推断；要深色底加 --theme dark
    python inject_template.py --source in.pptx --theme dark --out out.pptx

产物：
    输出 pptx
    <out>.inject-report.json / .md   注入报告

红线：
    - 不修改任何文字、不增删页面
    - 保留原母版（除非 --drop-old-master）；所有页面指向新模板版式
"""

import argparse
import copy
import json
import posixpath
import re
import shutil
import zipfile
from collections import OrderedDict
from pathlib import Path

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
}
REL_T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
PKG_REL_T = "http://schemas.openxmlformats.org/package/2006/relationships/"
# 注意：r:id 属性的命名空间 URI 末尾**不带**斜杠，与关系类型前缀 REL_T 不同
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CT = {
    "slideMaster": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "slideLayout": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "tags": "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
    "themeManager": "application/vnd.openxmlformats-officedocument.themeManager+xml",
}
DEFAULT_EXT_CT = {
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "fntdata": "application/x-fontdata",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "png": "image/png",
    "svg": "image/svg+xml",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "mp4": "video/mp4",
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
    "xml": "application/xml",
}

# 模板注入时搬运的目录（相对 base 目录）
MOVE_DIRS = (
    "slideMasters/",
    "slideLayouts/",
    "theme/",
    "media/",
    "embeddings/",
    "fonts/",
    "tags/",
)

# 不搬运的部件（文件名小写命中即跳过）：
#   themeManager / 主题缩略图是 .thmx 的包级外壳，注入 pptx 后只会成为孤儿部件
EXCLUDE_HINTS = ("thememanager", "thumbnail")

# 版式名称 -> 用途推断（用于自动映射）
KIND_COVER = "cover"
KIND_TOC = "toc"
KIND_CONTENT = "content"
KIND_TITLE_ONLY = "title_only"
KIND_BLANK = "blank"
KIND_BACK = "back"
KIND_SECTION = "section"

# auto_map 的结构判据阈值（**与 SKILL.md 的映射表一一对应**，改这里要同步改文档）
BODY_GROUP_MIN = 4        # body 成组的下限个数（≥4 视为"目录/索引"形态）
BODY_GROUP_IDX0 = 17      # 模板里成组 body 的起始 idx
TOC_LINE_MAX = 40         # 目录条目单行字数上限（超过就不是目录条目，是有内容的正文格子）
TOC_CHARS_MAX = 200       # 目录页合计正文字数上限
SECTION_CHARS_MAX = 80    # 章节过渡页合计字数上限（如 "01" + "New Trends"）


def qn(tag):
    pfx, local = tag.split(":")
    return "{%s}%s" % (NS[pfx], local)


def _local(tag):
    """取 '{ns}LocalName' 的 LocalName（无命名空间时原样返回）。"""
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def read_xml(blob):
    return etree.fromstring(blob)


def to_bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# -----------------------------------------------------------------------------
# 1. 解析模板
# -----------------------------------------------------------------------------

def template_base(names):
    if "theme/presentation.xml" in names:
        return "theme/"
    if "ppt/presentation.xml" in names:
        return "ppt/"
    raise SystemExit("模板包内找不到 presentation.xml")


def layout_kind(name, ltype, ph_types, ph_list):
    """判定版式用途。

    以**占位符结构**为准，名字只做兜底 —— 两个模板的命名并不一致：
      Light 版封面叫 "General cover"，Dark 版同类版式叫 "smart home"；
      Dark 版内容页叫 "Title and content"，若按名字含 content 判会误判成目录。
    """
    n = (name or "").strip().lower()

    if "封底" in n or "back cover" in n or n in ("back", "ending", "end", "thank you"):
        return KIND_BACK
    if "section" in n or "章节" in n:
        return KIND_SECTION
    if "blank" in n or "空白" in n:
        return KIND_BLANK

    has_ctr = "ctrTitle" in ph_types
    has_title = "title" in ph_types
    has_sub = "subTitle" in ph_types
    body_idx = sorted(int(i) for t, i in ph_list if t == "body" and i is not None)

    # 封面：居中标题（通常还带副标题）—— 名字里有没有 cover 都算
    if has_ctr:
        return KIND_COVER
    # 仅标题
    if has_title and "body" not in ph_types:
        return KIND_TITLE_ONLY
    # 标题 + 内容
    if has_title and "body" in ph_types:
        if len(body_idx) >= 4 and body_idx[0] >= 17:
            return KIND_TOC                 # body 成组（idx 从 17 起）→ 目录/索引
        return KIND_CONTENT
    # 没有标题、只有 body
    if body_idx:
        if len(body_idx) >= 4 and body_idx[0] >= 17:
            return KIND_TOC                 # 如 Light 版 layout10「目录」
        return KIND_SECTION                 # 单/双 body，如 layout15「Section Title」
    return KIND_BLANK


def pick_cover(covers, prefer=None):
    """从封面类版式里挑一个作为默认封面。covers = [(index, name), ...]"""
    if prefer:
        return prefer
    for i, nm in covers:
        low = (nm or "").lower()
        if "general" in low or "office" in low:
            return i
    return covers[0][0] if covers else None


def parse_template(z, base):
    """返回 {layouts:[...], masters:[...], base, ...}"""
    names = z.namelist()
    out = {"base": base, "layouts": [], "masters": [], "media": [], "fonts": [], "parts": []}

    for n in names:
        rel = n[len(base):] if n.startswith(base) else None
        if not rel:
            continue
        if any(rel.startswith(d) for d in MOVE_DIRS):
            out["parts"].append(n)

    # layouts
    for n in sorted(
        (x for x in names if re.match(re.escape(base) + r"slideLayouts/slideLayout\d+\.xml$", x)),
        key=lambda s: int(re.search(r"slideLayout(\d+)", s).group(1)),
    ):
        root = read_xml(z.read(n))
        cSld = root.find(qn("p:cSld"))
        lname = cSld.get("name") if cSld is not None else ""
        ph_list = []
        for ph in root.iter(qn("p:ph")):
            ph_list.append((ph.get("type") or "body", ph.get("idx")))
        ph_types = {t for t, _ in ph_list}
        idx = int(re.search(r"slideLayout(\d+)", n).group(1))
        out["layouts"].append({
            "index": idx,
            "part": n,
            "name": lname,
            "type": root.get("type") or "cust",
            "ph_types": sorted(ph_types),
            "ph_list": ph_list,
            "kind": layout_kind(lname, root.get("type"), ph_types, ph_list),
        })

    # masters
    for n in sorted(x for x in names if re.match(re.escape(base) + r"slideMasters/slideMaster\d+\.xml$", x)):
        out["masters"].append({"part": n, "index": int(re.search(r"slideMaster(\d+)", n).group(1))})

    out["media"] = sorted(x for x in names if "/media/" in x)
    out["fonts"] = sorted(x for x in names if x.endswith(".fntdata"))
    return out


# -----------------------------------------------------------------------------
# 2. 目标文件解析
# -----------------------------------------------------------------------------

def parse_target(z):
    names = z.namelist()
    out = {"names": names, "slides": [], "masters": [], "layouts": [], "media": [], "size": None,
           "has_notes_master": "ppt/notesMasters/notesMaster1.xml" in names}
    pres = read_xml(z.read("ppt/presentation.xml"))
    sz = pres.find(qn("p:sldSz"))
    if sz is not None:
        out["size"] = (int(sz.get("cx")), int(sz.get("cy")))

    for n in sorted(names, key=lambda s: int(re.search(r"slide(\d+)\.xml$", s).group(1))
                    if re.search(r"slide(\d+)\.xml$", s) else 0):
        if not re.match(r"ppt/slides/slide\d+\.xml$", n):
            continue
        root = read_xml(z.read(n))
        title = None
        for ph in root.iter(qn("p:ph")):
            t = ph.get("type")
            if t in ("title", "ctrTitle"):
                # 找到对应形状的文字
                sp = ph.getparent().getparent().getparent()
                txt = "".join(t.text or "" for t in sp.iter(qn("a:t"))) if sp is not None else ""
                title = (txt or "").strip()
                break
        if title is None:
            # 回退：找第一个有文字的 sp
            for sp in root.iter(qn("p:sp")):
                txt = "".join(t.text or "" for t in sp.iter(qn("a:t"))).strip()
                if txt:
                    title = txt
                    break
        phs = []
        for ph in root.iter(qn("p:ph")):
            sp = ph.getparent().getparent().getparent()
            txt = "".join(x.text or "" for x in sp.iter(qn("a:t"))) if sp is not None else ""
            phs.append({"type": ph.get("type") or "body", "idx": ph.get("idx"),
                        "text": txt.strip()})
        # 当前挂的 layout
        cur_layout = None
        rels_path = n.replace("slides/", "slides/_rels/") + ".rels"
        if rels_path in names:
            rx = read_xml(z.read(rels_path))
            for rel in rx:
                if rel.get("Type", "").endswith("/slideLayout"):
                    cur_layout = posixpath.normpath(
                        posixpath.join("ppt/slides", rel.get("Target")))
        # 当前版式自身的形态：名字 + 占位符数量。
        # 判"这页是不是设计成满版页"必须靠它 —— 页面没有占位符、而它挂的版式也没有占位符，
        # 才说明作者本来就把整页交给自由形状（如深色满版页 / Demo 页）。
        cur_lay_name, cur_lay_ph = "", None
        if cur_layout and cur_layout in names:
            lx = read_xml(z.read(cur_layout))
            csld = lx.find(qn("p:cSld"))
            cur_lay_name = ((csld.get("name") if csld is not None else "") or "")
            cur_lay_ph = sum(1 for _ in lx.iter(qn("p:ph")))
        out["slides"].append({
            "index": int(re.search(r"slide(\d+)\.xml$", n).group(1)),
            "part": n,
            "title": (title or "")[:70],
            "placeholders": phs,
            "current_layout": cur_layout,
            "current_layout_name": cur_lay_name,
            "current_layout_ph": cur_lay_ph,
            "text_len": sum(len(t.text or "") for t in root.iter(qn("a:t"))),
        })

    for n in names:
        if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n):
            out["masters"].append(n)
        if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n):
            out["layouts"].append(n)
        if n.startswith("ppt/media/"):
            out["media"].append(n)
    return out


# -----------------------------------------------------------------------------
# 3. 映射与重写
# -----------------------------------------------------------------------------

def build_path_map(tpl, prefix):
    """模板内部路径 -> 目标 pptx 内部路径"""
    m = {}
    for part in tpl["parts"]:
        rel = part[len(tpl["base"]):]                    # 例：slideLayouts/slideLayout12.xml
        bn = posixpath.basename(rel).lower()
        if any(h in bn for h in EXCLUDE_HINTS):
            continue                                     # themeManager / 缩略图不搬
        if rel.startswith("media/"):
            newname = prefix + "_" + posixpath.basename(rel)
            m[part] = "ppt/media/" + newname
        elif rel.startswith("fonts/"):
            newname = prefix + "_" + posixpath.basename(rel)
            m[part] = "ppt/fonts/" + newname
        elif rel.startswith("embeddings/"):
            newname = prefix + "_" + posixpath.basename(rel)
            m[part] = "ppt/embeddings/" + newname
        else:
            # slideMasters/slideLayouts/theme/tags：目录名保留，文件名加前缀
            d, f = posixpath.split(rel)
            m[part] = "ppt/" + d + "/" + prefix + f[0].upper() + f[1:]
        # rels 文件同步
        if "/_rels/" in part or part.endswith(".rels"):
            pass
    # 递归补上所有 .rels：由宿主部件路径推导
    for part in list(m):
        d, f = posixpath.split(part)
        if "/_rels/" in d:
            continue
        host_dir, host_file = d, f
        rels_part = posixpath.join(host_dir, "_rels", host_file + ".rels")
        if rels_part in tpl["parts"] or rels_part in tpl["_all_names"]:
            target_dir, target_file = posixpath.split(m[part])
            m[rels_part] = posixpath.join(target_dir, "_rels", target_file + ".rels")
    return m


def resolve_target(host_part, target):
    """把 rels 里的相对 Target 解析成包内绝对路径"""
    if target.startswith("/"):
        return target.lstrip("/")
    if "://" in target or target.startswith("http"):
        return None
    base_dir = posixpath.dirname(host_part)
    return posixpath.normpath(posixpath.join(base_dir, target))


def relativize(host_part, abs_target):
    base_dir = posixpath.dirname(host_part)
    return posixpath.relpath(abs_target, base_dir)


def rewrite_rels(blob, tpl_host, dst_host, path_map):
    """重写 .rels 里所有 Target。

    关键：用**模板内的宿主路径** tpl_host 解析出模板内绝对路径（映射表的 key），
    查出目标绝对路径后，再相对化到**目标宿主** dst_host。两者不能混用。
    """
    root = read_xml(blob)
    changes = []
    for rel in root:
        t = rel.get("Target")
        mode = rel.get("TargetMode")
        if not t or mode == "External":
            continue
        abs_tpl = resolve_target(tpl_host, t)
        if abs_tpl is None:
            continue
        if abs_tpl in path_map:
            new_abs = path_map[abs_tpl]
            rel.set("Target", relativize(dst_host, new_abs))
            changes.append({"rel_id": rel.get("Id"), "type": rel.get("Type", "").rsplit("/", 1)[-1],
                            "from": t, "to": rel.get("Target")})
    return to_bytes(root), changes


def add_content_type(ct_root, part_name, content_type):
    """给 [Content_Types].xml 增加 Override（若不存在）"""
    pn = "/" + part_name
    for ov in ct_root.findall(qn("ct:Override")):
        if ov.get("PartName") == pn:
            if ov.get("ContentType") != content_type:
                ov.set("ContentType", content_type)
            return False
    el = etree.SubElement(ct_root, qn("ct:Override"))
    el.set("PartName", pn)
    el.set("ContentType", content_type)
    return True


def ensure_default(ct_root, ext, content_type):
    ext = ext.lower()
    for d in ct_root.findall(qn("ct:Default")):
        if (d.get("Extension") or "").lower() == ext:
            return False
    el = etree.SubElement(ct_root, qn("ct:Default"))
    el.set("Extension", ext)
    el.set("ContentType", content_type)
    return True


def max_rid(rels_root):
    mx = 0
    for rel in rels_root:
        m = re.match(r"rId(\d+)$", rel.get("Id") or "")
        if m:
            mx = max(mx, int(m.group(1)))
    return mx


# -----------------------------------------------------------------------------
# 4. 自动映射
# -----------------------------------------------------------------------------

def auto_map(tpl, target, prefer_cover=None):
    """给每张幻灯片推断目标版式 index。

    判定顺序 —— **结构优先，文字量与标题关键词只做兜底**（与 SKILL.md 的映射表一致）：

      0. 第 1 页 → 封面；末页 → 封底
      1. 标题关键词：目录/agenda/contents → 目录；thank/封底/Q&A → 封底
      2. 结构性判定（看**占位符槽位**，不看文字多少）：
         | 结构特征                                    | 判定       |
         | 有 ctrTitle                                  | 封面       |
         | 无 title + body 成组(≥4, idx≥17) + 每行很短   | 目录       |
         | title + body 成组 + 每行很短 + 合计很短       | 目录       |
         | 无 title + body + 合计很短                    | 章节过渡   |
         | 无占位符，且**原版式本身也没有占位符**         | 空白页     |
      3. 兜底：文字很少 → 仅标题；其余 → 标题+内容。

    **为什么必须有第 2 档**（2026-09-18 Public Wi-Fi 实测）：旧版只按"文字长度"猜，
    把"无 title、两个 body 槽位(17/22)、合计 12 字符"的章节过渡页判成 `title_only`，
    于是注入后章节页拿不到 `Section Title` 版式的 `01` 序号位与色带，整页退化成裸文字。
    `KIND_SECTION` 一直是**定义了却从不产出**的常量 —— 模板版式能正确归类成 section，
    映射器却永远不会选它。SKILL.md 的映射表本来就写着按占位符结构判，是代码没实现。
    """
    layouts = tpl["layouts"]
    by_kind = {}
    for L in layouts:
        by_kind.setdefault(L["kind"], []).append((L["index"], L["name"]))

    none = [(None, None)]
    cover = pick_cover(by_kind.get(KIND_COVER) or [], prefer_cover)
    toc = (by_kind.get(KIND_TOC) or none)[0][0]
    content = (by_kind.get(KIND_CONTENT) or none)[0][0]
    title_only = (by_kind.get(KIND_TITLE_ONLY) or none)[0][0] or content
    back = (by_kind.get(KIND_BACK) or none)[0][0]
    blank = (by_kind.get(KIND_BLANK) or none)[0][0]
    section = (by_kind.get(KIND_SECTION) or none)[0][0] or title_only

    slides = target["slides"]
    n = len(slides)
    mapping = OrderedDict()
    reasons = OrderedDict()
    for i, s in enumerate(slides, 1):
        t = s["title"] or ""
        phs = s["placeholders"] or []
        n_ctr = sum(1 for p in phs if p["type"] == "ctrTitle")
        n_title = sum(1 for p in phs if p["type"] in ("title", "ctrTitle"))
        bodies = [p for p in phs if p["type"] not in ("title", "ctrTitle")]
        body_idx = sorted(int(p["idx"]) for p in bodies if (p["idx"] or "").isdigit())
        body_text = [p["text"] for p in bodies if p["text"]]
        body_chars = s["text_len"] - len(t) if n_title else s["text_len"]
        grouped = (len(body_idx) >= BODY_GROUP_MIN
                   and body_idx[0] >= BODY_GROUP_IDX0)
        lines_short = all(len(x) <= TOC_LINE_MAX for x in body_text)

        def put(idx, why):
            mapping[i] = idx
            reasons[i] = why

        # --- 0. 首末页 ---
        if i == 1:
            put(cover, "首页→封面")
            continue
        if i == n and n > 1:
            put(back or blank or content, "末页→封底")
            continue

        # --- 1. 标题关键词 ---
        if re.search(r"目录|agenda|contents|大纲|catalog", t, re.I):
            put(toc or content, "标题含目录关键词")
            continue
        if re.search(r"thank|谢谢|封底|Q\s*&\s*A|Q&A", t, re.I):
            put(back or blank or content, "标题含封底关键词")
            continue

        # --- 2. 结构性判定（占位符槽位） ---
        if n_ctr:
            put(cover, "有 ctrTitle→封面")
            continue
        if n_title == 0 and grouped and lines_short and body_chars <= TOC_CHARS_MAX:
            put(toc or content, "无 title+body 成组(%d)且每行≤%d字→目录"
                % (len(body_idx), TOC_LINE_MAX))
            continue
        if n_title and grouped and lines_short and body_chars <= TOC_CHARS_MAX:
            put(toc or content, "title+body 成组且每行≤%d字→目录" % TOC_LINE_MAX)
            continue
        if n_title == 0 and bodies and body_chars <= SECTION_CHARS_MAX:
            put(section, "无 title+%d 个 body+合计 %d 字→章节过渡"
                % (len(bodies), body_chars))
            continue
        if (not phs and s.get("current_layout_ph") == 0
                and s.get("current_layout")):
            put(blank or content, "无占位符且原版式「%s」也无占位符→空白页(满版设计)"
                % (s.get("current_layout_name") or "?"))
            continue

        # --- 3. 兜底 ---
        if s["text_len"] <= len(t) + 8 and s["text_len"] < 60:
            put(title_only, "文字极少→仅标题")
            continue
        put(content, "标题+内容")

    return mapping, {"cover": cover, "toc": toc, "content": content,
                     "title_only": title_only, "back": back, "blank": blank,
                     "section": section, "reasons": reasons}


# -----------------------------------------------------------------------------
# 4.5 模板定位（模板随包内嵌，命令不写绝对路径）
# -----------------------------------------------------------------------------

SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = SKILL_DIR / "assets" / "master"

# 别名 → 包内模板文件名。包内只放这两个 Reyee 模板；换模板时同步更新这里。
TEMPLATE_ALIASES = {
    "default": "2025 Ruijie Reyee PPT Template-20250530.thmx",          # 浅色底
    "light": "2025 Ruijie Reyee PPT Template-20250530.thmx",
    "dark": "2025 Ruijie Reyee PPT Template-Dark-20250620.thmx",        # 深色底
}


def bundled_templates():
    """包内 assets/master/ 下的模板文件（按名排序）。"""
    if not TEMPLATE_DIR.is_dir():
        return []
    exts = (".thmx", ".potx", ".pptx")
    return sorted(p for p in TEMPLATE_DIR.iterdir()
                  if p.is_file() and p.suffix.lower() in exts)


def resolve_theme(spec):
    """把 --theme 的取值解析成实际模板路径。

    接受四种写法，都在**包内**解析，不依赖任何绝对路径：
      * 别名            --theme default（浅色，等同 light）／--theme dark
      * 裸文件名        --theme "2025 Ruijie Reyee PPT Template-Dark-20250620.thmx"
      * 唯一子串        --theme dark 之外，如 --theme 20250530
      * 真实路径        --theme D:/somewhere/other.thmx（外部模板，仍支持）

    这份逻辑存在的理由：模板文件虽然随包分发，但命令里若写绝对路径，
    发到别人机器上就找不到 —— 文件"在包里"不等于"能用包里的"。
    """
    if spec is None or str(spec).strip() == "":
        spec = "default"
    s = str(spec).strip()

    # ① 真实路径（外部模板 / 显式指定）
    p = Path(s)
    if p.exists():
        return p

    key = s.lower()
    avail = bundled_templates()

    # ② 包内裸文件名
    for c in avail:
        if c.name.lower() == key:
            return c

    # ③ 别名
    if key in TEMPLATE_ALIASES:
        c = TEMPLATE_DIR / TEMPLATE_ALIASES[key]
        if c.is_file():
            return c

    # ④ 唯一子串（只记得 "20250620" 这种片段时）
    hits = [c for c in avail if key in c.name.lower()]
    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        raise SystemExit("模板名不唯一，匹配到 %d 个：\n  %s" % (
            len(hits), "\n  ".join(c.name for c in hits)))

    # ⑤ 失败：把能用的都列出来
    msg = ["找不到模板：%s" % spec]
    if avail:
        msg.append("包内可用模板（%s）：" % TEMPLATE_DIR)
        msg += ["  - %s" % c.name for c in avail]
        msg.append("也可直接用别名：default / light（浅色），dark（深色）")
    else:
        msg.append("模板目录不存在或为空：%s" % TEMPLATE_DIR)
    raise SystemExit("\n".join(msg))


# -----------------------------------------------------------------------------
# 5. 主流程
# -----------------------------------------------------------------------------

def run(args):
    src = Path(args.source)
    tpl_path = resolve_theme(getattr(args, "theme", None))
    out = Path(args.out) if args.out else src.with_name(src.stem + "_tpl.pptx")
    prefix = args.prefix or "tpl"

    if not src.exists():
        raise SystemExit("源文件不存在：%s" % src)

    zt = zipfile.ZipFile(tpl_path)
    tpl_base = template_base(zt.namelist())
    tpl = parse_template(zt, tpl_base)
    tpl["_all_names"] = zt.namelist()

    path_map = build_path_map(tpl, prefix)

    zs = zipfile.ZipFile(src)
    target = parse_target(zs)

    report = {
        "source": str(src),
        "theme": str(tpl_path),
        "theme_base": tpl_base,
        "prefix": prefix,
        "template_layouts": tpl["layouts"],
        "template_masters": tpl["masters"],
        "target_slide_count": len(target["slides"]),
        "target_size": target["size"],
        "template_size": None,
        "path_map": path_map,
        "rels_changes": {},
        "content_type_added": [],
        "slides_remapped": [],
        "warnings": [],
        "orphan_placeholders": [],
    }

    # 模板画布
    pres_t = read_xml(zt.read(tpl_base + "presentation.xml"))
    sz_t = pres_t.find(qn("p:sldSz"))
    if sz_t is not None:
        report["template_size"] = (int(sz_t.get("cx")), int(sz_t.get("cy")))

    if report["target_size"] and report["template_size"] and report["target_size"] != report["template_size"]:
        report["warnings"].append(
            "画布尺寸不一致：源 %s vs 模板 %s。%s" % (
                report["target_size"], report["template_size"],
                "已按 --sync-size 同步" if args.sync_size else
                "建议加 --sync-size 同步，否则版式对不齐（内容位置会错乱）"))

    # ---- outline 模式：只出清单和自动映射建议
    if args.outline:
        mapping, auto = auto_map(tpl, target, args.cover_layout)
        report["auto_map"] = {"suggest": {str(k): v for k, v in mapping.items()}, "kinds": auto}
        print_outline(tpl, target, mapping, auto)
        return report, out, None

    # ---- 映射
    if args.map:
        mapping = OrderedDict()
        for piece in args.map.split(","):
            k, _, v = piece.partition(":")
            mapping[int(k.strip())] = int(v.strip())
        auto = {"source": "manual"}
    else:
        mapping, auto = auto_map(tpl, target, args.cover_layout)
        auto = {"source": "auto", **auto}
    report["auto_map"] = {"mapping": {str(k): v for k, v in mapping.items()}, **auto}

    # 版式 index -> 模板部件 & 目标部件
    by_index = {L["index"]: L for L in tpl["layouts"]}
    layout_ph = {}
    for L in tpl["layouts"]:
        root = read_xml(zt.read(L["part"]))
        phs = []
        for ph in root.iter(qn("p:ph")):
            phs.append((ph.get("type") or "body", ph.get("idx")))
        layout_ph[L["index"]] = phs

    # ---- 预演：只算不写
    for i, s in enumerate(target["slides"], 1):
        want = mapping.get(i)
        if want is None or want not in by_index:
            report["warnings"].append("第 %d 页未找到可用版式（map 值=%s）" % (i, want))
            continue
        L = by_index[want]
        # 占位符承接检查
        want_phs = layout_ph[want]
        want_types = {t for t, _ in want_phs}
        want_pairs = set(want_phs)
        orphans = []
        for ph in s["placeholders"]:
            t, idx = ph["type"], ph["idx"]
            if (t, idx) in want_pairs:
                continue
            # title / ctrTitle 互相兼容
            if t in ("title", "ctrTitle") and ({"title", "ctrTitle"} & want_types):
                continue
            if t == "body" and idx not in {i2 for t2, i2 in want_phs if t2 == "body"}:
                orphans.append(ph)
        rec = {
            "slide": i,
            "title": s["title"],
            "from_layout": s["current_layout"],
            "to_layout_index": want,
            "to_layout_name": L["name"],
            "to_layout_part": path_map.get(L["part"], L["part"]),
            "to_layout_name_cn": L["kind"],
            "placeholders_before": s["placeholders"],
            "placeholders_after": [{"type": t, "idx": ix} for t, ix in want_phs],
            "orphan_placeholders": orphans,
        }
        if s["current_layout"] == path_map.get(L["part"], L["part"]):
            rec["noop"] = True
        report["slides_remapped"].append(rec)
        if orphans:
            report["orphan_placeholders"].append({"slide": i, "orphans": orphans})

    if report["orphan_placeholders"]:
        report["warnings"].append(
            "有 %d 页存在无法承接的占位符，内容可能孤悬或丢失，必须逐页处理" % len(report["orphan_placeholders"]))

    if args.dry_run:
        return report, out, None

    # ---- 实跑：重建 zip
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".building.pptx")
    src_names = set(zs.namelist())
    slide_rels_parts = {s["part"].replace("slides/", "slides/_rels/") + ".rels"
                        for s in target["slides"]}

    # 旧母版清理：只有在所有页面都完成重映射后才允许删。
    # 否则页面会指向已删除的版式 → PowerPoint 报文件损坏。
    map_ok = not any("未找到可用版式" in w for w in report["warnings"])
    drop_parts = set()
    if args.drop_old_master and map_ok and target["masters"]:
        for part in list(target["masters"]) + list(target["layouts"]):
            drop_parts.add(part)
            if "/slideMasters/" in part:
                drop_parts.add(part.replace("/slideMasters/", "/slideMasters/_rels/") + ".rels")
            elif "/slideLayouts/" in part:
                drop_parts.add(part.replace("/slideLayouts/", "/slideLayouts/_rels/") + ".rels")
        report["dropped_old"] = sorted(drop_parts)
    elif args.drop_old_master and not map_ok:
        report["warnings"].append("有页面未能完成重映射，已放弃删除旧母版/版式（避免文件损坏）")
    else:
        report["warnings"].append(
            "包内仍保留旧母版/旧版式（%d 个），PowerPoint 的版式列表会同时出现两套；"
            "要清理请加 --drop-old-master" % len(target["layouts"]))

    # 旧主题回收（陷阱「删了旧母版，却把演示文稿级主题留在旧母版的主题上」）
    # 删掉旧母版/旧版式后，`presentation.xml.rels` 的**演示文稿级主题**往往还指着旧主题。
    # 后果：包看着换过模板，主题库里却仍是旧模板的血统 —— 新建形状/取色/取字体全继承旧的，
    # 而 theme 被"旧母版 + presentation"双引用时引用计数=2，孤儿扫描又不敢删。
    # 所以判定必须按**血统**（这个 theme 挂在被删母版下），不能按引用计数。
    tpl_theme_dst = next((dst for part, dst in path_map.items()
                          if "/theme/" in part and part.endswith(".xml")
                          and "themanag" not in part.lower()), None)
    pres_rels_x = read_xml(zs.read("ppt/_rels/presentation.xml.rels"))
    pres_theme_rids = [r.get("Id") for r in pres_rels_x
                       if (r.get("Type") or "").endswith("/theme")]
    dead_themes = set()
    if drop_parts and tpl_theme_dst:
        for r in pres_rels_x:
            if not (r.get("Type") or "").endswith("/theme"):
                continue
            tgt = resolve_target("ppt/presentation.xml", r.get("Target"))
            if tgt and tgt != tpl_theme_dst:
                dead_themes.add(tgt)
        # 被删母版自己引用的主题也要算进来
        for part in list(drop_parts):
            if "/_rels/" in part or not part.endswith(".xml"):
                continue
            rp = posixpath.join(posixpath.dirname(part), "_rels",
                                posixpath.basename(part) + ".rels")
            if rp not in src_names:
                continue
            for r in read_xml(zs.read(rp)):
                if (r.get("Type") or "").endswith("/theme"):
                    tgt = resolve_target(part, r.get("Target"))
                    if tgt and tgt != tpl_theme_dst:
                        dead_themes.add(tgt)
        # 只回收"改指之后没有任何存活部件再引用"的主题
        for th in sorted(dead_themes):
            still = False
            for rp in [x for x in src_names if x.endswith(".rels")
                       and x not in drop_parts]:
                if rp == "ppt/_rels/presentation.xml.rels":
                    continue                     # 这一步就是要去改指它
                host = posixpath.dirname(posixpath.dirname(rp))
                host = posixpath.join(host, posixpath.basename(rp)[:-5])
                for r in read_xml(zs.read(rp)):
                    if (r.get("Type") or "").endswith("/theme"):
                        if resolve_target(host, r.get("Target")) == th:
                            still = True
                            break
                if still:
                    break
            if not still:
                drop_parts.add(th)
                report.setdefault("dropped_old_themes", []).append(th)
        report["theme_retarget"] = {"from": sorted(dead_themes), "to": tpl_theme_dst}

    # ---- 0. [Content_Types].xml 必须**先算好并以第一个条目写入**：
    #        OPC 规范要求 content-types 流是包内第一个部件，PowerPoint 严格执行。
    #        实测把 CT 留到最后写（旧实现），PowerPoint 直接拒收 —— 而 python-pptx
    #        能打开、verify_pptx.py 也报 P0=0，是典型的假阴性。
    #        （旧流程之所以没暴露：normalize_deck.py 用 python-pptx 重新保存整包，
    #          顺带把 CT 排回了首位。）
    path_map_written = set()
    for tpl_part, dst in path_map.items():
        if tpl_part not in src_names and tpl_part not in zt.namelist():
            continue
        path_map_written.add(dst)
    ct = read_xml(zs.read("[Content_Types].xml"))
    if drop_parts:
        for ov in list(ct.findall(qn("ct:Override"))):
            pn = (ov.get("PartName") or "").lstrip("/")
            if pn in drop_parts:
                ct.remove(ov)
    ct_added = []
    for tpl_part, dst in path_map.items():
        if dst not in path_map_written:
            continue
        ext = dst.rsplit(".", 1)[-1].lower()
        base = posixpath.basename(dst)
        d = posixpath.dirname(dst)

        # ---- .rels 部件的 content-type 恒为 relationships+xml（走 Default rels）。
        #      绝不能落进下面按目录名的分支：版式/母版的伴生 rels 文件
        #      （ppt/slideLayouts/_rels/tplSlideLayout1.xml.rels）路径里同样含
        #      "/slideLayouts/"，会被误写成 slideLayout+xml。这是硬性错误，
        #      PowerPoint 直接拒收（E_FAIL），而 python-pptx 从不校验 content-type，
        #      所以能照常打开 —— 典型假阴性。用 dirname 全等判定 + 本分支双保险。
        if base.endswith(".rels"):
            if ensure_default(ct, "rels", DEFAULT_EXT_CT["rels"]):
                ct_added.append("Default:rels")
            continue

        # 注意：一律用 dirname 全等，不用子串包含 —— 子串会命中 _rels 子目录。
        if d == "ppt/slideMasters":
            if add_content_type(ct, dst, CT["slideMaster"]):
                ct_added.append(dst)
        elif d == "ppt/slideLayouts":
            if add_content_type(ct, dst, CT["slideLayout"]):
                ct_added.append(dst)
        elif d == "ppt/theme" and base.endswith(".xml") and "themeManager" not in base:
            if add_content_type(ct, dst, CT["theme"]):
                ct_added.append(dst)
        elif d == "ppt/tags":
            if add_content_type(ct, dst, CT["tags"]):
                ct_added.append(dst)
        else:
            if ext in DEFAULT_EXT_CT and ensure_default(ct, ext, DEFAULT_EXT_CT[ext]):
                ct_added.append("Default:%s" % ext)
    report["content_type_added"] = ct_added

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        written = set()
        # ---- 0. 第一个条目永远是 [Content_Types].xml
        zo.writestr("[Content_Types].xml", to_bytes(ct))
        # ---- 1. 目标原有部件（跳过要重建的 / 要删除的）
        skip = {"[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"}
        skip |= slide_rels_parts
        skip |= drop_parts
        for n in zs.namelist():
            if n in skip:
                continue
            zo.writestr(n, zs.read(n))
            written.add(n)

        # ---- 2. 注入模板部件
        for tpl_part, dst in path_map.items():
            if tpl_part not in src_names and tpl_part not in zt.namelist():
                continue
            blob = zt.read(tpl_part)
            if tpl_part.endswith(".rels"):
                host_tpl = tpl_part[:-5].replace("/_rels/", "/")
                host_dst = path_map.get(host_tpl, host_tpl)
                blob, changes = rewrite_rels(blob, host_tpl, host_dst, path_map)
                if changes:
                    report["rels_changes"][dst] = changes
            # 若目标已有同名部件，用模板覆盖（注入优先）
            zo.writestr(dst, blob)
            written.add(dst)

        # ---- 3. 重写目标 slides 的 rels（把 sldLayout 指向新模板版式）
        for s in target["slides"]:
            i = s["index"]
            rels_part = s["part"].replace("slides/", "slides/_rels/") + ".rels"
            if rels_part in zs.namelist():
                root = read_xml(zs.read(rels_part))
            else:
                root = etree.Element(qn("rel:Relationships"), nsmap={None: NS["rel"]})
            rec = next((r for r in report["slides_remapped"] if r["slide"] == i), None)
            if rec is None:
                zo.writestr(rels_part, to_bytes(root))
                written.add(rels_part)
                continue
            tpl_layout_part = next(L["part"] for L in tpl["layouts"]
                                   if L["index"] == rec["to_layout_index"])
            dst_layout = path_map.get(tpl_layout_part, tpl_layout_part)
            target_rel = relativize(s["part"], dst_layout)
            done = False
            for rel in root:
                if rel.get("Type", "").endswith("/slideLayout"):
                    rel.set("Target", target_rel)
                    done = True
            if not done:
                rid = "rId%d" % (max_rid(root) + 1)
                el = etree.SubElement(root, qn("rel:Relationship"))
                el.set("Id", rid)
                el.set("Type", REL_T + "slideLayout")
                el.set("Target", target_rel)
            zo.writestr(rels_part, to_bytes(root))
            written.add(rels_part)

        # ---- 4. presentation.xml.rels：移除旧母版关系，加新 slideMaster 关系
        pres_rels = read_xml(zs.read("ppt/_rels/presentation.xml.rels"))
        old_master_rids = set()
        if drop_parts:
            for rel in list(pres_rels):
                if rel.get("Type", "").endswith("/slideMaster"):
                    tgt = resolve_target("ppt/presentation.xml", rel.get("Target"))
                    if tgt in drop_parts:
                        old_master_rids.add(rel.get("Id"))
                        pres_rels.remove(rel)
        # 演示文稿级主题改指到模板主题（与上面 drop_parts 的回收配套）
        if drop_parts and tpl_theme_dst:
            for rel in pres_rels:
                if (rel.get("Type") or "").endswith("/theme"):
                    rel.set("Target", relativize("ppt/presentation.xml", tpl_theme_dst))
                    break
        master_dst = path_map.get(tpl["masters"][0]["part"], tpl["masters"][0]["part"])
        rid_new = "rId%d" % (max_rid(pres_rels) + 1)
        el = etree.SubElement(pres_rels, qn("rel:Relationship"))
        el.set("Id", rid_new)
        el.set("Type", REL_T + "slideMaster")
        el.set("Target", relativize("ppt/presentation.xml", master_dst))

        # ---- 嵌入字体关系（可选）
        font_rids = {}
        if args.with_fonts:
            fonts = sorted(x for x in path_map if "/fonts/" in x)
            for fp in fonts:
                rid = "rId%d" % (max_rid(pres_rels) + 1)
                fel = etree.SubElement(pres_rels, qn("rel:Relationship"))
                fel.set("Id", rid)
                fel.set("Type", REL_T + "font")
                fel.set("Target", relativize("ppt/presentation.xml", path_map[fp]))
                font_rids[posixpath.basename(path_map[fp])] = rid
        zo.writestr("ppt/_rels/presentation.xml.rels", to_bytes(pres_rels))
        written.add("ppt/_rels/presentation.xml.rels")

        # ---- 5. presentation.xml：加 sldMasterId / 同步尺寸 / 嵌入字体
        pres = read_xml(zs.read("ppt/presentation.xml"))
        master_lst = pres.find(qn("p:sldMasterIdLst"))
        if master_lst is None:
            master_lst = etree.Element(qn("p:sldMasterIdLst"))
            pres.insert(0, master_lst)
        if old_master_rids:
            for e in list(master_lst):
                if e.get("{%s}id" % REL_NS) in old_master_rids:
                    master_lst.remove(e)
        existing_ids = [int(e.get("id")) for e in master_lst if e.get("id")]
        new_mid = (max(existing_ids) + 1) if existing_ids else 2147483887
        mid = etree.SubElement(master_lst, qn("p:sldMasterId"))
        mid.set("id", str(new_mid))
        mid.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", rid_new)

        if args.sync_size and report["template_size"]:
            sz = pres.find(qn("p:sldSz"))
            if sz is not None:
                sz.set("cx", str(report["template_size"][0]))
                sz.set("cy", str(report["template_size"][1]))
                report["warnings"].append("已同步画布尺寸为 %s" % (report["template_size"],))

        if args.with_fonts and font_rids:
            tpl_pres = read_xml(zt.read(tpl_base + "presentation.xml"))
            efl_t = tpl_pres.find(qn("p:embeddedFontLst"))
            if efl_t is not None:
                old = pres.find(qn("p:embeddedFontLst"))
                if old is not None:
                    pres.remove(old)
                new_efl = copy.deepcopy(efl_t)
                # 重映射字体 rId：按顺序对齐
                tpl_font_rids = {}
                for fnt in new_efl:
                    for child in fnt:
                        rid = child.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                        if rid:
                            tpl_font_rids[child.get("Tag") or child.tag] = rid
                # 用模板 rels 顺序映射
                tpl_rels = read_xml(zt.read(tpl_base + "_rels/presentation.xml.rels"))
                tpl_map = {}
                for rel in tpl_rels:
                    if rel.get("Type", "").endswith("/font"):
                        tpl_map[rel.get("Id")] = posixpath.basename(rel.get("Target"))
                # 逐个字体项替换 rId
                attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                for fnt in new_efl:
                    for child in fnt:
                        old_rid = child.get(attr)
                        if old_rid and old_rid in tpl_map:
                            fname = tpl_map[old_rid]
                            new_rid = font_rids.get(prefix + "_" + fname) or font_rids.get(fname)
                            if new_rid:
                                child.set(attr, new_rid)
                # 插入位置 —— OOXML (CT_Presentation) 对 <p:presentation> 子元素顺序有硬要求：
                #   sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, sldIdLst,
                #   sldSz, notesSz, smartTags, embeddedFontLst,
                #   custShowLst, photoAlbum, custDataLst, kinsoku, defaultTextStyle,
                #   modifyVerifier, extLst
                # 即 embeddedFontLst 必须排在 custShowLst / photoAlbum / custDataLst / kinsoku /
                # defaultTextStyle 之前。实测旧实现只找 defaultTextStyle，
                # 当包里已有 custDataLst（think-cell / 标记标签常见）时会被插到它后面 → P0。
                AFTER_EFL = ("custShowLst", "photoAlbum", "custDataLst", "kinsoku",
                             "defaultTextStyle", "modifyVerifier", "extLst")
                anchor = None
                for child in pres:
                    if not isinstance(child.tag, str):
                        continue
                    if _local(child.tag) in AFTER_EFL:
                        anchor = child
                        break
                if anchor is not None:
                    anchor.addprevious(new_efl)
                else:
                    pres.append(new_efl)
                report["embedded_fonts_injected"] = len(font_rids)

        zo.writestr("ppt/presentation.xml", to_bytes(pres))
        written.add("ppt/presentation.xml")

        # ---- 6. [Content_Types].xml 已在第 0 步作为第一个条目写入（见上）。

    shutil.move(str(tmp), str(out))

    # ---- 自检：PowerPoint 会拒收的包级错误，python-pptx 一律照收（假阴性）。
    #      历史教训：模板块的伴生 .rels 被判成 slideLayout+xml，verify_pptx.py 报 P0=0、
    #      python-pptx 能开，但 PowerPoint 直接 E_FAIL。这里做硬校验，不通过就抛。
    problems = _package_self_check(out)
    if problems:
        report["warnings"].extend(problems)
        raise SystemExit(
            "自检失败 —— 产物会被 PowerPoint 拒收，已阻断：\n  - " + "\n  - ".join(problems))
    return report, out, None


RELS_CT = "application/vnd.openxmlformats-package.relationships+xml"


def _package_self_check(path):
    """包级硬校验。返回问题列表（空 = 通过）。

    覆盖 PowerPoint 会硬拒收、而 python-pptx 毫无察觉的几类错误：
      1. [Content_Types].xml 不是 ZIP 首条目（OPC 强制）
      2. 重复条目名
      3. .rels 部件的 content-type 不是 relationships+xml  ← 本文件历史 bug
      4. 有部件拿不到 content-type（Default/Override 都覆盖不到）
      5. 关系 Target / 宿主引用的 r:id 悬空
    """
    import posixpath as _pp

    CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
    out = []
    z = zipfile.ZipFile(path)
    names = z.namelist()

    if not names or names[0] != "[Content_Types].xml":
        out.append("ZIP 首条目不是 [Content_Types].xml，而是 %r" % (names[0] if names else None))

    seen, dups = set(), []
    for n in names:
        if n in seen:
            dups.append(n)
        seen.add(n)
    if dups:
        out.append("重复条目名：%s" % dups[:5])

    ctr = read_xml(z.read("[Content_Types].xml"))
    defaults, overrides = {}, {}
    for ch in ctr:
        t = _local(ch.tag)
        if t == "Default":
            defaults[(ch.get("Extension") or "").lower()] = ch.get("ContentType")
        elif t == "Override":
            overrides[(ch.get("PartName") or "").lstrip("/")] = ch.get("ContentType")

    for pn, ctype in sorted(overrides.items()):
        if pn.endswith(".rels") and ctype != RELS_CT:
            out.append("rels 部件 content-type 错误：%s => %s（应为 relationships+xml）" % (pn, ctype))

    for n in names:
        if n.endswith("/") or n == "[Content_Types].xml":
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if n not in overrides and ext not in defaults:
            out.append("部件没有任何 content-type：%s" % n)

    # 关系与 r:id 完整性
    rels_of = {}
    for n in names:
        if not n.endswith(".rels"):
            continue
        if n == "_rels/.rels":
            host = ""
        else:
            d, f = _pp.split(n)
            host = _pp.join(_pp.dirname(d), f[:-5])
        try:
            rx = read_xml(z.read(n))
        except Exception as e:
            out.append("rels 解析失败 %s: %s" % (n, e))
            continue
        m = {}
        for rel in rx:
            if (rel.get("TargetMode") or "Internal") == "External":
                continue
            resolved = _pp.normpath(_pp.join(_pp.dirname(host), rel.get("Target") or ""))
            if resolved not in seen:
                out.append("悬空关系：%s %s -> %s" % (n, rel.get("Id"), rel.get("Target")))
            m[rel.get("Id")] = resolved
        rels_of[host] = m

    for n in names:
        if not n.endswith(".xml"):
            continue
        if not (n.startswith("ppt/slides/slide") or n.startswith("ppt/slideLayouts")
                or n.startswith("ppt/slideMasters") or n == "ppt/presentation.xml"):
            continue
        blob = z.read(n)
        used = {u.decode() for u in re.findall(rb'r:(?:id|embed|link|pict|dm)="([^"]+)"', blob)}
        miss = used - set(rels_of.get(n, {}).keys())
        if miss:
            out.append("%s 引用了未声明的 rId：%s" % (n, sorted(miss)))

    z.close()
    return out


def print_outline(tpl, target, mapping, auto):
    print("=" * 78)
    print("模板版式清单（%d 个）" % len(tpl["layouts"]))
    print("=" * 78)
    print("%-5s %-22s %-14s %s" % ("idx", "名称", "用途", "占位符类型"))
    print("-" * 78)
    for L in tpl["layouts"]:
        print("%-5d %-22s %-14s %s" % (
            L["index"], L["name"][:22], L["kind"], ",".join(L["ph_types"]) or "-"))
    print()
    print("用途归类：%s" % {k: v for k, v in auto.items()
                            if k not in ("source", "reasons")})
    print()
    print("=" * 78)
    print("目标文件：%d 页" % len(target["slides"]))
    print("=" * 78)
    print("%-4s %-40s %-6s %-18s %s" % ("页", "现有标题", "→版式", "版式名", "判定依据"))
    print("-" * 78)
    reasons = auto.get("reasons") or {}
    for s in target["slides"]:
        want = mapping.get(s["index"])
        nm = next((L["name"] for L in tpl["layouts"] if L["index"] == want), "?")
        print("%-4d %-40s %-6s %-18s %s" % (
            s["index"], (s["title"] or "(无标题)")[:40], want, nm[:18],
            reasons.get(s["index"], "")))
    print()
    print("用 --map \"1:2,3:12,...\" 覆盖自动映射；用 --dry-run 预演。")
    print("判定依据一栏是逐页可审计的：结构判据命中时会写明命中哪一条。")


def to_md(rep):
    L = []
    L.append("# 模板注入报告\n")
    L.append("- 源文件：`%s`" % rep["source"])
    L.append("- 模板：`%s`（基准目录 `%s`）" % (rep["theme"], rep["theme_base"]))
    L.append("- 部件前缀：`%s`" % rep["prefix"])
    L.append("- 源画布：%s　模板画布：%s" % (rep["target_size"], rep["template_size"]))
    L.append("- 注入模板版式：%d 个" % len(rep["template_layouts"]))
    L.append("")

    if rep["warnings"]:
        L.append("## 警告\n")
        for w in rep["warnings"]:
            L.append("- **%s**" % w)
        L.append("")

    L.append("## 逐页重映射\n")
    L.append("| 页 | 现有标题 | 原版式 | → 新版式 | 新版式占位符 | 孤儿占位符 |")
    L.append("|---|---|---|---|---|---|")
    for r in rep["slides_remapped"]:
        pa = ", ".join("%s%s" % (p["type"], "" if p["idx"] in (None, "0") else "#" + p["idx"])
                       for p in r["placeholders_after"]) or "-"
        orph = ", ".join("%s%s" % (p.get("type"), "" if p.get("idx") in (None, "0") else "#" + str(p.get("idx")))
                         for p in r["orphan_placeholders"]) or "-"
        L.append("| %d | %s | %s | **%d** %s | %s | %s |" % (
            r["slide"], (r["title"] or "-")[:40],
            posixpath.basename(r["from_layout"] or "-"),
            r["to_layout_index"], r["to_layout_name"],
            pa, orph))
    L.append("")

    if rep["orphan_placeholders"]:
        L.append("## 孤儿占位符（必须逐页处理）\n")
        L.append("这些页面的内容占位符在新版式里找不到对应槽位，PowerPoint 打开后内容会跑到空白区或丢失：\n")
        for o in rep["orphan_placeholders"]:
            L.append("### 第 %d 页" % o["slide"])
            for p in o["orphans"]:
                L.append("- `%s` idx=%s" % (p.get("type"), p.get("idx")))
            L.append("")

    L.append("## 版式清单\n")
    L.append("| idx | 名称 | 用途 | 占位符 |")
    L.append("|---|---|---|---|")
    for L2 in rep["template_layouts"]:
        L.append("| %d | %s | %s | %s |" % (
            L2["index"], L2["name"], L2["kind"], ", ".join(L2["ph_types"]) or "-"))
    L.append("")

    L.append("## 部件映射\n")
    L.append("| 模板内路径 | → 目标路径 |")
    L.append("|---|---|")
    for k, v in sorted(rep["path_map"].items()):
        L.append("| `%s` | `%s` |" % (k, v))
    L.append("")

    if rep["content_type_added"]:
        L.append("## Content_Types 新增\n")
        for c in rep["content_type_added"]:
            L.append("- `%s`" % c)
        L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="把模板母版体系注入到目标 pptx")
    ap.add_argument("--source", required=True, help="目标 pptx")
    ap.add_argument("--theme", default="default",
                    help="模板：别名 default/light（浅色）、dark（深色），"
                         "或包内模板文件名，或外部 .thmx/.pptx 路径（默认 default）")
    ap.add_argument("--out", default=None, help="输出 pptx")
    ap.add_argument("--prefix", default="tpl", help="注入部件的前缀（默认 tpl）")
    ap.add_argument("--map", default=None, help='页→版式映射，如 "1:1,2:10,3:12,5:16"')
    ap.add_argument("--cover-layout", type=int, default=None, help="封面用哪个版式 index")
    ap.add_argument("--outline", action="store_true", help="只输出版式清单与自动映射建议")
    ap.add_argument("--dry-run", action="store_true", help="预演，不写文件")
    ap.add_argument("--sync-size", action="store_true", help="把画布尺寸同步为模板的")
    ap.add_argument("--drop-old-master", dest="drop_old_master", action="store_true", default=True,
                    help="删除目标原有的旧母版与旧版式，只留模板版式（默认开）")
    ap.add_argument("--keep-old-master", dest="drop_old_master", action="store_false",
                    help="保留旧母版/旧版式（版式列表会同时出现两套）")
    ap.add_argument("--with-fonts", dest="with_fonts", action="store_true", default=True,
                    help="注入嵌入字体（默认开）")
    ap.add_argument("--no-fonts", dest="with_fonts", action="store_false")
    ap.add_argument("--report", default=None, help="报告输出目录")
    args = ap.parse_args()

    rep, out, _ = run(args)
    if args.outline:
        return

    rp = Path(args.report) if args.report else out.with_suffix("")
    rp.mkdir(parents=True, exist_ok=True)
    (rp / "inject-report.json").write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    (rp / "inject-report.md").write_text(to_md(rep), encoding="utf-8")

    print("=" * 78)
    print("注入%s完成" % ("预演" if args.dry_run else ""))
    print("=" * 78)
    print("模板版式：%d 个　目标页数：%d" % (len(rep["template_layouts"]), rep["target_slide_count"]))
    print("部件映射：%d 条" % len(rep["path_map"]))
    print()
    print("%-4s %-40s %-10s %s" % ("页", "标题", "→版式", "孤儿占位符"))
    print("-" * 78)
    for r in rep["slides_remapped"]:
        orph = ",".join("%s#%s" % (p.get("type"), p.get("idx")) for p in r["orphan_placeholders"]) or "-"
        print("%-4d %-40s %-10s %s" % (r["slide"], (r["title"] or "-")[:40],
                                       "%d(%s)" % (r["to_layout_index"], r["to_layout_name"][:8]), orph))
    if rep["warnings"]:
        print()
        print("警告：")
        for w in rep["warnings"]:
            print("  ! %s" % w)
    if not args.dry_run:
        print()
        print("输出：%s" % out)
        print("报告：%s" % (rp / "inject-report.md"))


if __name__ == "__main__":
    main()
