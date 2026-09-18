# -*- coding: utf-8 -*-
"""把页面上"压过版式设定"的字体/字号覆盖清掉，让文字回归版式（→母版）的继承值。

解决什么问题
------------
原稿（或上游工具）会在**页面**上给占位符写死字体/字号：
  · 标题 sp 的 <a:lstStyle><a:lvl1pPr><a:defRPr sz="2600"><a:latin typeface="Arial"/>…
  · 标题段落里每个 <a:r><a:rPr sz="2400" b="1" latin="微软雅黑"/>…
这些覆盖的优先级**高于版式**，于是同一份 PPT 里标题会出现 26/28/24pt、
Arial/微软雅黑/Noto Sans 混着来 —— 表现为"标题上上下下、字体不对"。
版式本身（模板设定）通常是对的，用户要的就是"回归模板"。

本脚本按 OOXML 继承链处理：**删页面级覆盖，让版式（→母版）话说**。

继承链（从高到低，谁赢看谁在前面有值）
  ① slide 的 <a:lstStyle><a:lvl1pPr><a:defRPr>     ← 本脚本清掉
  ② slide 段落/运行的 <a:rPr> / <a:pPr><a:defRPr>  ← 本脚本清掉（只清字体字号粗斜）
  ③ layout 槽位的 <a:lstStyle><a:lvl1pPr><a:defRPr>  ← 模板设定，保留
  ④ master <p:txStyles> 同名 style                    ← 保留

清哪几个属性（只清"字体字号"这一类，不碰颜色）
  · 属性：sz, b, i, kern, spc, baseline
  · 子元素：a:latin, a:ea, a:cs, a:sym
  **保留 a:solidFill**（颜色不在本次范围；且留着的值本就是品牌色）

用法
----
    # 先体检（只读，报告每个占位符的覆盖项与生效值）
    python fix_title_style.py "<deck.pptx>" --check

    # 清标题覆盖
    python fix_title_style.py "<deck.pptx>" --out-deck "<new.pptx>" --scope title

    # 顺带把版式/母版标题色从模板残留的 043CC1 归一到品牌 0055CD
    python fix_title_style.py "<deck.pptx>" --out-deck "<new.pptx>" --scope title \
        --unify-title-color 0055CD

产出后**必须**跑 verify_pptx.py + open_test.py + 渲染复核（字号变化可能引起换行/溢出）。
"""
from __future__ import annotations

import argparse
import os
import posixpath
import re
import sys
import zipfile

from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'a': A, 'p': P, 'r': R}
CT_PART = '[Content_Types].xml'

# 要清掉的字体/字号类属性
FONT_ATTRS = ('sz', 'b', 'i', 'kern', 'spc', 'baseline', 'cap', 'strike', 'u')
FONT_TAGS = ('latin', 'ea', 'cs', 'sym')


def strip_font(rpr, log=None, prefix=''):
    """清掉一个 rPr/defRPr 上的字体字号覆盖；**保留 a:solidFill 等其余项**。

    只动"字体/字号/粗斜"这一类，颜色与段落属性一概不碰 —— 这样本步骤是
    **颜色中立**的：修完不影响任何元素的颜色（换色属 Step 4，不在本步骤范围）。
    """
    if rpr is None:
        return False
    hit = False
    for at in FONT_ATTRS:
        if rpr.get(at) is not None:
            if log is not None:
                log.append('%s%s=%s' % (prefix, at, rpr.get(at)))
            del rpr.attrib[at]
            hit = True
    for tag in FONT_TAGS:
        e = rpr.find('{%s}%s' % (A, tag))
        if e is not None:
            if log is not None:
                log.append('%s%s=%s' % (prefix, tag, e.get('typeface')))
            rpr.remove(e)
            hit = True
    return hit


def drop_lst_style(txBody, log=None):
    """（激进档，默认不用）整块丢弃 <a:lstStyle> 的内容。

    只清字体字号时**不要**用这个：lstStyle 里往往还带着该页的 fill（颜色），
    整块丢掉会把颜色一并交给版式，等于顺手改了颜色。"""
    ls = txBody.find('a:lstStyle', NS)
    if ls is None or len(ls) == 0:
        return False
    for lvl in ls:
        d = lvl.find('a:defRPr', NS)
        if d is not None and log is not None:
            c = d.find('a:solidFill/a:srgbClr', NS)
            if c is not None:
                log.append('lstStyle:丢弃fill=%s' % c.get('val'))
    for ch in list(ls):
        ls.remove(ch)
    return True


# 旧实现 text_width_pt() / fit_scale() / DEFAULT_FONT 已于 2026-09-18 删除：
# 它们把可用宽度与版式字号写死、字体路径写死到某台机器的用户字体目录，
# 已由 pptx_title_scope.py（度量 + 范围判定）+ fix_title_oneline.py（收口）取代。
# 上面的 --overflow-shrink 会直接返回退出码 2 并提示迁移到新脚本。



def process_slide(root, scope, log, drop_ls=False):
    """处理一页。scope: title / body / all"""
    changed = False
    for sp in root.findall('.//p:cSld/p:spTree/p:sp', NS):
        ph = sp.find('.//p:nvPr/p:ph', NS)
        if ph is None:
            continue
        typ = ph.get('type') or 'body'
        is_title = typ in ('title', 'ctrTitle')
        if scope == 'title' and not is_title:
            continue
        if scope == 'body' and is_title:
            continue
        if typ == 'sldNum':            # 页码不动
            continue
        tb = sp.find('p:txBody', NS)
        if tb is None:
            continue
        if drop_ls:
            if drop_lst_style(tb, log):
                changed = True
        else:
            # 默认（颜色中立）：只清字体字号；lstStyle 里的 fill / 段落属性保留
            # 注：iter() 只吃 tag，不吃路径 —— txBody 下所有 a:defRPr
            # （含 lstStyle 的与 pPr 的）都在这一条里覆盖
            for rpr in tb.iter('{%s}defRPr' % A):
                if strip_font(rpr, log, prefix='defRPr:'):
                    changed = True
        for rpr in tb.iter('{%s}rPr' % A):
            if strip_font(rpr, log):
                changed = True
        for rpr in tb.iter('{%s}endParaRPr' % A):
            if strip_font(rpr, log):
                changed = True
    return changed


def unify_color(root, old, new):
    """把 root 里所有 old 色值的 srgbClr 改成 new。"""
    n = 0
    for c in root.iter('{%s}srgbClr' % A):
        if (c.get('val') or '').upper() == old.upper():
            c.set('val', new)
            n += 1
    return n


def main():
    ap = argparse.ArgumentParser(description='清掉页面级的字体/字号覆盖，回归版式设定')
    ap.add_argument('deck')
    ap.add_argument('--check', action='store_true', help='只体检，不输出')
    ap.add_argument('--out-deck', default=None)
    ap.add_argument('--scope', default='title', choices=['title', 'body', 'all'])
    ap.add_argument('--unify-title-color', default=None,
                    help='把版式/母版/页面里标题残留的旧色统一到该 hex，如 0055CD')
    ap.add_argument('--from-color', default='043CC1',
                    help='要被替换掉的旧色（默认 043CC1，模板残留的标题蓝）')
    ap.add_argument('--allow-template-write', action='store_true',
                    help='允许把色归一写到版式/母版（模板本体）。默认**不允许** —— '
                         '改版式/母版就是改模板，不再是"用模板格式改 PPT"')
    ap.add_argument('--drop-lst-style', action='store_true',
                    help='激进档：整块丢弃 lstStyle（会连带把该页的 fill 交给版式，等于顺手改色）')
    ap.add_argument('--overflow-shrink', action='store_true',
                    help='【已弃用】标题过长缩字号 —— 请改用 fix_title_oneline.py。'
                         '本开关只认 type="title"、宽高写死 700.6pt/28pt，会漏掉'
                         '标题槽位（目录/分节名/序号）与内边距，导致缩了仍换行')
    args = ap.parse_args()

    z = zipfile.ZipFile(args.deck)
    names = set(z.namelist())

    reports = {}
    slide_new = {}
    for n in sorted(names, key=lambda s: int(m.group(1)) if (m := re.match(r'ppt/slides/slide(\d+)\.xml$', s)) else 0):
        if not re.match(r'ppt/slides/slide\d+\.xml$', n):
            continue
        root = etree.fromstring(z.read(n))
        log = []
        if process_slide(root, args.scope, log, drop_ls=args.drop_lst_style):
            slide_new[n] = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
            reports[n] = log
        elif args.check and log:
            reports[n] = log

    # 版式：给"单行放不下"的内容页标题写溢出缩排，避免第二行穿过标题下分隔线
    # ⚠ 2026-09-18 弃用：本实现只认 type="title"（漏掉标题槽位）、把可用宽度写死成
    #   cx=700.6pt（漏掉左右内边距 14.4pt）、把版式字号写死成 28pt（漏掉 44pt 的
    #   Section Title / 26.46pt 的目录槽位），算出偏大的缩放比 —— 缩了仍然换行，
    #   而且看起来"已经处理过"，更难排查。统一改走 fix_title_oneline.py（范围与
    #   度量都从模板推导，且有 min-scale 下限与 UNRESOLVED 上报）。
    extra_new = {}
    n_shrink = 0
    if args.overflow_shrink:
        print('!! --overflow-shrink 已弃用（缩了仍可能换行）。')
        print('   请改用：python fix_title_oneline.py "<deck.pptx>" --out "<new.pptx>"')
        print('   复核：  python verify_title_lines.py "<new.pptx>"')
        z.close()
        return 2

    # 版式 / 母版 / 页面 的标题色归一
    # ⚠ 默认**只改页面**：版式与母版是模板本体，改它们等于改模板（2026-09-18 事故）。
    n_col = 0
    if args.unify_title_color:
        new = args.unify_title_color
        old_color = args.from_color
        skipped_tpl = []
        for n in sorted(names):
            is_slide = bool(re.match(r'ppt/slides/slide\d+\.xml$', n))
            is_lay = bool(re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n))
            is_mas = bool(re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n))
            if not (is_slide or is_lay or is_mas):
                continue
            if (is_lay or is_mas) and not args.allow_template_write:
                skipped_tpl.append(n)
                continue
            # 该部件若已被前序步骤改过，就在**改过的版本**上继续改，避免互相覆盖
            if n in slide_new:
                root = etree.fromstring(slide_new[n])
            elif n in extra_new:
                root = etree.fromstring(extra_new[n])
            else:
                root = etree.fromstring(z.read(n))
            targets = []
            for sp in root.findall('.//p:cSld/p:spTree/p:sp', NS):
                ph = sp.find('.//p:nvPr/p:ph', NS)
                if ph is not None and (ph.get('type') or '') in ('title', 'ctrTitle'):
                    targets.append(sp)
            for ts in root.findall('.//p:txStyles/p:titleStyle', NS):
                targets.append(ts)
            n_hit = 0
            for t in targets:
                n_hit += unify_color(t, old_color, new)
            if not n_hit:
                continue
            n_col += n_hit
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            if n in slide_new:
                slide_new[n] = data
            else:
                extra_new[n] = data
            print('  色归一 %-40s %d 处 %s -> %s' % (posixpath.basename(n), n_hit, old_color, new))
        if skipped_tpl:
            print('  色归一：跳过 %d 个模板部件（版式/母版）—— 模板本体只读。'
                  % len(skipped_tpl))
            print('           %s' % ', '.join(posixpath.basename(x) for x in skipped_tpl[:10]))
            print('           确实要改模板请加 --allow-template-write（改完产物不再等于模板格式）')

    print('=== 清覆盖报告（scope=%s）===' % args.scope)
    for n in sorted(reports, key=lambda s: int(re.search(r'(\d+)', posixpath.basename(s)).group(1))):
        pg = re.search(r'(\d+)', posixpath.basename(n)).group(1)
        print('  P%-3s %s' % (pg, ' | '.join(reports[n])[:150]))
    print('受影响页数: %d' % len(slide_new))
    if args.unify_title_color:
        print('色归一合计: %d 处 -> %s' % (n_col, args.unify_title_color))

    if not args.out_deck:
        z.close()
        print('\n（--check 模式，未输出文件）')
        return 0

    if os.path.exists(args.out_deck):
        os.remove(args.out_deck)
    zo = zipfile.ZipFile(args.out_deck, 'w', zipfile.ZIP_DEFLATED)
    zo.writestr(zipfile.ZipInfo(CT_PART), z.read(CT_PART))     # CT 必须第一个条目
    for item in z.infolist():
        n = item.filename
        if n == CT_PART or n.endswith('/'):
            continue
        if n in slide_new:
            zo.writestr(item, slide_new[n])
        elif n in extra_new:
            zo.writestr(item, extra_new[n])
        else:
            zo.writestr(item, z.read(n))
    zo.close()
    z.close()
    print('\n产出 %s (%.1f MB)' % (args.out_deck, os.path.getsize(args.out_deck) / 1048576))
    return 0


if __name__ == '__main__':
    sys.exit(main())
