#!/usr/bin/env python3
# Copyright 2026 jsksjosk199-lang
# SPDX-License-Identifier: Apache-2.0
"""Inspect or repair section breaks in the supported three-table DOCX layout.

This helper only handles section boundaries. It does not restyle report cells,
invent content, or insert breaks within sections with more than four events.
"""
import argparse
import json
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit('Missing python-docx; use the bundled Codex Python runtime or install the dependency.')

LABELS = {
    '本月完成工作': ('本月完成工作', '本月工作完成情况', '本月工作完成'),
    '下月工作计划': ('下月工作计划', '下月计划'),
    '问题复盘与总结': ('问题复盘与总结', '问题复盘总结'),
}

def normalize(text):
    return re.sub(r'[\s\u3000]+', '', text)

def heading(table):
    return ''.join(table._element.tr_lst[0].xpath('.//w:t/text()')) if table._element.tr_lst else ''

def page_breaks(node):
    return [b for b in node.xpath('.//w:br') if b.get(qn('w:type')) == 'page']

def section_tables(doc):
    matches = {name: [] for name in LABELS}
    for i, table in enumerate(doc.tables):
        text = normalize(heading(table))
        for name, variants in LABELS.items():
            if any(v in text for v in variants):
                matches[name].append((i, table, text))
    wrong = {key: len(vals) for key, vals in matches.items() if len(vals) != 1}
    if wrong:
        raise ValueError('未能唯一识别三张栏目表；不自动修改。匹配数：' + json.dumps(wrong,ensure_ascii=False))
    result=[matches[k][0] for k in LABELS]
    if [x[0] for x in result] != sorted(x[0] for x in result):
        raise ValueError('栏目表的顺序与预期不同，不自动修改。')
    return result

def boundary_nodes(a, b):
    nodes=[]
    n=a.getnext()
    while n is not None and n is not b:
        if n.tag != qn('w:p'):
            raise ValueError('栏目之间包含非段落对象；需要人工检查，未修改文件。')
        # Only remove/reformat paragraphs proven empty apart from page breaks,
        # paragraph properties, run properties and retained bookmarks. A w:sym,
        # drawing, tab, field or other non-text element can carry real content.
        for child in n:
            if child.tag == qn('w:pPr'):
                if child.find(qn('w:sectPr')) is not None:
                    raise ValueError('栏目之间包含分节；需要人工检查，未修改文件。')
            elif child.tag in (qn('w:bookmarkStart'),qn('w:bookmarkEnd')):
                continue
            elif child.tag == qn('w:r'):
                for item in child:
                    if item.tag == qn('w:rPr'): continue
                    if item.tag == qn('w:br') and item.get(qn('w:type')) == 'page': continue
                    raise ValueError('栏目之间含有文字、符号或其他内容；需要人工检查，未修改文件。')
            else:
                raise ValueError('栏目之间含有非空对象；需要人工检查，未修改文件。')
        nodes.append(n)
        n=n.getnext()
    if n is None:
        raise ValueError('栏目不在同一文档正文中，未修改文件。')
    return nodes

def inspect_document(doc):
    matched=section_tables(doc)
    sections=[]
    for name, (idx, table, title) in zip(LABELS,matched):
        sections.append({'name':name, 'table_index':idx, 'heading':title,
                         'rows':len(table.rows),
                         'floating':table._element.tblPr.find(qn('w:tblpPr')) is not None,
                         'page_breaks_in_table':len(page_breaks(table._element))})
    embedded=sum(len(page_breaks(t._element.tr_lst[0])) for _,t,_ in matched[1:])
    boundaries=[]
    for (_,a,_),(_,b,_) in zip(matched,matched[1:]):
        try:
            parts=boundary_nodes(a._element,b._element)
            boundaries.append({'supported':True,'paragraphs':len(parts),
                               'manual_page_breaks':sum(len(page_breaks(p)) for p in parts)})
        except ValueError as exc:
            boundaries.append({'supported':False,'reason':str(exc)})
    return {'sections':sections, 'embedded_header_page_breaks':embedded, 'boundaries':boundaries}

def delete_break(br, remove_empty_run=True):
    parent=br.getparent()
    parent.remove(br)
    # Avoid leaving the empty run that previously held the page break.
    if remove_empty_run and parent.tag == qn('w:r') and all(child.tag==qn('w:rPr') for child in parent):
        parent.getparent().remove(parent)

def heading_break_in_effect(paragraph):
    """Resolve direct and inherited paragraph-style page-break-before."""
    value=paragraph.paragraph_format.page_break_before
    if value is not None: return value
    style=paragraph.style
    seen=set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        value=style.paragraph_format.page_break_before
        if value is not None: return value
        style=style.base_style
    defaults=paragraph.part.document.styles.element.xpath(
        './w:docDefaults/w:pPrDefault/w:pPr/w:pageBreakBefore')
    return bool(defaults and defaults[0].get(qn('w:val')) not in ('0','false','off'))

def make_break(paragraph, wps_center):
    p=paragraph
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    f=p.paragraph_format
    f.page_break_before=False
    f.keep_with_next=False
    f.space_before=Pt(0); f.space_after=Pt(0)
    f.first_line_indent=Pt(0); f.left_indent=Pt(0)
    # Empirically calibrated for the local WPS editor. Visual QA is still needed.
    f.right_indent=Cm(5 if wps_center else 0)
    f.line_spacing_rule=WD_LINE_SPACING.EXACTLY; f.line_spacing=Pt(20)
    run=p.add_run()
    run.font.name='仿宋'
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'仿宋')
    run.font.size=Pt(13)
    run.add_break(WD_BREAK.PAGE)

def repair_document(doc, boundaries, wps_center, inline_floating, bottom_margin_cm):
    matched=section_tables(doc)
    # Preflight every boundary before mutating the in-memory document.
    pairs=[]
    for i,((_,a,_),(_,b,_)) in enumerate(zip(matched,matched[1:])):
        if boundaries!='both' and i != (0 if boundaries=='month-plan' else 1):
            continue
        nodes=boundary_nodes(a._element,b._element)
        if any(n.xpath('.//w:bookmarkStart | .//w:bookmarkEnd') for n in nodes[:-1]):
            raise ValueError('多段空段落中存在书签；需要人工检查，未修改文件。')
        pairs.append((a,b,nodes))
    moves=0
    cleared_heading_break_before=0
    inlined=[]
    for a,b,nodes in pairs:
        # Only move breaks in the upcoming heading row; never remove breaks in data rows.
        for br in page_breaks(b._element.tr_lst[0]):
            delete_break(br, remove_empty_run=False); moves+=1
        for p_node in b._element.tr_lst[0].xpath('.//w:p'):
            paragraph=Paragraph(p_node,doc._body)
            if heading_break_in_effect(paragraph):
                paragraph.paragraph_format.page_break_before=False
                cleared_heading_break_before+=1
        for p in nodes:
            for br in page_breaks(p): delete_break(br)
        if nodes:
            chosen=nodes[-1]
            for n in nodes[:-1]: n.getparent().remove(n)
        else:
            chosen=OxmlElement('w:p')
            b._element.addprevious(chosen)
        make_break(Paragraph(chosen,doc._body),wps_center)
        if inline_floating:
            pr=b._element.tblPr
            floating=pr.find(qn('w:tblpPr'))
            if floating is not None:
                pr.remove(floating)
                overlap=pr.find(qn('w:tblOverlap'))
                if overlap is not None: pr.remove(overlap)
                inlined.append(normalize(heading(b)))
    old_margins=[]
    if bottom_margin_cm is not None:
        if not 0.5 <= bottom_margin_cm <= 5:
            raise ValueError('底边距应在 0.5 至 5 厘米之间。')
        for section in doc.sections:
            old_margins.append(round(section.bottom_margin.cm,2))
            section.bottom_margin=Cm(bottom_margin_cm)
    return {'boundaries':boundaries,'embedded_breaks_moved':moves,
            'heading_page_break_before_cleared':cleared_heading_break_before,
            'floating_tables_changed_to_inline':inlined,
            'old_bottom_margins_cm':old_margins,'new_bottom_margin_cm':bottom_margin_cm,
            'wps_center_requested':wps_center}

def main():
    p=argparse.ArgumentParser(description=__doc__)
    sub=p.add_subparsers(dest='command',required=True)
    a=sub.add_parser('inspect'); a.add_argument('source',type=Path)
    b=sub.add_parser('repair'); b.add_argument('source',type=Path); b.add_argument('--out',type=Path,required=True)
    b.add_argument('--boundaries',choices=('month-plan','plan-review','both'),required=True,
                   help='Select exactly which section boundaries the user requested.')
    b.add_argument('--wps-center',action='store_true',help='Apply the locally tested WPS marker alignment; verify visually.')
    b.add_argument('--inline-floating',action='store_true',help='Convert floating following tables to inline to avoid WPS marker overlaps.')
    b.add_argument('--bottom-margin-cm',type=float,help='Change bottom margin only when needed; report the deviation.')
    args=p.parse_args()
    try:
        if args.command=='repair':
            if args.source.resolve()==args.out.resolve():
                raise ValueError('输出路径不能与源文件相同。')
            if args.out.exists():
                raise ValueError('输出文件已存在；请使用新路径。')
        doc=Document(args.source)
        if args.command=='inspect': result=inspect_document(doc)
        else:
            result=repair_document(doc,args.boundaries,args.wps_center,args.inline_floating,args.bottom_margin_cm)
            args.out.parent.mkdir(parents=True,exist_ok=True)
            doc.save(args.out)
            result['output']=str(args.out)
        print(json.dumps(result,ensure_ascii=False,indent=2))
    except Exception as exc:
        print(f'{type(exc).__name__}: {exc}',file=sys.stderr)
        return 2
    return 0

if __name__=='__main__':
    raise SystemExit(main())
