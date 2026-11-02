# -*- coding: utf-8 -*-
"""
通用影视全流程出版级 Word 排版引擎 (pub_layout_engine.py)
- 严格执行中西文混排规范（中文宋体/黑体，西文 Times New Roman）
- 表头深海蓝底白字加粗并跨页重复（tblHeader），数据行不设 cantSplit（彻底杜绝跨页截断）
- 支持纵向标准页面与横向宽幅（8列导演分镜专用）页面
"""
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def _set_cell_margins(cell, top=110, bottom=110, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def _set_table_borders(table, color="B8C4D4", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def init_pub_doc(landscape=False):
    doc = docx.Document()
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width = Inches(11.69)
        sec.page_height = Inches(8.27)
        sec.top_margin = Inches(0.65)
        sec.bottom_margin = Inches(0.65)
        sec.left_margin = Inches(0.65)
        sec.right_margin = Inches(0.65)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width = Inches(8.5)
        sec.page_height = Inches(11.0)
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.65)
        sec.right_margin = Inches(0.65)
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    return doc

def add_pub_p(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size_pt=10.5,
              font_cn="宋体", space_before_pt=2, space_after_pt=4, color_rgb=None,
              first_line_indent_pt=0, line_spacing=1.35):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.line_spacing = line_spacing
    if first_line_indent_pt:
        pf.first_line_indent = Pt(first_line_indent_pt)
    r = p.add_run(str(text))
    r.bold = bold
    r.font.name = 'Times New Roman'
    r.font.size = Pt(size_pt)
    r._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
    if color_rgb:
        r.font.color.rgb = RGBColor(*color_rgb)
    return p

def add_pub_table(doc, headers, rows, col_widths_in=None, font_size_pt=9.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _set_table_borders(t)

    hdr_row = t.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:tblHeader'))
    trPr.append(OxmlElement('w:cantSplit'))

    for i, h_text in enumerate(headers):
        c = hdr_row.cells[i]
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_margins(c, top=120, bottom=120, left=140, right=140)
        c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1F3864"/>'))
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run(str(h_text))
        r.bold = True
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size_pt)
        r.font.color.rgb = RGBColor(255, 255, 255)
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

    for r_idx, row_data in enumerate(rows):
        row = t.add_row()
        # 注意：数据行严禁注入 w:cantSplit，确保超长分集大纲或跨页内容在Word中100%完整跨页显示
        bg_hex = "F4F7FA" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>'))
            val_str = str(val).strip()
            lines = [ln.strip() for ln in val_str.split("\n") if ln.strip()]
            if not lines:
                lines = ["-"]
            for l_i, ln in enumerate(lines):
                p = cell.paragraphs[0] if l_i == 0 else cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(1.5)
                p.paragraph_format.space_after = Pt(2.0)
                p.paragraph_format.line_spacing = 1.32
                is_b = (c_idx == 0 and len(headers) <= 4)
                r = p.add_run(ln)
                r.bold = is_b
                r.font.name = 'Times New Roman'
                r.font.size = Pt(font_size_pt)
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体' if is_b else '宋体')

    if col_widths_in:
        for row in t.rows:
            for i, w in enumerate(col_widths_in):
                if i < len(row.cells):
                    row.cells[i].width = Inches(w)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)
    return t
