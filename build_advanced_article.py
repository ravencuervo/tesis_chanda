# build_advanced_article.py
# -*- coding: utf-8 -*-
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import json
import xml.sax.saxutils as saxutils
import os

# Load CSL database
with open('article_csl_db.json', 'r', encoding='utf-8') as f:
    csl_db = json.load(f)

doc = docx.Document()

# Configure page margins: 2.5 cm on all sides
for s in doc.sections:
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(2.5)
    s.bottom_margin = Cm(2.5)
    s.left_margin = Cm(2.5)
    s.right_margin = Cm(2.5)

def set_run_font(run, font_name="Times New Roman", size_pt=12, bold=False, italic=False, color_rgb=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

def add_csl_field(p, keys, display_text=None):
    if isinstance(keys, str):
        keys = [keys]
    items = []
    fmts = []
    for k in keys:
        if k in csl_db:
            items.append(csl_db[k]['csl_json']['citationItems'][0])
            fmts.append(csl_db[k]['formatted_cite'].strip('()'))
        else:
            print("Missing citation key:", k)
    if display_text is None:
        display_text = f"({'; '.join(fmts)})"
        
    csl_payload = {
        "citationItems": items,
        "mendeley": {
            "formattedCitation": display_text,
            "plainTextFormattedCitation": display_text,
            "previouslyFormattedCitation": display_text
        },
        "properties": {"noteIndex": 0},
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
    }
    csl_str = json.dumps(csl_payload, ensure_ascii=False)
    escaped_csl_str = saxutils.escape(csl_str)
    escaped_display_text = saxutils.escape(display_text)
    
    r_b = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin" w:fldLock="1"/></w:r>')
    r_i = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:instrText xml:space="preserve">ADDIN CSL_CITATION {escaped_csl_str}</w:instrText></w:r>')
    r_s = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>')
    r_t = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:noProof/></w:rPr><w:t>{escaped_display_text}</w:t></w:r>')
    r_e = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>')
    
    p._p.append(r_b)
    p._p.append(r_i)
    p._p.append(r_s)
    p._p.append(r_t)
    p._p.append(r_e)

def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, line_spacing=1.5, bold=False, italic=False, size=12):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        r = p.add_run(text)
        set_run_font(r, size_pt=size, bold=bold, italic=italic)
    return p

def add_h1(text):
    return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=16, space_after=6, bold=True, size=13)

def add_h2(text):
    return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=12, space_after=4, bold=True, italic=True, size=12)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def apply_apa_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)
    if len(table.rows) > 0:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

def add_custom_table(headers, rows_data, col_widths=None):
    t = doc.add_table(rows=len(rows_data)+1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    apply_apa_table_borders(t)
    
    # Headers
    for c_i, h in enumerate(headers):
        cell = t.cell(0, c_i)
        cell.paragraphs[0].text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].paragraph_format.line_spacing = 1.0
        cell.paragraphs[0].paragraph_format.space_after = Pt(2)
        cell.paragraphs[0].paragraph_format.space_before = Pt(2)
        set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=True)
        set_cell_margins(cell)
        if col_widths and c_i < len(col_widths):
            cell.width = Inches(col_widths[c_i])
            
    # Data rows
    for r_i, r_data in enumerate(rows_data):
        for c_i, val in enumerate(r_data):
            cell = t.cell(r_i+1, c_i)
            cell.paragraphs[0].text = str(val)
            cell.paragraphs[0].paragraph_format.line_spacing = 1.0
            cell.paragraphs[0].paragraph_format.space_after = Pt(2)
            cell.paragraphs[0].paragraph_format.space_before = Pt(2)
            # Align center for numbers, left for text
            if any(char.isdigit() for char in str(val)) or val in ["--", "Total", "TOTAL"]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            bold_flag = True if (r_i == len(rows_data)-1 and ("TOTAL" in str(r_data[0]).upper() or "TOTAL" in str(val).upper())) or "**" in str(val) else False
            set_run_font(cell.paragraphs[0].runs[0], size_pt=9.5, bold=bold_flag)
            set_cell_margins(cell)
            if col_widths and c_i < len(col_widths):
                cell.width = Inches(col_widths[c_i])
    return t

def add_figure_image(image_path, fig_num, title, note_text, width_inches=5.6):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(4)
    r_img = p_img.add_run()
    r_img.add_picture(image_path, width=Inches(width_inches))
    
    p_cap = add_p(f"Figura {fig_num}", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=1, bold=True, size=10)
    p_title = add_p(title, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=3, italic=True, size=10)
    p_note = add_p(f"Nota. {note_text}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=10, italic=True, size=9)
    return p_note

print("Core helper functions defined.")
