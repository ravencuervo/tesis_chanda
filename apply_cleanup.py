import docx
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import paraphrase_cleanup_all

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

def update_paragraph_text(p, new_text):
    font_name = "Times New Roman"
    font_size = 12
    if p.runs and p.runs[0].font.name:
        font_name = p.runs[0].font.name
    if p.runs and p.runs[0].font.size:
        font_size = p.runs[0].font.size.pt

    p.text = ""
    run = p.add_run(new_text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

applied = 0
for idx, new_text in paraphrase_cleanup_all.cleanup.items():
    if idx < len(doc.paragraphs):
        update_paragraph_text(doc.paragraphs[idx], new_text)
        applied += 1

print(f"Applied {applied} cleanup replacements!")
doc.save('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')
doc.save('ANEXO 004 - BORRADOR DE TESIS EPG UNA - TURNITIN_CORREGIDO_MENOR_15.docx')
print("Successfully saved updated documents!")
