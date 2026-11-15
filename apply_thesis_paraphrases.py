import docx
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import paraphrase_batch1
import paraphrase_batch2
import paraphrase_batch3
import paraphrase_batch4
import paraphrase_batch5

# Consolidate all replacements
all_replacements = {}
all_replacements.update(paraphrase_batch1.replacements)
all_replacements.update(paraphrase_batch2.batch2)
all_replacements.update(paraphrase_batch3.batch3)
all_replacements.update(paraphrase_batch4.batch4)
all_replacements.update(paraphrase_batch5.batch5)

print(f"Total paragraphs scheduled for replacement: {len(all_replacements)}")

# Load document
doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

def update_paragraph_text(p, new_text):
    # If paragraph has runs, keep style and font of first run if possible
    font_name = "Times New Roman"
    font_size = 12
    if p.runs and p.runs[0].font.name:
        font_name = p.runs[0].font.name
    if p.runs and p.runs[0].font.size:
        font_size = p.runs[0].font.size.pt

    # Clear runs
    p.text = ""
    run = p.add_run(new_text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    
    # Ensure w:rFonts is explicitly set for Word rendering
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

applied_count = 0
for idx, new_text in all_replacements.items():
    if idx < len(doc.paragraphs):
        update_paragraph_text(doc.paragraphs[idx], new_text)
        applied_count += 1
    else:
        print(f"Warning: Index {idx} out of range (max {len(doc.paragraphs)})")

print(f"Successfully applied {applied_count} paraphrased paragraphs to the document!")

output_filename = 'ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx'
doc.save(output_filename)
print(f"Saved updated thesis file as: {output_filename}")

# Also save a distinct copy so the user can easily distinguish if desired
doc.save('ANEXO 004 - BORRADOR DE TESIS EPG UNA - TURNITIN_CORREGIDO_MENOR_15.docx')
print("Saved clean copy as: ANEXO 004 - BORRADOR DE TESIS EPG UNA - TURNITIN_CORREGIDO_MENOR_15.docx")
