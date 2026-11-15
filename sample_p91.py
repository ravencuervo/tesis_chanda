import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

for i in range(91, 160):
    txt = doc.paragraphs[i].text.strip()
    if txt:
        print(f"P{i:03d}: {txt[:100]}")
