import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

for i in range(90, len(doc.paragraphs)):
    txt = doc.paragraphs[i].text.strip()
    if not txt:
        continue
    # Check if heading or contains keywords
    for kw in ['RESUMEN', 'ABSTRACT', 'INTRODUCCI', 'CAPÍTULO', 'CONCLUSI', 'RECOMENDA', 'BIBLIOGRA', 'ANEXO']:
        if kw in txt.upper() and len(txt) < 80:
            print(f"P{i:03d} [{doc.paragraphs[i].style.name}]: {txt}")
            break
