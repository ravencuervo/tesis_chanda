import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt:
        continue
    # print headers or lines starting with chapter/numbers
    if any(txt.startswith(prefix) for prefix in ['CAPÍTULO', 'INTRODUCCIÓN', 'RESUMEN', 'ABSTRACT', '1.', '2.', '3.', '4.', 'CONCLUSIONES', 'RECOMENDACIONES', 'BIBLIOGRAFÍA', 'ANEXO']):
        print(f"P{i:03d} [{p.style.name}]: {txt[:90]}")
