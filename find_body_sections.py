import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

print("=== BODY SECTIONS ===")
for i in range(90, len(doc.paragraphs)):
    txt = doc.paragraphs[i].text.strip()
    if any(txt.upper().startswith(k) for k in [
        'RESUMEN', 'ABSTRACT', 'INTRODUCCIÓN', 'CAPÍTULO I', 'CAPÍTULO II', 'CAPÍTULO III', 
        'CAPÍTULO IV', '4.1', '4.2', 'CONCLUSIONES', 'RECOMENDACIONES', 'BIBLIOGRAFÍA', 'ANEXOS'
    ]):
        print(f"P{i:03d}: {txt[:80]}")
