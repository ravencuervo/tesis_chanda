import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

for idx in [270, 271, 272, 273, 274, 275, 282, 283, 284, 285, 286, 287, 288]:
    if idx < len(doc.paragraphs):
        print(f"P{idx:03d}: {doc.paragraphs[idx].text[:120]}")
