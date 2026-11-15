import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

print("=== KEY METHODOLOGICAL AND STATISTICAL DATA ===")
for i in [139, 140, 141, 142, 289, 290, 291, 292, 293, 294, 295, 296, 297, 298, 299, 300, 301, 302, 303, 307, 308, 309, 310, 311, 312, 313, 314, 315, 316]:
    if i < len(doc.paragraphs):
        txt = doc.paragraphs[i].text.strip()
        if txt:
            print(f"P{i}: {txt}")
