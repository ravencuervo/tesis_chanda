import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

print("=== P160 to P350 sample ===")
for i in range(160, 360, 10):
    txt = doc.paragraphs[i].text.strip()
    if txt:
        print(f"P{i:03d}: {txt[:100]}")

print("\n=== P360 to P640 sample ===")
for i in range(360, len(doc.paragraphs), 20):
    txt = doc.paragraphs[i].text.strip()
    if txt:
        print(f"P{i:03d}: {txt[:100]}")
