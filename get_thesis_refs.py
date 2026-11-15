import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

refs = []
for i in range(589, 644):
    txt = doc.paragraphs[i].text.strip()
    if txt:
        refs.append((i, txt))

print(f"Total references found in thesis: {len(refs)}")
for idx, (p_idx, r) in enumerate(refs[:30]):
    print(f"{idx+1}. [P{p_idx}]: {r[:100]}...")
