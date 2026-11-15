import docx
import json
import re

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

with open('turnitin_extracted_matches.json', 'r', encoding='utf-8') as f:
    pdf_matches = json.load(f)

def normalize(txt):
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', '', txt.lower())).strip()

doc_paras = [(i, normalize(p.text)) for i, p in enumerate(doc.paragraphs) if len(normalize(p.text)) > 20]

matched_paras = set()
for p_num, data in pdf_matches.items():
    pdf_clean = normalize(data['text'])
    words = pdf_clean.split()
    for w_i in range(0, len(words) - 5, 4):
        shingle = " ".join(words[w_i:w_i+6])
        for p_idx, p_clean in doc_paras:
            if shingle in p_clean:
                matched_paras.add(p_idx)

body_matched = [p for p in sorted(list(matched_paras)) if p > 130 and p < 590]
print(f"Total body paragraphs matching: {len(body_matched)}")
for p in body_matched:
    print(f"P{p:03d} [{doc.paragraphs[p].style.name}]: {doc.paragraphs[p].text[:90]}")
