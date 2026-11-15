import docx
import json
import re

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

with open('turnitin_extracted_matches.json', 'r', encoding='utf-8') as f:
    pdf_matches = json.load(f)

def normalize(txt):
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', '', txt.lower())).strip()

doc_paras = []
for i, p in enumerate(doc.paragraphs):
    clean = normalize(p.text)
    if len(clean) > 20:
        doc_paras.append((i, clean))

matched_doc_paras = {}
total_shingles_checked = 0
total_shingles_hit = 0

for p_num, data in pdf_matches.items():
    pdf_text = data['text']
    pdf_clean = normalize(pdf_text)
    words = pdf_clean.split()
    if len(words) < 5:
        continue
    
    found_in_paras = set()
    for w_i in range(0, len(words) - 5, 4):
        total_shingles_checked += 1
        shingle = " ".join(words[w_i:w_i+6])
        for p_idx, p_clean in doc_paras:
            if shingle in p_clean:
                found_in_paras.add(p_idx)
                total_shingles_hit += 1
                break
                
    if found_in_paras:
        matched_doc_paras[p_num] = list(sorted(found_in_paras))

all_flagged_paras = sorted(list(set(idx for paras in matched_doc_paras.values() for idx in paras)))
print(f"Total PDF pages with shingle matches in updated document: {len(matched_doc_paras)} (down from 69)")
print(f"Total unique Word paragraphs matching Turnitin shingles: {len(all_flagged_paras)} (down from 181)")
print(f"Shingles matching: {total_shingles_hit} / {total_shingles_checked} ({total_shingles_hit/max(1, total_shingles_checked)*100:.2f}%)")

print("\nRemaining matching paragraphs sample:")
for p_idx in all_flagged_paras[:15]:
    print(f"  P{p_idx:03d}: {doc.paragraphs[p_idx].text[:80]}...")
