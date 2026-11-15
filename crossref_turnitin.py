import docx
import json
import re

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

with open('turnitin_extracted_matches.json', 'r', encoding='utf-8') as f:
    pdf_matches = json.load(f)

# Normalize text for matching
def normalize(txt):
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', '', txt.lower())).strip()

# Build index of doc paragraphs
doc_paras = []
for i, p in enumerate(doc.paragraphs):
    clean = normalize(p.text)
    if len(clean) > 20:
        doc_paras.append((i, p.text.strip(), clean))

# For each page in pdf_matches, find matching doc paragraphs
matched_doc_paras = {}
for p_num, data in pdf_matches.items():
    pdf_text = data['text']
    pdf_clean = normalize(pdf_text)
    words = pdf_clean.split()
    if len(words) < 5:
        continue
    
    # check 6-word shingles
    found_in_paras = set()
    for w_i in range(0, len(words) - 5, 4):
        shingle = " ".join(words[w_i:w_i+6])
        for p_idx, orig_text, p_clean in doc_paras:
            if shingle in p_clean:
                found_in_paras.add(p_idx)
                
    if found_in_paras:
        matched_doc_paras[p_num] = list(sorted(found_in_paras))

print(f"Total PDF pages matched to Word paragraphs: {len(matched_doc_paras)}")
# Flatten all matched paragraph indices
all_flagged_paras = sorted(list(set(idx for paras in matched_doc_paras.values() for idx in paras)))
print(f"Total unique Word paragraphs flagged by Turnitin: {len(all_flagged_paras)}")
print("Paragraph index ranges:", [(min(g), max(g)) for g in [
    [all_flagged_paras[0]]
]])

# Group consecutive paragraphs
ranges = []
start = all_flagged_paras[0]
prev = start
for x in all_flagged_paras[1:]:
    if x <= prev + 2:
        prev = x
    else:
        ranges.append((start, prev))
        start = x
        prev = x
ranges.append((start, prev))

print("Flagged paragraph clusters in Word:")
for r_start, r_end in ranges:
    sample = doc.paragraphs[r_start].text[:70].strip()
    print(f"  P{r_start:03d} to P{r_end:03d} ({r_end - r_start + 1} paras): {sample}...")
