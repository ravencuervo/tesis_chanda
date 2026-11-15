import json

with open('paragraphs_to_paraphrase.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

for g in ['RESUMEN_INTRO', 'PLANTEAMIENTO_PROBLEMA']:
    print(f"==================== {g} ====================")
    for item in data[g]:
        print(f"P{item['index']:03d} [{item['style']}]:\n{item['text']}\n")
