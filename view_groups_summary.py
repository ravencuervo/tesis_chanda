import json

with open('paragraphs_to_paraphrase.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

for g in ['MARCO_TEORICO', 'METODOLOGIA', 'RESULTADOS_NORMALIDAD_DECISION', 'DISCUSION_CONTRASTACION', 'CONCLUSIONES_RECOMENDACIONES']:
    print(f"=== {g} (Count: {len(data[g])}) ===")
    for item in data[g][:5]:
        print(f"P{item['index']:03d}: {item['text'][:120]}...")
