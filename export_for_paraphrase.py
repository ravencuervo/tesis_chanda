import docx
import json

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

# Let's inspect paragraphs in groups and export them to a json file for precise paraphrasing
flagged_groups = [
    ("RESUMEN_INTRO", list(range(139, 143)) + list(range(149, 157)) + [161]),
    ("MARCO_TEORICO", list(range(182, 189)) + [199] + list(range(205, 208)) + list(range(214, 216)) + [219] + list(range(226, 228)) + [230, 246]),
    ("ANTECEDENTES", list(range(253, 276))),
    ("PLANTEAMIENTO_PROBLEMA", list(range(276, 304))),
    ("METODOLOGIA", list(range(304, 306)) + [308] + list(range(317, 322)) + [333, 337] + list(range(341, 343)) + list(range(346, 350))),
    ("RESULTADOS_NORMALIDAD_DECISION", [356, 357, 362, 363, 364, 369, 370, 371, 372, 384, 388, 389, 390, 395, 396, 397, 402, 403, 404, 405, 416, 417, 418, 430, 431, 432, 439, 440, 441, 442, 453, 454, 459, 460, 461, 467, 468, 469, 474, 475, 476, 477, 488, 489, 490, 494, 495, 496]),
    ("DISCUSION_CONTRASTACION", list(range(543, 552)) + list(range(555, 574))),
    ("CONCLUSIONES_RECOMENDACIONES", list(range(575, 580)) + list(range(582, 587)))
]

data_to_paraphrase = {}
for g_name, indices in flagged_groups:
    group_list = []
    for idx in indices:
        if idx < len(doc.paragraphs):
            t = doc.paragraphs[idx].text.strip()
            if t:
                group_list.append({"index": idx, "style": doc.paragraphs[idx].style.name, "text": t})
    data_to_paraphrase[g_name] = group_list

with open('paragraphs_to_paraphrase.json', 'w', encoding='utf-8') as f:
    json.dump(data_to_paraphrase, f, ensure_ascii=False, indent=2)

print("Exported flagged paragraphs:")
for g_name, items in data_to_paraphrase.items():
    print(f"  {g_name}: {len(items)} paragraphs")
