import docx
import json

doc_word = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')
print('Word paragraphs count:', len(doc_word.paragraphs))

# Let's inspect the sections in ANEXO 004
sections = []
current_sec = "PRELIMINARES"
sec_p_counts = {}

for i, p in enumerate(doc_word.paragraphs):
    txt = p.text.strip()
    if p.style.name.startswith('Heading 1') or 'CAPÍTULO' in txt.upper() or 'INTRODUCCIÓN' in txt.upper() or 'RESUMEN' in txt.upper() or 'CONCLUSIONES' in txt.upper() or 'RECOMENDACIONES' in txt.upper() or 'BIBLIOGRAFÍA' in txt.upper() or 'ANEXOS' in txt.upper():
        current_sec = txt[:50]
    sec_p_counts[current_sec] = sec_p_counts.get(current_sec, 0) + 1

print("\nSections detected in Word doc:")
for s, count in sec_p_counts.items():
    print(f"  - {s} ({count} paragraphs)")
