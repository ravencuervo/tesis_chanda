import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

print(f"Total tables: {len(doc.tables)}")

for t_idx, table in enumerate(doc.tables):
    rows_text = []
    for row in table.rows[:3]:
        cells = [c.text.strip().replace('\n', ' ') for c in row.cells]
        rows_text.append(" | ".join(cells))
    header_summary = " /// ".join(rows_text[:2])
    if any(k in header_summary.lower() for k in ['spearman', 'correlac', 'rho', 'desempeño', 'clima', 'comunicación']):
        print(f"\n--- Table {t_idx} ---")
        for r in table.rows:
            print(" | ".join([c.text.strip().replace('\n', ' ') for c in r.cells]))
