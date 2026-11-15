import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

print(f"Total tables: {len(doc.tables)}")

for idx, t in enumerate(doc.tables):
    rows = len(t.rows)
    cols = len(t.columns) if rows > 0 else 0
    # get sample text from first 2 rows
    sample = []
    for r in t.rows[:2]:
        sample.append(" | ".join([c.text.strip().replace('\n', ' ') for c in r.cells]))
    # find preceding paragraph (table title)
    title = "UNKNOWN"
    # search backwards from table's position if possible or match table index
    print(f"Table {idx+1} ({rows}x{cols}): {' /// '.join(sample[:2])}")
