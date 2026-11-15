import docx
import os
import pymupdf

def search_text(filename, query):
    doc = docx.Document(filename)
    print(f"=== Searching '{query}' in {filename} ===")
    found = 0
    for i, p in enumerate(doc.paragraphs):
        if query.lower() in p.text.lower():
            print(f"P{i}: {p.text.strip()[:140]}")
            found += 1
            if found >= 10:
                break
    # also search in tables
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                if query.lower() in cell.text.lower():
                    print(f"Table {t_idx} [R{r_idx}C{c_idx}]: {cell.text.strip()[:140]}")
                    found += 1
                    if found >= 15:
                        break
    if found == 0:
        print("None found")

for fn in ["Informe_final_cuantitativo.docx", "ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx"]:
    for q in ["artículo", "articulo", "directri", "anexo 004", "anexo 4", "mendeley", "formato"]:
        search_text(fn, q)
