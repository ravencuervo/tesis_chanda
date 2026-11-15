import docx
import json

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

def dump_table(t, name):
    print(f"\n==================== {name} ({len(t.rows)}x{len(t.columns)}) ====================")
    for r in t.rows:
        row_str = " | ".join([c.text.strip().replace('\n', ' ') for c in r.cells])
        print(row_str)

dump_table(doc.tables[2], "Tabla 3: Distribución Muestral por Servicios Eléctricos")
dump_table(doc.tables[3], "Tabla 4: Operacionalización de Variables")
dump_table(doc.tables[9], "Tabla 10: Prueba de Normalidad General")
dump_table(doc.tables[10], "Tabla 11: Frecuencias Comunicación Interna")
dump_table(doc.tables[11], "Tabla 12: Frecuencias Desempeño Laboral")
