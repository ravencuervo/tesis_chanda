import docx

doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

def dump_table(t, name):
    print(f"\n==================== {name} ({len(t.rows)}x{len(t.columns)}) ====================")
    for r in t.rows:
        row_str = " | ".join([c.text.strip().replace('\n', ' ') for c in r.cells])
        print(row_str)

dump_table(doc.tables[15], "Tabla 16: Frecuencias Canales y Medios")
dump_table(doc.tables[16], "Tabla 17: Frecuencias Percepción Clima Laboral")
dump_table(doc.tables[20], "Tabla 21: Frecuencias Calidad Flujo Comunicacional")
dump_table(doc.tables[21], "Tabla 22: Frecuencias Satisfacción y Cohesión Laboral")
dump_table(doc.tables[25], "Tabla 26: Frecuencias Clima Laboral")
dump_table(doc.tables[26], "Tabla 27: Frecuencias Desempeño Laboral (OE3)")
