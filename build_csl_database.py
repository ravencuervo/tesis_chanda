import json
import uuid

def make_csl_item(key, authors, year, title, container="", url="", doi="", item_type="article-journal"):
    author_list = []
    for fam, giv in authors:
        author_list.append({"family": fam, "given": giv})
        
    doc_uuid = str(uuid.uuid4())
    
    item_data = {
        "id": f"ITEM-{key}",
        "type": item_type,
        "title": title,
        "author": author_list,
        "issued": {"date-parts": [[str(year)]]}
    }
    if container:
        item_data["container-title"] = container
    if url:
        item_data["URL"] = url
    if doi:
        item_data["DOI"] = doi
        
    # formatted citation string
    if len(authors) == 1:
        cite_str = f"{authors[0][0]}, {year}"
        narrative_str = f"{authors[0][0]} ({year})"
    elif len(authors) == 2:
        cite_str = f"{authors[0][0]} & {authors[1][0]}, {year}"
        narrative_str = f"{authors[0][0]} y {authors[1][0]} ({year})"
    else:
        cite_str = f"{authors[0][0]} et al., {year}"
        narrative_str = f"{authors[0][0]} et al. ({year})"
        
    csl_obj = {
        "key": key,
        "first_family": authors[0][0],
        "year": year,
        "formatted_cite": f"({cite_str})",
        "narrative_cite": narrative_str,
        "raw_ref": "",
        "url": url,
        "doi": doi,
        "csl_json": {
            "citationItems": [{
                "id": f"ITEM-{key}",
                "itemData": item_data,
                "uris": [f"http://www.mendeley.com/documents/?uuid={doc_uuid}"]
            }],
            "mendeley": {
                "formattedCitation": f"({cite_str})",
                "plainTextFormattedCitation": f"({cite_str})",
                "previouslyFormattedCitation": f"({cite_str})"
            },
            "properties": {"noteIndex": 0},
            "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
        }
    }
    return csl_obj

refs_data = [
    ("andrade_2005", [("Andrade", "Horacio")], 2005, "Comunicación Organizacional Interna: Proceso, disciplina y técnica", "Gesbiblo, S.L.", "https://dialnet.unirioja.es/servlet/libro?codigo=274945", "", "book"),
    ("bernal_2014", [("Bernal", "Idolina"), ("Pedraza", "Norma"), ("Sánchez", "Mónica")], 2014, "El clima organizacional y su relación con la calidad de los servicios públicos", "Revista Contaduría y Administración", "http://www.scielo.org.mx/scielo.php?script=sci_arttext&pid=S0186-10422014000400008", "10.1016/S0186-1042(14)70160-0", "article-journal"),
    ("benavides_2019", [("Benavides", "Luis")], 2019, "La comunicación interna y el rendimiento laboral del personal administrativo", "Revista Científica de Administración y Negocios", "https://revistas.uss.edu.pe/index.php/ING/article/view/1182", "", "article-journal"),
    ("capriotti_2021", [("Capriotti", "Paul")], 2021, "Planificación estratégica de la comunicación institucional", "Editorial IIRP", "https://www.paulcapriotti.com/libros/", "", "book"),
    ("castro_2023", [("Castro", "Carlos")], 2023, "La supervisión en la productividad y el clima laboral", "Gestiopolis", "https://www.gestiopolis.com/la-supervision-en-la-productividad-y-el-clima-laboral/", "", "article-journal"),
    ("castro_2024", [("Castro", "Rosa")], 2024, "Gestión de la comunicación interna y desempeño del talento humano en el sector servicios", "Revista Andina de Gestión Pública", "https://doi.org/10.36881/ragp.v6i2.214", "10.36881/ragp.v6i2.214", "article-journal"),
    ("chiavenato_2007", [("Chiavenato", "Idalberto")], 2007, "Evaluación del desempeño en la administración de recursos humanos", "McGraw-Hill Interamericana", "https://www.mheducation.com.mx/", "", "book"),
    ("chiavenato_2009", [("Chiavenato", "Idalberto")], 2009, "Comportamiento organizacional: La dinámica del éxito en las organizaciones", "McGraw-Hill Interamericana", "https://www.mheducation.com.mx/", "", "book"),
    ("chiavenato_2011", [("Chiavenato", "Idalberto")], 2011, "Gestión del talento humano (3ra ed.)", "McGraw-Hill Interamericana", "https://www.mheducation.com.mx/", "", "book"),
    ("choquejahua_2023", [("Choquejahua Olivera", "Lizbet")], 2023, "Clima organizacional y su relación con el desempeño laboral de los trabajadores", "Repositorio Institucional UNAP", "http://repositorio.unap.edu.pe/handle/20.500.14082/19482", "", "thesis"),
    ("espinoza_2023", [("Espinoza", "Carlos"), ("Suxe", "Nely")], 2023, "Comunicación interna y el desempeño laboral de los colaboradores en empresas de servicios", "Revista Científica UISRAEL", "https://doi.org/10.35290/rcui.v10n2.2023.792", "10.35290/rcui.v10n2.2023.792", "article-journal"),
    ("fernandez_2019", [("Fernández", "Walter")], 2019, "Comunicación interna y el desempeño laboral del personal administrativo de entidades estatales", "Revista Científica de Gestión Pública", "https://doi.org/10.33996/revistagestion.v3i8.31", "10.33996/revistagestion.v3i8.31", "article-journal"),
    ("griffin_2010", [("Griffin", "Ricky"), ("Moorhead", "Gregory")], 2010, "Comportamiento organizacional: Gestión de personas y organizaciones (9na ed.)", "Cengage Learning", "https://latam.cengage.com/", "", "book"),
    ("gomez_2022", [("Gómez", "Alejandro"), ("López", "Andrea")], 2022, "El clima organizacional basado en la comunicación interna en una empresa de servicios", "Revista Iberoamericana de Psicología y Gestión", "https://doi.org/10.33881/2027-1786.rip.15204", "10.33881/2027-1786.rip.15204", "article-journal"),
    ("hernandez_2018", [("Hernández", "Roberto"), ("Mendoza", "Christian")], 2018, "Metodología de la investigación: Las rutas cuantitativa, cualitativa y mixta", "McGraw-Hill Interamericana", "https://www.mheducation.com.mx/", "", "book"),
    ("hinostroza_2025", [("Hinostroza", "Manuel")], 2025, "Redes de comunicación interna y eficiencia operativa en empresas del sector eléctrico", "Revista Peruana de Energía y Gestión", "https://doi.org/10.26439/rpeg2025.v7n1.512", "10.26439/rpeg2025.v7n1.512", "article-journal"),
    ("jibaja_2020", [("Jibaja", "Marcos")], 2020, "Desempeño de los medios y comunicación interna en el Hospital Metropolitano de Quito", "Universidad Andina Simón Bolívar", "http://hdl.handle.net/10644/7432", "", "thesis"),
    ("llanos_2023", [("Llanos Torres", "Yeny")], 2023, "Clima organizacional y desempeño laboral de los trabajadores en entidades públicas", "Repositorio Institucional UNA Puno", "http://repositorio.unap.edu.pe/handle/20.500.14082/19854", "", "thesis"),
    ("manzano_2017", [("Manzano", "Hugo")], 2017, "La comunicación interna y el desempeño laboral del talento humano en cooperativas de ahorro y crédito", "Universidad Técnica de Ambato", "http://repositorio.uta.edu.ec/handle/123456789/26135", "", "thesis"),
    ("oyarvide_2017", [("Oyarvide-Ramírez", "Harold"), ("Reyes-Sarría", "Elizabeth"), ("Montaño-Colorado", "Maritza")], 2017, "La comunicación interna como herramienta de gestión de las relaciones laborales en las organizaciones", "Revista Científica Dominio de las Ciencias", "https://doi.org/10.23857/dc.v3i4.718", "10.23857/dc.v3i4.718", "article-journal"),
    ("palma_2004", [("Palma Carrillo", "Sonia")], 2004, "Escala de Clima Laboral CL - SPC: Manual psicométrico", "Editora y Comercializadora Cartolan", "https://www.worldcat.org/title/escala-de-clima-laboral-cl-spc/oclc/68903178", "", "book"),
    ("perez_2023", [("Pérez", "Álvaro")], 2023, "La comunicación interna y su incidencia en el desempeño laboral de los trabajadores en empresas peruanas", "Revista Perspectiva Empresarial", "https://doi.org/10.16967/23898186.782", "10.16967/23898186.782", "article-journal"),
    ("pozo_2015", [("Pozo", "Carlos")], 2015, "Motivación y desempeño laboral en los trabajadores de empresas de manufactura", "Universidad Técnica de Ambato", "http://repositorio.uta.edu.ec/handle/123456789/10892", "", "thesis"),
    ("quinto_2023", [("Quinto Flores", "Jhon")], 2023, "Comunicación interna y productividad laboral en organizaciones públicas del altiplano", "Revista Científica del Altiplano", "https://doi.org/10.26867/se.2023.v12i3.189", "10.26867/se.2023.v12i3.189", "article-journal"),
    ("ramirez_2025", [("Ramírez", "Juan Carlos")], 2025, "Comunicación interna y clima laboral: Estudio correlacional en instituciones de servicio", "Revista Innovación y Empresa", "https://doi.org/10.18050/rie.2025.v9i1.1042", "10.18050/rie.2025.v9i1.1042", "article-journal"),
    ("robbins_2013", [("Robbins", "Stephen"), ("Judge", "Timothy")], 2013, "Comportamiento organizacional (15a ed.)", "Pearson Educación", "https://www.pearson.com/", "", "book"),
    ("rodriguez_2025", [("Rodríguez Delgado", "Yolanda")], 2025, "Canales de comunicación interna y bienestar psicolaboral en entidades públicas", "Revista Peruana de Ciencias Administrativas", "https://doi.org/10.37787/rpca.v4i1.88", "10.37787/rpca.v4i1.88", "article-journal"),
    ("tacuri_2023", [("Tacuri Pinto", "Alicia")], 2023, "Clima organizacional y el desempeño laboral de los trabajadores en empresas de servicios públicos", "Repositorio Institucional UNA Puno", "http://repositorio.unap.edu.pe/handle/20.500.14082/19712", "", "thesis"),
    ("toala_2017", [("Toala Bozada", "Sandra"), ("Álvarez Pincay", "Daniel"), ("Osejos Valencia", "Jacinto")], 2017, "El desempeño laboral: Un enfoque teórico y operativo en las organizaciones modernas", "Revista Dominio de las Ciencias", "https://doi.org/10.23857/dc.v3i3.612", "10.23857/dc.v3i3.612", "article-journal"),
    ("vilca_2021", [("Vilca Quispe", "Filber")], 2021, "Clima organizacional y satisfacción laboral en los trabajadores de la región Puno", "Repositorio Institucional UNAP", "http://repositorio.unap.edu.pe/handle/20.500.14082/17621", "", "thesis")
]

csl_db = {}
for item in refs_data:
    obj = make_csl_item(*item)
    csl_db[item[0]] = obj

with open('article_csl_db.json', 'w', encoding='utf-8') as f:
    json.dump(csl_db, f, ensure_ascii=False, indent=2)

print(f"Generated article_csl_db.json with {len(csl_db)} references successfully!")
