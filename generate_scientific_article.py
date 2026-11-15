# generate_scientific_article.py
# -*- coding: utf-8 -*-
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import json
import xml.sax.saxutils as saxutils

# Load CSL database
with open('article_csl_db.json', 'r', encoding='utf-8') as f:
    csl_db = json.load(f)

doc = docx.Document()

# Configure page margins: 2.5 cm on all sides
for s in doc.sections:
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(2.5)
    s.bottom_margin = Cm(2.5)
    s.left_margin = Cm(2.5)
    s.right_margin = Cm(2.5)

def set_run_font(run, font_name="Times New Roman", size_pt=12, bold=False, italic=False, color_rgb=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

def add_csl_field(p, keys, display_text=None):
    if isinstance(keys, str):
        keys = [keys]
    items = []
    fmts = []
    for k in keys:
        if k in csl_db:
            items.append(csl_db[k]['csl_json']['citationItems'][0])
            fmts.append(csl_db[k]['formatted_cite'].strip('()'))
        else:
            print("Missing citation key:", k)
    if display_text is None:
        display_text = f"({'; '.join(fmts)})"
        
    csl_payload = {
        "citationItems": items,
        "mendeley": {
            "formattedCitation": display_text,
            "plainTextFormattedCitation": display_text,
            "previouslyFormattedCitation": display_text
        },
        "properties": {"noteIndex": 0},
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
    }
    csl_str = json.dumps(csl_payload, ensure_ascii=False)
    escaped_csl_str = saxutils.escape(csl_str)
    escaped_display_text = saxutils.escape(display_text)
    
    r_b = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin" w:fldLock="1"/></w:r>')
    r_i = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:instrText xml:space="preserve">ADDIN CSL_CITATION {escaped_csl_str}</w:instrText></w:r>')
    r_s = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>')
    r_t = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:noProof/></w:rPr><w:t>{escaped_display_text}</w:t></w:r>')
    r_e = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>')
    
    p._p.append(r_b)
    p._p.append(r_i)
    p._p.append(r_s)
    p._p.append(r_t)
    p._p.append(r_e)

def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, line_spacing=1.5, bold=False, italic=False, size=12):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        r = p.add_run(text)
        set_run_font(r, size_pt=size, bold=bold, italic=italic)
    return p

def add_h1(text):
    return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=14, space_after=4, bold=True, size=13)

def add_h2(text):
    return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=3, bold=True, italic=True, size=12)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def apply_apa_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)
    
    # Border below header row
    if len(table.rows) > 0:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

# ==============================================================================
# HEADER / TITLE SECTION
# ==============================================================================
p_title_es = add_p("COMUNICACIÓN INTERNA, CLIMA ORGANIZACIONAL Y DESEMPEÑO LABORAL EN PROFESIONALES Y TÉCNICOS DEL SECTOR ELÉCTRICO EN PUNO, PERÚ", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=6, bold=True, size=14)
p_title_en = add_p("INTERNAL COMMUNICATION, ORGANIZATIONAL CLIMATE, AND JOB PERFORMANCE AMONG PROFESSIONAL AND TECHNICAL STAFF IN THE ELECTRICAL SECTOR IN PUNO, PERU", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=14, bold=True, italic=True, size=12)

p_author = add_p("Chanda Ketty Sosa Quispe", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, bold=True, size=11)
p_affil = add_p("Maestría en Contabilidad y Administración, Escuela de Posgrado, Universidad Nacional del Altiplano, Puno, Perú\nORCID: https://orcid.org/0009-0004-9218-4720 | chanda.sosa@unap.edu.pe", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14, italic=True, size=10)

# ==============================================================================
# RESUMEN Y PALABRAS CLAVE
# ==============================================================================
add_h1("RESUMEN")
p_res = add_p(
    "El propósito fundamental de la investigación consistió en determinar la relación entre la comunicación interna, "
    "el clima organizacional y el desempeño laboral exhibido por el personal técnico y profesional de la Empresa Regional "
    "de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.) durante el periodo 2022-2024. La metodología "
    "respondió a un enfoque cuantitativo, de nivel descriptivo-correlacional y diseño no experimental transversal. Se trabajó "
    "con una muestra censal integrada por 158 trabajadores clasificados en los grupos ocupacionales profesionales y técnicos. "
    "Para el levantamiento empírico se aplicó una batería psicométrica estructurada en escala Likert validada mediante juicio "
    "de expertos, demostrando una consistencia interna sobresaliente (Alfa de Cronbach de 0,876 para Comunicación Interna; "
    "0,927 para Clima Laboral; y 0,846 para Desempeño Laboral). Las pruebas de Kolmogórov-Smirnov y Shapiro-Wilk confirmaron "
    "un comportamiento no paramétrico (p < 0,05), por lo que se empleó el estadístico Rho de Spearman para la contrastación "
    "de hipótesis. Los hallazgos revelaron una correlación positiva considerable y estadísticamente significativa entre la "
    "comunicación interna y el desempeño laboral (rs = 0,743; p = 0,000). Asimismo, se constataron asociaciones directas muy "
    "fuertes entre los canales comunicacionales y el clima laboral (rs = 0,840; p = 0,000), una correlación moderada entre la "
    "calidad del flujo informativo y la satisfacción laboral (rs = 0,567; p = 0,000), y una correlación positiva muy fuerte "
    "entre el clima organizacional y el desempeño (rs = 0,806; p = 0,000). Se concluye que un sistema comunicacional interno "
    "eficaz y multidireccional consolida un entorno de trabajo armónico y potencia la productividad en el servicio eléctrico."
)

p_kw_es = add_p("Palabras clave: ", bold=True, space_after=12)
r_kw = p_kw_es.add_run("Clima organizacional, comunicación interna, desempeño laboral, eficiencia técnica, sector eléctrico.")
set_run_font(r_kw, italic=True)

# ==============================================================================
# ABSTRACT Y KEYWORDS
# ==============================================================================
add_h1("ABSTRACT")
p_abs = add_p(
    "The central aim of this investigation was to determine the relationship between internal communication dynamics, "
    "organizational climate, and job performance among professional and technical staff at the Empresa Regional de Servicios "
    "Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.) during the 2022–2024 period. Methodologically, the study "
    "followed a quantitative paradigm with a descriptive-correlational scope and a non-experimental, cross-sectional design. "
    "A census sample comprising 158 professional and technical employees across all regional electrical divisions was evaluated. "
    "Data collection employed a five-point Likert psychometric questionnaire validated by expert panel judgment, exhibiting "
    "exceptional internal reliability (Cronbach's alpha of 0.876 for Internal Communication, 0.927 for Organizational Climate, "
    "and 0.846 for Job Performance). Normality testing through Kolmogorov-Smirnov and Shapiro-Wilk statistics rejected Gaussian "
    "distribution (p < 0.05), warranting non-parametric Spearman's Rho correlation. Results demonstrated a substantial and "
    "statistically significant positive correlation between internal communication and job performance (rs = 0.743; p = 0.000). "
    "Furthermore, strong associations emerged between institutional communication media and workplace climate perception "
    "(rs = 0.840; p = 0.000), a moderate correlation between informational flow quality and employee satisfaction (rs = 0.567; "
    "p = 0.000), and a very strong correlation between organizational climate and operational productivity (rs = 0.806; p = 0.000). "
    "In conclusion, strengthening multidirectional internal communication channels optimizes the socio-labor ecosystem and drives "
    "operational excellence in public electrical distribution services."
)

p_kw_en = add_p("Keywords: ", bold=True, space_after=14)
r_kw_en = p_kw_en.add_run("Electrical sector, internal communication, job performance, organizational climate, technical efficiency.")
set_run_font(r_kw_en, italic=True)

# ==============================================================================
# INTRODUCCIÓN (Con 25-30 citas Mendeley integradas)
# ==============================================================================
add_h1("INTRODUCCIÓN")

p_intro1 = add_p(
    "En la administración corporativa del siglo XXI, el capital humano se posiciona como el factor diferenciador indispensable "
    "para alcanzar altos estándares de competitividad, calidad y sostenibilidad institucional "
)
add_csl_field(p_intro1, "chiavenato_2011", "(Chiavenato, 2011)")
p_intro1.add_run(
    ". En las entidades prestadoras de servicios públicos esenciales, donde la continuidad operativa y la seguridad técnica "
    "resultan determinantes para el bienestar colectivo, las interacciones comunicativas internas constituyen el sistema "
    "circulatorio mediante el cual se transmiten las directrices, los valores compartidos y las metas estratégicas corporativas "
)
add_csl_field(p_intro1, "andrade_2005", "(Andrade, 2005)")
p_intro1.add_run(". No obstante, diversas investigaciones empíricas demuestran que las fallas en la difusión de información originan distorsiones que debilitan el clima laboral y menoscaban el rendimiento productivo de los colaboradores.")

p_intro2 = add_p(
    "Desde el punto de vista teórico, la comunicación interna se cimenta en la teoría de las relaciones humanas y en el enfoque "
    "sistémico de las organizaciones "
)
add_csl_field(p_intro2, "capriotti_2021", "(Capriotti, 2021)")
p_intro2.add_run(
    ". Dicho constructo no se limita a la emisión vertical de instrucciones, sino que abarca un proceso dialógico multidireccional "
    "—descendente, ascendente y diagonal— orientado a cohesionar las distintas áreas de la empresa "
)
add_csl_field(p_intro2, "oyarvide_2017", "(Oyarvide-Ramírez et al., 2017)")
p_intro2.add_run(
    ". Cuando las instituciones logran articular canales formales accesibles y confiables, los colaboradores experimentan mayor certidumbre funcional y claridad en sus roles laborales "
)
add_csl_field(p_intro2, "fernandez_2019", "(Fernández, 2019)")
p_intro2.add_run(". Por el contrario, la opacidad informativa favorece la propagación de rumores informales y tensiones interpersonales.")

p_intro3 = add_p(
    "En estrecha correspondencia, el clima organizacional alude a la configuración de percepciones globales que los colaboradores "
    "construyen en torno a las políticas, prácticas y procedimientos de su entorno cotidiano "
)
add_csl_field(p_intro3, "palma_2004", "(Palma Carrillo, 2004)")
p_intro3.add_run(
    ". Dimensiones cardinales como la autorrealización personal, el involucramiento laboral, la supervisión de soporte, la comunicación y las condiciones físicas de trabajo modulan la experiencia subjetiva del colaborador "
)
add_csl_field(p_intro3, "bernal_2014", "(Bernal et al., 2014)")
p_intro3.add_run(
    ". Un ambiente institucional favorable fomenta la reciprocidad y la confianza mutua, mientras que un microclima hostil propicia desmotivación, ausentismo y deterioro en la calidad de la atención al usuario "
)
add_csl_field(p_intro3, "gomez_2022", "(Gómez & López, 2022)")
p_intro3.add_run(".")

p_intro4 = add_p(
    "Por su parte, el desempeño laboral representa el valor agregado que las conductas y resultados de los trabajadores aportan al logro de las metas institucionales "
)
add_csl_field(p_intro4, "toala_2017", "(Toala Bozada et al., 2017)")
p_intro4.add_run(
    ". Este constructo multidimensional comprende tanto el rendimiento técnico vinculado a la tarea (eficiencia y eficacia en la utilización de recursos) como el desempeño contextual manifestado en conductas de colaboración, proactividad y disciplina "
)
add_csl_field(p_intro4, "robbins_2013", "(Robbins & Judge, 2013)")
p_intro4.add_run(
    ". La evaluación del desempeño debe concebirse como un diagnóstico pedagógico y estimulador del talento, evitando enfoques meramente punitivos "
)
add_csl_field(p_intro4, ["chiavenato_2007", "pozo_2015"], "(Chiavenato, 2007; Pozo, 2015)")
p_intro4.add_run(".")

p_intro5 = add_p(
    "En la arena internacional, las investigaciones corroboran el nexo indisoluble entre estos factores. En Ecuador, "
)
add_csl_field(p_intro5, "jibaja_2020", "Jibaja (2020)")
p_intro5.add_run(
    " comprobó que la oportunidad en los medios de comunicación interna se asocia directamente con la eficiencia operativa en organizaciones de servicios. En el ámbito cooperativo, "
)
add_csl_field(p_intro5, "manzano_2017", "Manzano (2017)")
p_intro5.add_run(
    " evidenció que la retroalimentación fluida entre jefaturas y personal operativo minimiza errores transaccionales y eleva el compromiso laboral."
)

p_intro6 = add_p(
    "En el contexto peruano, diversos estudios destacan la urgencia de optimizar la comunicación corporativa. En Lima, "
)
add_csl_field(p_intro6, "rodriguez_2025", "Rodríguez Delgado (2025)")
p_intro6.add_run(
    " halló correlaciones directas significativas entre la claridad de los canales organizacionales y la satisfacción del personal estatal. De modo convergente, "
)
add_csl_field(p_intro6, "ramirez_2025", "Ramírez (2025)")
p_intro6.add_run(
    " confirmó que la comunicación interna actúa como predictor medular de la estabilidad del clima laboral. En el sector energético nacional, "
)
add_csl_field(p_intro6, "hinostroza_2025", "Hinostroza (2025)")
p_intro6.add_run(
    " reportó que la disponibilidad de directivas técnicas oportunas reduce sustancialmente los tiempos de respuesta ante contingencias en redes eléctricas."
)

p_intro7 = add_p(
    "En la región andina y el sur del país, "
)
add_csl_field(p_intro7, "castro_2024", "Castro (2024)")
p_intro7.add_run(
    " constató que el diálogo asertivo potencia el trabajo interdisciplinario en empresas de servicios. Específicamente en el altiplano de Puno, las investigaciones de "
)
add_csl_field(p_intro7, "berolatti_2023", "Berolatti Obando (2023)")
p_intro7.add_run(", ")
add_csl_field(p_intro7, "llanos_2023", "Llanos Torres (2023)")
p_intro7.add_run(" y ")
add_csl_field(p_intro7, "tacuri_2023", "Tacuri Pinto (2023)")
p_intro7.add_run(
    " convergen en que la percepción del clima organizacional se encuentra estrechamente vinculada al reconocimiento profesional y al soporte comunicacional de las jefaturas. Asimismo, "
)
add_csl_field(p_intro7, "choquejahua_2023", "Choquejahua Olivera (2023)")
p_intro7.add_run(" y ")
add_csl_field(p_intro7, "vilca_2021", "Vilca Quispe (2021)")
p_intro7.add_run(
    " ratificaron que el confort psicosocial predice favorablemente la eficiencia del talento humano en instituciones puneñas."
)

p_intro8 = add_p(
    "En la Empresa Regional de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.), entidad encargada de "
    "la distribución y comercialización de energía en el departamento de Puno, la dispersión geográfica de los centros operativos "
    "ha generado asimetrías en los canales comunicativos entre la sede central y los servicios eléctricos zonales. Dicha desconexión "
    "provoca retrasos en la atención de directivas técnicas, percepción de inequidad y desmotivación en los operarios. Frente a esta "
    "problemática empírica, la presente investigación formuló como objetivo general determinar la influencia de la comunicación interna "
    "en el clima laboral y en el desempeño de los profesionales y técnicos de Electro Puno S.A.A. durante el periodo 2022-2024, "
    "contrastando rigurosamente las hipótesis de relación entre sus componentes estructurales."
)

# ==============================================================================
# MÉTODOS
# ==============================================================================
add_h1("MÉTODOS")

add_h2("Ámbito o lugar de estudio")
add_p(
    "La investigación se desarrolló en la Empresa Regional de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.), "
    "abarcando tanto su sede central corporativa ubicada en la ciudad de Puno como sus centros de operaciones y servicios eléctricos "
    "distribuidos estratégicamente en las provincias de Puno, San Román (Juliaca), Melgar (Ayaviri), Azángaro, Huancané y El Collao (Ilave), "
    "en el departamento de Puno, Perú."
)

add_h2("Descripción de métodos")
add_p(
    "a) Periodo de estudio y diseño de investigación:\n"
    "El levantamiento de información de campo y análisis cuantitativo comprendió el bienio 2022–2024. El estudio adoptó una tipología "
    "aplicada con enfoque cuantitativo, alcance descriptivo-correlacional y un diseño no experimental de corte transversal, "
    "observando las variables en su contexto institucional natural sin manipulación deliberada "
)
add_csl_field(doc.paragraphs[-1], "hernandez_2018", "(Hernández & Mendoza, 2018)")
doc.paragraphs[-1].add_run(".")

add_p(
    "b) Población y delimitación muestral:\n"
    "La población estuvo constituida por la totalidad del personal profesional y técnico con vínculo laboral vigente en Electro Puno S.A.A. "
    "registrado en la nómina corporativa. Dado el tamaño abordable y la dispersión geográfica controlada, se optó por un muestreo censal "
    "no probabilístico, evaluando a 158 colaboradores (100% del marco poblacional), conformados por 63 profesionales universitarios (39,87%) "
    "y 95 técnicos operativos y de mantenimiento de redes (60,13%)."
)

add_p(
    "c) Materiales e instrumentos de recolección de datos:\n"
    "Se diseñó una cédula psicométrica estructurada compuesta por tres escalas en formato Likert graduadas en cinco opciones ordinales "
    "(1: Totalmente en desacuerdo, 2: En desacuerdo, 3: Indiferente, 4: De acuerdo, 5: Totalmente de acuerdo). La escala de Comunicación "
    "Interna constó de 13 ítems; la escala de Clima Organizacional comprendió 16 reactivos adaptados del modelo CL-SPC "
)
add_csl_field(doc.paragraphs[-1], "palma_2004", "(Palma Carrillo, 2004)")
doc.paragraphs[-1].add_run(
    "; y la escala de Desempeño Laboral incluyó 15 reactivos sustentados en las dimensiones de eficiencia, eficacia y conducta laboral "
)
add_csl_field(doc.paragraphs[-1], "chiavenato_2009", "(Chiavenato, 2009)")
doc.paragraphs[-1].add_run(
    ". La validez de contenido fue dictaminada favorablemente mediante juicio pericial emitido por tres expertos académicos con grado de doctor. "
    "La confiabilidad por consistencia interna se evaluó mediante el coeficiente Alfa de Cronbach, obteniéndose valores excelentes: "
    "Comunicación Interna (α = 0,876), Clima Laboral (α = 0,927) y Desempeño Laboral (α = 0,846), así como en sus subdimensiones "
    "(Canales comunicacionales α = 0,936; Percepción del clima α = 0,786)."
)

add_p(
    "d) Variables analizadas:\n"
    "Variable X (Independiente): Comunicación Interna, analizada a través de las dimensiones de Canales y medios comunicativos, y Calidad del flujo informativo (claridad, oportunidad y retroalimentación).\n"
    "Variable Y (Dependiente): Desempeño Laboral, abordada mediante Eficiencia operativa, Eficacia en el cumplimiento de tareas y Conducta laboral.\n"
    "Variable Z (Mediadora / Interviniente): Clima Organizacional, examinada a través del Vínculo organizacional, Involucramiento, Supervisión y Satisfacción laboral."
)

add_p(
    "e) Tratamiento y prueba estadística aplicada:\n"
    "El procesamiento estadístico se llevó a cabo en el software IBM SPSS Statistics versión 26. Previo al análisis inferencial, "
    "se sometieron los datos a las pruebas de bondad de ajuste de Kolmogórov-Smirnov y Shapiro-Wilk (n = 158), arrojando niveles "
    "de significancia bilateral asintótica p = 0,000 (p < 0,05) en todas las variables, evidenciando un comportamiento no gaussiano. "
    "Por consiguiente, se aplicó la prueba no paramétrica del coeficiente de correlación por rangos de Spearman (rs) con un nivel de "
    "confianza del 99% (significancia bilateral α = 0,01)."
)

# ==============================================================================
# RESULTADOS Y DISCUSIÓN
# ==============================================================================
add_h1("RESULTADOS Y DISCUSIÓN")

# --- OBJETIVO GENERAL ---
add_h2("Análisis y contrastación del Objetivo General: Comunicación interna y desempeño laboral")
add_p(
    "Conforme a los cálculos inferenciales, se contrastó la hipótesis general de investigación sobre la muestra censal de 158 trabajadores. "
    "En la Tabla 1 se presentan los coeficientes de correlación no paramétrica obtenidos entre la comunicación interna y el desempeño laboral."
)

# TABLA 1
add_p("Tabla 1", bold=True, space_after=1)
add_p("Correlación de Spearman entre Comunicación Interna y Desempeño Laboral", italic=True, space_after=4)

t1 = doc.add_table(rows=4, cols=4)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
apply_apa_table_borders(t1)

headers_t1 = ["Estadístico", "Variable", "Comunicación Interna", "Desempeño Laboral"]
for c_i, h in enumerate(headers_t1):
    cell = t1.cell(0, c_i)
    cell.paragraphs[0].text = h
    set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=True)
    set_cell_margins(cell)

data_t1 = [
    ["Rho de Spearman", "Comunicación Interna", "1,000", ",743**"],
    ["", "Sig. (bilateral)", "--", ",000"],
    ["", "N", "158", "158"]
]
for r_i, row in enumerate(data_t1):
    for c_i, val in enumerate(row):
        cell = t1.cell(r_i+1, c_i)
        cell.paragraphs[0].text = val
        bold_flag = True if "**" in val else False
        set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=bold_flag)
        set_cell_margins(cell)

add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral). Fuente: Matriz de datos procesada en SPSS v.26 a partir de encuestas a profesionales y técnicos de Electro Puno S.A.A.", italic=True, size=10, space_after=8)

p_disc1 = add_p(
    "Los resultados expuestos en la Tabla 1 constatan la existencia de una correlación positiva considerable y estadísticamente "
    "significativa entre la comunicación interna y el desempeño laboral (rs = 0,743; p = 0,000). Al ser el p-valor inferior al "
    "nivel de significancia establecido (p < 0,01), se rechaza la hipótesis nula y se corrobora que una adecuada gestión de los "
    "flujos comunicacionales optimiza el rendimiento del capital humano. Este hallazgo concuerda plenamente con las pesquisas de "
)
add_csl_field(p_disc1, "quinto_2023", "Quinto Flores (2023)")
p_disc1.add_run(
    ", quien registró coeficientes robustos entre la comunicación corporativa y la productividad laboral en el sector público del altiplano. De igual modo, respalda los postulados teóricos de "
)
add_csl_field(p_disc1, "andrade_2005", "Andrade (2005)")
p_disc1.add_run(" y ")
add_csl_field(p_disc1, "fernandez_2019", "Fernández (2019)")
p_disc1.add_run(
    ", quienes sostienen que la comunicación interna opera como una herramienta estratégica que previene la ambigüedad en el rol y maximiza la entrega funcional en puestos técnicos de alta responsabilidad."
)

# --- OBJETIVO ESPECÍFICO 1 ---
add_h2("Objetivo Específico 1: Canales y medios de comunicación interna y percepción del clima laboral")
add_p(
    "Para dilucidar la correspondencia entre la provisión de canales comunicacionales y la atmósfera socioemocional de trabajo, "
    "se procesaron los puntajes bivariados mostrados a continuación en la Tabla 2."
)

# TABLA 2
add_p("Tabla 2", bold=True, space_after=1)
add_p("Correlación de Spearman entre Canales y Medios de Comunicación Interna y Percepción del Clima Laboral", italic=True, space_after=4)

t2 = doc.add_table(rows=4, cols=4)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
apply_apa_table_borders(t2)

headers_t2 = ["Estadístico", "Variable / Dimensión", "Canales y Medios", "Percepción del Clima"]
for c_i, h in enumerate(headers_t2):
    cell = t2.cell(0, c_i)
    cell.paragraphs[0].text = h
    set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=True)
    set_cell_margins(cell)

data_t2 = [
    ["Rho de Spearman", "Canales y Medios Comunicativos", "1,000", ",840**"],
    ["", "Sig. (bilateral)", "--", ",000"],
    ["", "N", "158", "158"]
]
for r_i, row in enumerate(data_t2):
    for c_i, val in enumerate(row):
        cell = t2.cell(r_i+1, c_i)
        cell.paragraphs[0].text = val
        bold_flag = True if "**" in val else False
        set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=bold_flag)
        set_cell_margins(cell)

add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral). Fuente: Base de datos institucional de campo en Electro Puno S.A.A.", italic=True, size=10, space_after=8)

p_disc2 = add_p(
    "A partir de los cálculos plasmados en la Tabla 2, se aprecia una correlación positiva muy fuerte y altamente significativa "
    "(rs = 0,840; p = 0,000) entre la disponibilidad de canales comunicacionales y la percepción del clima laboral. Dicho coeficiente "
    "demuestra que cuando los canales informativos (circulares técnicas, correo corporativo, reuniones de coordinación) operan con "
    "celeridad y accesibilidad, el sentido de respaldo institucional se incrementa notablemente. Este resultado es congruente con "
    "lo determinado por "
)
add_csl_field(p_disc2, "espinoza_2023", "Espinoza y Suxe (2023)")
p_disc2.add_run(
    ", quienes comprobaron que los canales descendentes y horizontales fortalecen los vínculos de confianza grupal, y con las observaciones de "
)
add_csl_field(p_disc2, "hinostroza_2025", "Hinostroza (2025)")
p_disc2.add_run(
    " en el sector de energía eléctrica de la capital, donde la modernización de los canales colaborativos redujo las discrepancias interdepartamentales."
)

# --- OBJETIVO ESPECÍFICO 2 ---
add_h2("Objetivo Específico 2: Calidad del flujo comunicacional y nivel de cohesión y satisfacción laboral")
add_p(
    "En lo concerniente a los atributos intrínsecos de la información (precisión, inmediatez y retroalimentación bidireccional), "
    "se correlacionaron frente a los índices de cohesión grupal y bienestar ocupacional, según se aprecia en la Tabla 3."
)

# TABLA 3
add_p("Tabla 3", bold=True, space_after=1)
add_p("Correlación de Spearman entre Calidad del Flujo Comunicacional y Satisfacción y Cohesión Laboral", italic=True, space_after=4)

t3 = doc.add_table(rows=4, cols=4)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER
apply_apa_table_borders(t3)

headers_t3 = ["Estadístico", "Dimensión", "Calidad del Flujo", "Satisfacción y Cohesión"]
for c_i, h in enumerate(headers_t3):
    cell = t3.cell(0, c_i)
    cell.paragraphs[0].text = h
    set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=True)
    set_cell_margins(cell)

data_t3 = [
    ["Rho de Spearman", "Calidad del Flujo Comunicacional", "1,000", ",567**"],
    ["", "Sig. (bilateral)", "--", ",000"],
    ["", "N", "158", "158"]
]
for r_i, row in enumerate(data_t3):
    for c_i, val in enumerate(row):
        cell = t3.cell(r_i+1, c_i)
        cell.paragraphs[0].text = val
        bold_flag = True if "**" in val else False
        set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=bold_flag)
        set_cell_margins(cell)

add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral). Fuente: Registro de encuestas estructuradas a técnicos y profesionales.", italic=True, size=10, space_after=8)

p_disc3 = add_p(
    "La evidencia resumida en la Tabla 3 constata una correlación positiva moderada-considerable (rs = 0,567; p = 0,000) entre la calidad "
    "del flujo de mensajes y la satisfacción con cohesión de los operarios. Se infiere que la claridad en las instrucciones y la "
    "presencia de retroalimentación oportuna desde los niveles jerárquicos superiores atenúan la incertidumbre laboral y estimulan "
    "el trabajo solidario en terreno. Estos resultados validan las conclusiones de "
)
add_csl_field(p_disc3, "rodriguez_2025", "Rodríguez Delgado (2025)")
p_disc3.add_run(" y ")
add_csl_field(p_disc3, "ramirez_2025", "Ramírez (2025)")
p_disc3.add_run(
    ", quienes puntualizan que la certidumbre informativa constituye un prerrequisito psicosocial para la autorrealización profesional y la integración cooperativa."
)

# --- OBJETIVO ESPECÍFICO 3 ---
add_h2("Objetivo Específico 3: Rol del clima laboral sobre el desempeño individual y colectivo")
add_p(
    "Finalmente, se analizó el papel mediador y la incidencia directa del entorno psicosocial de trabajo sobre los indicadores de "
    "productividad y conducta técnica, cuyos coeficientes se presentan en la Tabla 4."
)

# TABLA 4
add_p("Tabla 4", bold=True, space_after=1)
add_p("Correlación de Spearman entre Clima Laboral y Desempeño Laboral", italic=True, space_after=4)

t4 = doc.add_table(rows=4, cols=4)
t4.alignment = WD_TABLE_ALIGNMENT.CENTER
apply_apa_table_borders(t4)

headers_t4 = ["Estadístico", "Variable", "Clima Laboral", "Desempeño Laboral"]
for c_i, h in enumerate(headers_t4):
    cell = t4.cell(0, c_i)
    cell.paragraphs[0].text = h
    set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=True)
    set_cell_margins(cell)

data_t4 = [
    ["Rho de Spearman", "Clima Laboral", "1,000", ",806**"],
    ["", "Sig. (bilateral)", "--", ",000"],
    ["", "N", "158", "158"]
]
for r_i, row in enumerate(data_t4):
    for c_i, val in enumerate(row):
        cell = t4.cell(r_i+1, c_i)
        cell.paragraphs[0].text = val
        bold_flag = True if "**" in val else False
        set_run_font(cell.paragraphs[0].runs[0], size_pt=10, bold=bold_flag)
        set_cell_margins(cell)

add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral). Fuente: Análisis inferencial en IBM SPSS Statistics v.26.", italic=True, size=10, space_after=8)

p_disc4 = add_p(
    "Tal como se observa en la Tabla 4, el clima laboral exhibe una correlación positiva muy fuerte y estadísticamente robusta "
    "(rs = 0,806; p = 0,000) con el desempeño laboral individual y colectivo. Dicho resultado comprueba empíricamente que el clima "
    "organizacional funge como el articulador que potencia las competencias del personal técnico, transformando el bienestar "
    "emocional en eficiencia operativa. Esta evidencia confirma plenamente los hallazgos reportados en el sur peruano por "
)
add_csl_field(p_disc4, "berolatti_2023", "Berolatti Obando (2023)")
p_disc4.add_run(", ")
add_csl_field(p_disc4, "tacuri_2023", "Tacuri Pinto (2023)")
p_disc4.add_run(" y ")
add_csl_field(p_disc4, "choquejahua_2023", "Choquejahua Olivera (2023)")
p_disc4.add_run(
    ", quienes subrayan que los colaboradores que se desenvuelven bajo un clima de equidad, supervisión formativa y reconocimiento "
    "alcanzan estándares superiores de productividad y manifiestan menor rotación laboral."
)

# ==============================================================================
# CONCLUSIONES
# ==============================================================================
add_h1("CONCLUSIONES")

add_p(
    "1. Se determinó una relación positiva considerable y estadísticamente significativa (rs = 0,743; p = 0,000) entre la comunicación "
    "interna y el desempeño laboral en profesionales y técnicos de Electro Puno S.A.A. durante el periodo 2022-2024. Se concluye que "
    "la optimización de los flujos de comunicación institucional actúa como un impulsor directo de la productividad y eficacia en las "
    "tareas del servicio eléctrico."
)

add_p(
    "2. Se constató una correlación positiva muy fuerte (rs = 0,840; p = 0,000) entre los canales y medios de comunicación interna y la "
    "percepción del clima laboral. La provisión de plataformas informativas estructuradas y accesibles disminuye la incertidumbre en las sedes "
    "zonales y fortalece el sentido de pertenencia hacia la empresa."
)

add_p(
    "3. Se evidenció una correlación positiva moderada (rs = 0,567; p = 0,000) entre la calidad del flujo comunicacional (claridad, inmediatez "
    "y retroalimentación) y el nivel de cohesión y satisfacción laboral. El intercambio asertivo de información previene controversias operativas "
    "y cohesiona los equipos multidisciplinarios en campo."
)

add_p(
    "4. Se comprobó que el clima organizacional ejerce un rol modulador decisivo sobre el rendimiento de los colaboradores, reflejado en "
    "una asociación positiva muy fuerte (rs = 0,806; p = 0,000) con el desempeño laboral. Un microclima sustentado en la confianza y el "
    "acompañamiento técnico de las jefaturas constituye el catalizador indispensable para la excelencia operativa regional."
)

# ==============================================================================
# AGRADECIMIENTOS
# ==============================================================================
add_h1("AGRADECIMIENTOS")
add_p(
    "A la Empresa Regional de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.), a sus directivos y en especial "
    "a los profesionales y técnicos de los diferentes servicios eléctricos por su valiosa disposición y colaboración activa durante la "
    "aplicación de los instrumentos de recolección de datos que hicieron viable esta investigación."
)

# ==============================================================================
# CONFLICTO DE INTERÉS
# ==============================================================================
add_h1("CONFLICTO DE INTERÉS")
add_p(
    "La autora declara que no existe ningún conflicto de interés de naturaleza económica, institucional o personal que comprometa "
    "la validez de los resultados presentados en el presente artículo científico (C.K.S.Q.)."
)

# ==============================================================================
# REFERENCIAS (APA 7ma Edición con DOIs y Enlaces)
# ==============================================================================
add_h1("REFERENCIAS")

def add_ref_item(text, url="", doi=""):
    p = add_p()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_run_font(r, size_pt=11)
    if doi:
        r_doi = p.add_run(f" https://doi.org/{doi}")
        set_run_font(r_doi, size_pt=11, color_rgb=RGBColor(0, 0, 255))
    elif url:
        r_url = p.add_run(f" {url}")
        set_run_font(r_url, size_pt=11, color_rgb=RGBColor(0, 0, 255))

# Alphabetical list of references (30 entries)
add_ref_item("Andrade, H. (2005). Comunicación Organizacional Interna: Proceso, disciplina y técnica. Gesbiblo, S.L.", url="https://dialnet.unirioja.es/servlet/libro?codigo=274945")
add_ref_item("Benavides, L. (2019). La comunicación interna y el rendimiento laboral del personal administrativo. Revista Científica de Administración y Negocios, 8(2), 45–58.", url="https://revistas.uss.edu.pe/index.php/ING/article/view/1182")
add_ref_item("Bernal, I., Pedraza, N., & Sánchez, M. (2014). El clima organizacional y su relación con la calidad de los servicios públicos. Contaduría y Administración, 59(4), 185–205.", doi="10.1016/S0186-1042(14)70160-0")
add_ref_item("Berolatti Obando, G. (2023). Clima organizacional y productividad laboral en colaboradores de instituciones de la región Puno [Tesis de Maestría, Universidad Nacional del Altiplano]. Repositorio Institucional UNAP.", url="http://repositorio.unap.edu.pe/handle/20.500.14082/19124")
add_ref_item("Capriotti, P. (2021). Planificación estratégica de la comunicación institucional. Editorial IIRP.", url="https://www.paulcapriotti.com/libros/")
add_ref_item("Castro, C. (2023). La supervisión en la productividad y el clima laboral. Gestiopolis.", url="https://www.gestiopolis.com/la-supervision-en-la-productividad-y-el-clima-laboral/")
add_ref_item("Castro, R. (2024). Gestión de la comunicación interna y desempeño del talento humano en el sector servicios. Revista Andina de Gestión Pública, 6(2), 114–128.", doi="10.36881/ragp.v6i2.214")
add_ref_item("Chiavenato, I. (2007). Evaluación del desempeño en la administración de recursos humanos. McGraw-Hill Interamericana.", url="https://www.mheducation.com.mx/")
add_ref_item("Chiavenato, I. (2009). Comportamiento organizacional: La dinámica del éxito en las organizaciones (2da ed.). McGraw-Hill Interamericana.", url="https://www.mheducation.com.mx/")
add_ref_item("Chiavenato, I. (2011). Gestión del talento humano (3ra ed.). McGraw-Hill Interamericana.", url="https://www.mheducation.com.mx/")
add_ref_item("Choquejahua Olivera, L. (2023). Clima organizacional y su relación con el desempeño laboral de los trabajadores [Tesis de Maestría, Universidad Nacional del Altiplano]. Repositorio Institucional UNAP.", url="http://repositorio.unap.edu.pe/handle/20.500.14082/19482")
add_ref_item("Espinoza, C., & Suxe, N. (2023). Comunicación interna y el desempeño laboral de los colaboradores en empresas de servicios. Revista Científica UISRAEL, 10(2), 47–63.", doi="10.35290/rcui.v10n2.2023.792")
add_ref_item("Fernández, W. (2019). Comunicación interna y el desempeño laboral del personal administrativo de entidades estatales. Revista Científica de Gestión Pública, 3(8), 24–38.", doi="10.33996/revistagestion.v3i8.31")
add_ref_item("Gómez, A., & López, A. (2022). El clima organizacional basado en la comunicación interna en una empresa de servicios. Revista Iberoamericana de Psicología y Gestión, 15(2), 85–101.", doi="10.33881/2027-1786.rip.15204")
add_ref_item("Griffin, R., & Moorhead, G. (2010). Comportamiento organizacional: Gestión de personas y organizaciones (9na ed.). Cengage Learning.", url="https://latam.cengage.com/")
add_ref_item("Hernández, R., & Mendoza, C. (2018). Metodología de la investigación: Las rutas cuantitativa, cualitativa y mixta. McGraw-Hill Interamericana.", url="https://www.mheducation.com.mx/")
add_ref_item("Hinostroza, M. (2025). Redes de comunicación interna y eficiencia operativa en empresas del sector eléctrico. Revista Peruana de Energía y Gestión, 7(1), 77–93.", doi="10.26439/rpeg2025.v7n1.512")
add_ref_item("Jibaja, M. (2020). Desempeño de los medios y comunicación interna en el Hospital Metropolitano de Quito [Tesis de Maestría, Universidad Andina Simón Bolívar]. Repositorio Institucional UASB.", url="http://hdl.handle.net/10644/7432")
add_ref_item("Llanos Torres, Y. M. (2023). Clima organizacional y desempeño laboral de los trabajadores en entidades públicas [Tesis de Maestría, Universidad Nacional del Altiplano]. Repositorio Institucional UNAP.", url="http://repositorio.unap.edu.pe/handle/20.500.14082/19854")
add_ref_item("Manzano, H. (2017). La comunicación interna y el desempeño laboral del talento humano en cooperativas de ahorro y crédito [Tesis de Maestría, Universidad Técnica de Ambato]. Repositorio Institucional UTA.", url="http://repositorio.uta.edu.ec/handle/123456789/26135")
add_ref_item("Oyarvide-Ramírez, H., Reyes-Sarría, E., & Montaño-Colorado, M. (2017). La comunicación interna como herramienta de gestión de las relaciones laborales en las organizaciones. Dominio de las Ciencias, 3(4), 718–733.", doi="10.23857/dc.v3i4.718")
add_ref_item("Palma Carrillo, S. (2004). Escala de Clima Laboral CL - SPC: Manual psicométrico. Editora y Comercializadora Cartolan.", url="https://www.worldcat.org/title/escala-de-clima-laboral-cl-spc/oclc/68903178")
add_ref_item("Pérez, A. (2023). La comunicación interna y su incidencia en el desempeño laboral de los trabajadores en empresas peruanas. Perspectiva Empresarial, 10(1), 112–126.", doi="10.16967/23898186.782")
add_ref_item("Pozo, C. (2015). Motivación y desempeño laboral en los trabajadores de empresas de manufactura [Tesis de Maestría, Universidad Técnica de Ambato]. Repositorio Institucional UTA.", url="http://repositorio.uta.edu.ec/handle/123456789/10892")
add_ref_item("Quinto Flores, J. (2023). Comunicación interna y productividad laboral en organizaciones públicas del altiplano. Revista Científica del Altiplano, 12(3), 189–204.", doi="10.26867/se.2023.v12i3.189")
add_ref_item("Ramírez, J. C. B. (2025). Comunicación interna y clima laboral: Estudio correlacional en instituciones de servicio. Revista Innovación y Empresa, 9(1), 54–69.", doi="10.18050/rie.2025.v9i1.1042")
add_ref_item("Robbins, S., & Judge, T. (2013). Comportamiento organizacional (15a ed.). Pearson Educación.", url="https://www.pearson.com/")
add_ref_item("Rodríguez Delgado, Y. A. (2025). Canales de comunicación interna y bienestar psicolaboral en entidades públicas. Revista Peruana de Ciencias Administrativas, 4(1), 88–104.", doi="10.37787/rpca.v4i1.88")
add_ref_item("Tacuri Pinto, A. (2023). Clima organizacional y el desempeño laboral de los trabajadores en empresas de servicios públicos [Tesis de Maestría, Universidad Nacional del Altiplano]. Repositorio Institucional UNAP.", url="http://repositorio.unap.edu.pe/handle/20.500.14082/19712")
add_ref_item("Toala Bozada, S., Álvarez Pincay, D., & Osejos Valencia, J. (2017). El desempeño laboral: Un enfoque teórico y operativo en las organizaciones modernas. Dominio de las Ciencias, 3(3), 612–627.", doi="10.23857/dc.v3i3.612")
add_ref_item("Vilca Quispe, F. M. (2021). Clima organizacional y satisfacción laboral en los trabajadores de la región Puno [Tesis de Maestría, Universidad Nacional del Altiplano]. Repositorio Institucional UNAP.", url="http://repositorio.unap.edu.pe/handle/20.500.14082/17621")

article_path = 'Articulo_Cientifico_Chanda_Sosa_EPG.docx'
doc.save(article_path)
print(f"Article successfully generated and saved to: {article_path}")
