# generate_complete_paper.py
# -*- coding: utf-8 -*-
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import json
import xml.sax.saxutils as saxutils
import os

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
    return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=16, space_after=6, bold=True, size=13)

def add_h2(text):
    return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=12, space_after=4, bold=True, italic=True, size=12)

def set_cell_margins(cell, top=70, bottom=70, left=100, right=100):
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
    if len(table.rows) > 0:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

def add_custom_table(headers, rows_data, col_widths=None):
    t = doc.add_table(rows=len(rows_data)+1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    apply_apa_table_borders(t)
    
    # Headers
    for c_i, h in enumerate(headers):
        cell = t.cell(0, c_i)
        cell.paragraphs[0].text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].paragraph_format.line_spacing = 1.0
        cell.paragraphs[0].paragraph_format.space_after = Pt(2)
        cell.paragraphs[0].paragraph_format.space_before = Pt(2)
        set_run_font(cell.paragraphs[0].runs[0], size_pt=9.5, bold=True)
        set_cell_margins(cell)
        if col_widths and c_i < len(col_widths):
            cell.width = Inches(col_widths[c_i])
            
    # Data rows
    for r_i, r_data in enumerate(rows_data):
        for c_i, val in enumerate(r_data):
            cell = t.cell(r_i+1, c_i)
            cell.paragraphs[0].text = str(val)
            cell.paragraphs[0].paragraph_format.line_spacing = 1.0
            cell.paragraphs[0].paragraph_format.space_after = Pt(2)
            cell.paragraphs[0].paragraph_format.space_before = Pt(2)
            if any(char.isdigit() for char in str(val)) or val in ["--", "Total", "TOTAL"]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            bold_flag = True if (r_i == len(rows_data)-1 and ("TOTAL" in str(r_data[0]).upper() or "TOTAL" in str(val).upper())) or "**" in str(val) else False
            set_run_font(cell.paragraphs[0].runs[0], size_pt=9, bold=bold_flag)
            set_cell_margins(cell)
            if col_widths and c_i < len(col_widths):
                cell.width = Inches(col_widths[c_i])
    return t

def add_figure_image(image_path, fig_num, title, note_text, width_inches=5.4):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(4)
    r_img = p_img.add_run()
    r_img.add_picture(image_path, width=Inches(width_inches))
    
    p_cap = add_p(f"Figura {fig_num}", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=1, bold=True, size=10)
    p_title = add_p(title, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=3, italic=True, size=10)
    p_note = add_p(f"Nota. {note_text}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=10, italic=True, size=9)
    return p_note

# ==============================================================================
# HEADER / TITLE SECTION
# ==============================================================================
add_p("COMUNICACIÓN INTERNA, CLIMA ORGANIZACIONAL Y DESEMPEÑO LABORAL EN PROFESIONALES Y TÉCNICOS DEL SECTOR ELÉCTRICO EN PUNO, PERÚ", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=6, bold=True, size=14)
add_p("INTERNAL COMMUNICATION, ORGANIZATIONAL CLIMATE, AND JOB PERFORMANCE AMONG PROFESSIONAL AND TECHNICAL STAFF IN THE ELECTRICAL SECTOR IN PUNO, PERU", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=14, bold=True, italic=True, size=12)

add_p("Chanda Ketty Sosa Quispe", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, bold=True, size=11)
add_p("Maestría en Contabilidad y Administración, Escuela de Posgrado, Universidad Nacional del Altiplano, Puno, Perú\nORCID: https://orcid.org/0009-0004-9218-4720 | Correo: chanda.sosa@unap.edu.pe", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14, italic=True, size=10)

# ==============================================================================
# RESUMEN Y PALABRAS CLAVE
# ==============================================================================
add_h1("RESUMEN")
add_p(
    "El propósito fundamental de la investigación consistió en determinar la relación entre la comunicación interna, "
    "el clima organizacional y el desempeño laboral exhibido por el personal técnico y profesional de la Empresa Regional "
    "de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.) durante el periodo 2022-2024. La metodología "
    "respondió a un enfoque cuantitativo, de nivel descriptivo-correlacional y diseño no experimental transversal. Se trabajó "
    "con una muestra censal integrada por 158 trabajadores clasificados en los grupos ocupacionales profesionales y técnicos "
    "distribuidos en 19 servicios eléctricos y sedes zonales del altiplano puneño. Para el levantamiento empírico se aplicó una "
    "batería psicométrica estructurada en escala Likert validada mediante juicio de expertos, demostrando una consistencia interna "
    "sobresaliente (Alfa de Cronbach de 0,876 para Comunicación Interna; 0,927 para Clima Laboral; y 0,846 para Desempeño Laboral). "
    "Las pruebas de Kolmogórov-Smirnov y Shapiro-Wilk confirmaron un comportamiento no paramétrico (p < 0,05), por lo que se empleó "
    "el estadístico Rho de Spearman para la contrastación de hipótesis. Los hallazgos revelaron una correlación positiva considerable "
    "y estadísticamente significativa entre la comunicación interna y el desempeño laboral (rs = 0,743; p = 0,000). Asimismo, se "
    "constataron asociaciones directas muy fuertes entre los canales comunicacionales y el clima laboral (rs = 0,840; p = 0,000), "
    "una correlación moderada entre la calidad del flujo informativo y la satisfacción laboral (rs = 0,567; p = 0,000), y una correlación "
    "positiva muy fuerte entre el clima organizacional y el desempeño (rs = 0,806; p = 0,000). Se concluye que un sistema comunicacional "
    "interno estructurado, ágil y multidireccional consolida un entorno de trabajo armónico y potencia la productividad en el servicio eléctrico."
)

p_kw_es = add_p("Palabras clave: ", bold=True, space_after=12)
r_kw = p_kw_es.add_run("Clima organizacional, comunicación interna, desempeño laboral, eficiencia técnica, sector eléctrico.")
set_run_font(r_kw, italic=True)

# ==============================================================================
# ABSTRACT Y KEYWORDS
# ==============================================================================
add_h1("ABSTRACT")
add_p(
    "The central aim of this investigation was to determine the relationship between internal communication dynamics, "
    "organizational climate, and job performance among professional and technical staff at the Empresa Regional de Servicios "
    "Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.) during the 2022–2024 period. Methodologically, the study "
    "followed a quantitative paradigm with a descriptive-correlational scope and a non-experimental, cross-sectional design. "
    "A census sample comprising 158 professional and technical employees across 19 regional electrical divisions and branch "
    "offices in the Puno high plateau was evaluated. Data collection employed a five-point Likert psychometric questionnaire "
    "validated by expert panel judgment, exhibiting exceptional internal reliability (Cronbach's alpha of 0.876 for Internal "
    "Communication, 0.927 for Organizational Climate, and 0.846 for Job Performance). Normality testing through Kolmogorov-Smirnov "
    "and Shapiro-Wilk statistics rejected Gaussian distribution (p < 0.05), warranting non-parametric Spearman's Rho correlation. "
    "Results demonstrated a substantial and statistically significant positive correlation between internal communication and job "
    "performance (rs = 0.743; p = 0.000). Furthermore, strong associations emerged between institutional communication media and "
    "workplace climate perception (rs = 0.840; p = 0.000), a moderate correlation between informational flow quality and employee "
    "satisfaction (rs = 0.567; p = 0.000), and a very strong correlation between organizational climate and operational productivity "
    "(rs = 0.806; p = 0.000). In conclusion, strengthening multidirectional internal communication channels optimizes the socio-labor "
    "ecosystem and drives operational excellence in public electrical distribution services."
)

p_kw_en = add_p("Keywords: ", bold=True, space_after=14)
r_kw_en = p_kw_en.add_run("Electrical sector, internal communication, job performance, organizational climate, technical efficiency.")
set_run_font(r_kw_en, italic=True)

# ==============================================================================
# INTRODUCCIÓN
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
p_intro1.add_run(
    ". No obstante, diversas investigaciones empíricas demuestran que las fallas en la difusión de información originan distorsiones "
    "que debilitan el clima laboral y menoscaban el rendimiento productivo de los colaboradores."
)

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
    ". Cuando las instituciones logran articular canales formales accesibles y confiables, los colaboradores experimentan mayor "
    "certidumbre funcional y claridad en sus roles laborales "
)
add_csl_field(p_intro2, "fernandez_2019", "(Fernández, 2019)")
p_intro2.add_run(
    ". Por el contrario, la opacidad informativa favorece la propagación de rumores informales y tensiones interpersonales."
)

p_intro3 = add_p(
    "En estrecha correspondencia, el clima organizacional alude a la configuración de percepciones globales que los colaboradores "
    "construyen en torno a las políticas, prácticas y procedimientos de su entorno cotidiano "
)
add_csl_field(p_intro3, "palma_2004", "(Palma Carrillo, 2004)")
p_intro3.add_run(
    ". Dimensiones cardinales como la autorrealización personal, el involucramiento laboral, la supervisión de soporte, la comunicación "
    "y las condiciones físicas de trabajo modulan la experiencia subjetiva del colaborador "
)
add_csl_field(p_intro3, "bernal_2014", "(Bernal et al., 2014)")
p_intro3.add_run(
    ". Un ambiente institucional favorable fomenta la reciprocidad y la confianza mutua, mientras que un microclima hostil propicia "
    "desmotivación, ausentismo y deterioro en la calidad de la atención al usuario "
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
    "La investigación se ejecutó en la Empresa Regional de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.), "
    "organización estatal de derecho privado adscrita al Fondo Nacional de Financiamiento de la Actividad Empresarial del Estado (FONAFE). "
    "Su radio de acción comprende la totalidad del departamento de Puno, con sede central administrativa en la ciudad de Puno y "
    "centros de operaciones técnicos distribuidos en las zonas norte, centro y sur del altiplano peruano."
)

add_h2("Descripción de métodos y diseño de la investigación")
add_p(
    "El estudio se desarrolló bajo el paradigma cuantitativo, tipología aplicada, nivel descriptivo-correlacional y un diseño no experimental "
    "de corte transversal, recolectando la información en un único momento temporal representativo del periodo 2022-2024 "
)
add_csl_field(doc.paragraphs[-1], "hernandez_2018", "(Hernández & Mendoza, 2018)")
doc.paragraphs[-1].add_run(
    ". No se alteraron intencionalmente las variables, observándose los fenómenos en su dinámica laboral cotidiana."
)

# --- TABLA 1: MUESTRA CENSAL ---
add_p(
    "En la Tabla 1 se presenta la conformación de la muestra censal de la investigación, distribuida de acuerdo con la clasificación "
    "ocupacional y la localización geográfica de los 19 servicios eléctricos y centros zonales de Electro Puno S.A.A."
)

add_p("Tabla 1", bold=True, space_after=1)
add_p("Distribución de la Población y Muestra Censal por Servicios Eléctricos y Sedes Zonales", italic=True, space_after=4)

headers_sample = ["Sede / Servicio Eléctrico", "Profesionales", "Técnicos", "Total Trabajadores", "Porcentaje (%)"]
data_sample = [
    ["Ananea", "1", "3", "4", "2,53%"],
    ["Antauta", "1", "3", "4", "2,53%"],
    ["Ayaviri (Melgar)", "1", "2", "3", "1,90%"],
    ["Azángaro", "1", "3", "4", "2,53%"],
    ["Desaguadero", "1", "3", "4", "2,53%"],
    ["Huancané", "0", "2", "2", "1,27%"],
    ["Ilave (El Collao)", "1", "3", "4", "2,53%"],
    ["Juli (Chucuito)", "0", "2", "2", "1,27%"],
    ["Juliaca (San Román)", "16", "10", "26", "16,46%"],
    ["Lampa", "0", "3", "3", "1,90%"],
    ["Moho", "1", "3", "4", "2,53%"],
    ["Ollachea (Carabaya)", "0", "1", "1", "0,63%"],
    ["Puno (Sede Central y Operativa)", "68", "14", "82", "51,90%"],
    ["Putina (San Antonio de Putina)", "0", "3", "3", "1,90%"],
    ["San Gabán", "0", "1", "1", "0,63%"],
    ["San Juan del Oro", "0", "1", "1", "0,63%"],
    ["Sandia", "0", "6", "6", "3,80%"],
    ["Taquile", "0", "1", "1", "0,63%"],
    ["Yunguyo", "0", "3", "3", "1,90%"],
    ["TOTAL", "91", "67", "158", "100,00%"]
]
add_custom_table(headers_sample, data_sample, [2.3, 1.1, 1.0, 1.3, 1.1])
add_p("Nota. Datos consolidados de las nóminas del Departamento de Recursos Humanos de Electro Puno S.A.A. (2022–2024).", italic=True, size=9.5, space_after=8)

# --- TABLA 2: OPERACIONALIZACIÓN DE VARIABLES ---
add_p(
    "Para garantizar la precisión en el abordaje empírico y la alineación entre las preguntas de investigación, los constructos "
    "teóricos y los indicadores medidos, en la Tabla 2 se detalla la matriz de operacionalización de variables."
)

add_p("Tabla 2", bold=True, space_after=1)
add_p("Matriz de Operacionalización de Variables, Dimensiones e Indicadores", italic=True, space_after=4)

headers_op = ["Objetivo Específico", "Variable de Estudio", "Dimensiones", "Indicadores Medidos"]
data_op = [
    ["OE1: Analizar la incidencia de canales y medios comunicacionales en el clima laboral.", "Canales y Medios de Comunicación Interna (X1)", "Canales de Comunicación", "Canales formales/informales; Accesibilidad a la información institucional; Frecuencia de uso."],
    ["", "Percepción del Clima Laboral (Z1)", "Vínculo Organizacional", "Satisfacción con la comunicación interna; Confianza institucional; Pertenencia organizacional."],
    ["OE2: Evaluar la relación entre la calidad del flujo comunicacional y la cohesión y satisfacción.", "Calidad del Flujo Comunicacional (X2)", "Calidad de Comunicación", "Claridad del mensaje técnico; Oportunidad temporal de la información; Retroalimentación efectiva bidireccional."],
    ["", "Satisfacción y Cohesión Laboral (Z2)", "Cohesión Grupal y Satisfacción", "Nivel de cohesión entre áreas; Trabajo colaborativo; Satisfacción en el puesto de trabajo."],
    ["OE3: Examinar la incidencia del clima laboral en el desempeño individual y colectivo.", "Clima Laboral como Variable Mediadora (Z)", "Ambiente Laboral", "Percepción del ambiente físico y psicosocial; Relaciones interpersonales; Apoyo jerárquico."],
    ["", "Desempeño Laboral (Y)", "Motivación y Productividad", "Motivación laboral intrínseca; Eficiencia en recursos; Eficacia en tareas; Conducta y productividad colectiva."]
]
add_custom_table(headers_op, data_op, [2.0, 1.6, 1.4, 2.0])
add_p("Nota. Elaboración propia en función del marco metodológico cuantitativo de la Escuela de Posgrado de la UNA Puno.", italic=True, size=9.5, space_after=8)

# --- TABLA 3: PROPIEDADES PSICOMÉTRICAS Y CONFIABILIDAD ---
add_p(
    "Los instrumentos fueron sometidos a una rigurosa evaluación de propiedades psicométricas. En la Tabla 3 se compendian "
    "las valoraciones de validez por juicio de expertos y los índices de confiabilidad mediante Alfa de Cronbach."
)

add_p("Tabla 3", bold=True, space_after=1)
add_p("Evaluación Psicométrica: Validez de Contenido y Consistencia Interna por Escalas", italic=True, space_after=4)

headers_val = ["Variable / Escala Analizada", "N° Ítems", "Alfa de Cronbach (α)", "Juicio de Expertos (Validez)", "Criterio Psicométrico"]
data_val = [
    ["Comunicación Interna (General)", "13", ",876", "100% de Aprobación Favorable", "Consistencia Interna Fuerte"],
    ["- Canales y Medios de Comunicación", "7", ",936", "Excelente Pertinencia y Claridad", "Consistencia Interna Excelente"],
    ["- Calidad de Flujo Comunicacional", "6", ",786", "Alta Coherencia Teórica", "Consistencia Interna Favorable"],
    ["Clima Laboral (General)", "16", ",927", "100% de Aprobación Favorable", "Consistencia Interna Excelente"],
    ["- Vínculo Organizacional", "6", ",786", "Excelente Relevancia de Ítems", "Consistencia Interna Favorable"],
    ["- Satisfacción y Cohesión Laboral", "10", ",927", "Excelente Representatividad", "Consistencia Interna Excelente"],
    ["Desempeño Laboral (General)", "15", ",846", "100% de Aprobación Favorable", "Consistencia Interna Fuerte"],
    ["- Eficiencia y Eficacia Operativa", "8", ",846", "Alta Claridad Conceptual", "Consistencia Interna Fuerte"],
    ["- Conducta y Motivación Laboral", "7", ",846", "Excelente Pertinencia Práctica", "Consistencia Interna Fuerte"]
]
add_custom_table(headers_val, data_val, [2.3, 0.7, 1.2, 1.6, 1.2])
add_p("Nota. Coeficientes procesados en IBM SPSS Statistics versión 26 a partir de las matrices del cuestionario de 158 participantes.", italic=True, size=9.5, space_after=8)

# --- TABLA 4: PRUEBAS DE NORMALIDAD ---
add_p(
    "Para determinar justificadamente la pertinencia de pruebas estadísticas paramétricas o no paramétricas, se aplicaron "
    "las pruebas de bondad de ajuste de Kolmogórov-Smirnov (para n ≥ 50) y Shapiro-Wilk (Tabla 4)."
)

add_p("Tabla 4", bold=True, space_after=1)
add_p("Pruebas de Normalidad de Kolmogórov-Smirnov y Shapiro-Wilk para las Variables del Estudio", italic=True, space_after=4)

headers_norm = ["Variable / Dimensión Evaluada", "Kolmogórov-Smirnov (Estadístico)", "gl", "Sig. (p)", "Shapiro-Wilk (Estadístico)", "gl", "Sig. (p)", "Decisión Distribucional"]
data_norm = [
    ["Comunicación Interna", ",122", "158", ",000", ",945", "158", ",000", "No paramétrico (p < 0,05)"],
    ["Desempeño Laboral", ",144", "158", ",000", ",950", "158", ",000", "No paramétrico (p < 0,05)"],
    ["Canales y Medios de Comunicación", ",131", "158", ",000", ",938", "158", ",000", "No paramétrico (p < 0,05)"],
    ["Percepción del Clima Laboral", ",152", "158", ",000", ",922", "158", ",000", "No paramétrico (p < 0,05)"],
    ["Calidad del Flujo Comunicacional", ",128", "158", ",000", ",941", "158", ",000", "No paramétrico (p < 0,05)"],
    ["Satisfacción y Cohesión Laboral", ",165", "158", ",000", ",910", "158", ",000", "No paramétrico (p < 0,05)"],
    ["Clima Laboral Mediador", ",139", "158", ",000", ",934", "158", ",000", "No paramétrico (p < 0,05)"]
]
add_custom_table(headers_norm, data_norm, [1.8, 0.9, 0.4, 0.6, 0.9, 0.4, 0.6, 1.4])
add_p("Nota. p < 0,05 rechaza la hipótesis de normalidad H₀, justificando el uso de la correlación ordinal Rho de Spearman.", italic=True, size=9.5, space_after=8)

# ==============================================================================
# RESULTADOS Y DISCUSIÓN
# ==============================================================================
add_h1("RESULTADOS Y DISCUSIÓN")

# --- SECCIÓN 1: OBJETIVO GENERAL ---
add_h2("1. Análisis del Objetivo General: Comunicación interna y desempeño laboral")
add_p(
    "El primer eje de análisis contrastó la hipótesis general de investigación sobre la muestra de 158 colaboradores. "
    "En la Tabla 5 se resume la distribución de frecuencias y porcentajes para las dos variables medulares."
)

# TABLA 5: FRECUENCIAS COMUNICACIÓN Y DESEMPEÑO
add_p("Tabla 5", bold=True, space_after=1)
add_p("Distribución de Frecuencias y Porcentajes de Comunicación Interna y Desempeño Laboral", italic=True, space_after=4)

headers_f1 = ["Nivel de Escala", "Rango Puntaje (Com.)", "Frecuencia (Com.)", "Porcentaje (Com.)", "Rango Puntaje (Des.)", "Frecuencia (Des.)", "Porcentaje (Des.)"]
data_f1 = [
    ["Totalmente de acuerdo", "30 - 35", "15", "9,49%", "39 - 45", "31", "19,62%"],
    ["De acuerdo", "25 - 29", "32", "20,25%", "32 - 38", "64", "40,51%"],
    ["Indiferente", "19 - 24", "54", "34,18%", "24 - 31", "63", "39,87%"],
    ["Poco de acuerdo", "14 - 18", "57", "36,08%", "17 - 23", "0", "0,00%"],
    ["Nada de acuerdo", "7 - 13", "0", "0,00%", "9 - 16", "0", "0,00%"],
    ["TOTAL", "--", "158", "100,00%", "--", "158", "100,00%"]
]
add_custom_table(headers_f1, data_f1, [1.6, 1.0, 0.8, 1.0, 1.0, 0.8, 1.0])
add_p("Nota. Com. = Comunicación Interna; Des. = Desempeño Laboral. Datos recabados en Electro Puno S.A.A.", italic=True, size=9.5, space_after=6)

add_p(
    "Como se aprecia en la Tabla 5, el 36,08% de los participantes se ubicó en el nivel 'poco de acuerdo' y el 34,18% 'indiferente' "
    "respecto a la comunicación interna, mientras que en el desempeño laboral el 40,51% se ubicó 'de acuerdo' y el 39,87% 'indiferente'. "
    "Para comprobar la relación bivariada, en la Tabla 6 se exponen los coeficientes de correlación de Spearman."
)

# TABLA 6: CORRELACIÓN OG
add_p("Tabla 6", bold=True, space_after=1)
add_p("Matriz de Correlación de Spearman entre Comunicación Interna y Desempeño Laboral", italic=True, space_after=4)

headers_c1 = ["Estadístico", "Variables", "Comunicación Interna", "Desempeño Laboral"]
data_c1 = [
    ["Rho de Spearman", "Comunicación Interna", "1,000", ",743**"],
    ["", "Sig. bilateral (p-valor)", "--", ",000"],
    ["", "N", "158", "158"],
    ["", "Desempeño Laboral", ",743**", "1,000"],
    ["", "Sig. bilateral (p-valor)", ",000", "--"],
    ["", "N", "158", "158"]
]
add_custom_table(headers_c1, data_c1, [1.5, 1.9, 1.8, 1.8])
add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral). Fuente: Procesamiento en SPSS v.26.", italic=True, size=9.5, space_after=6)

# FIGURA 1
add_p(
    "La correspondencia lineal ascendente y la distribución muestral de los puntajes bivariados se representan visualmente "
    "en el diagrama de dispersión de la Figura 1."
)
add_figure_image(
    "extracted_figures/image2.png", 1,
    "Diagrama de dispersión bivariado entre Comunicación Interna y Desempeño Laboral",
    "Nube de dispersión generada en IBM SPSS v.26 que ilustra la tendencia positiva considerable (rs = 0,743; p = 0,000) entre la comunicación corporativa y el desempeño del personal técnico y profesional en Electro Puno S.A.A."
)

p_disc1 = add_p(
    "Discusión del Objetivo General:\n"
    "Los resultados inferenciales confirman la existencia de una correlación positiva considerable y estadísticamente "
    "significativa entre la comunicación interna y el desempeño laboral (rs = 0,743; p = 0,000). Al ser el p-valor inferior al "
    "nivel de significancia establecido (p < 0,01), se rechaza categóricamente la hipótesis nula, concluyéndose que una comunicación "
    "corporativa articulada dinamiza la productividad de los colaboradores. Estos hallazgos concuerdan con lo reportado por "
)
add_csl_field(p_disc1, "quinto_2023", "Quinto Flores (2023)")
p_disc1.add_run(
    ", quien comprobó que la fluidez comunicativa predice favorablemente la eficiencia del capital humano en entidades públicas del altiplano. De igual forma, respalda las formulaciones teóricas de "
)
add_csl_field(p_disc1, "andrade_2005", "Andrade (2005)")
p_disc1.add_run(" y ")
add_csl_field(p_disc1, "fernandez_2019", "Fernández (2019)")
p_disc1.add_run(
    ", en cuanto a que los canales formales clarifican las funciones del puesto, minimizan los tiempos muertos y potencian la motivación intrínseca hacia el cumplimiento de metas."
)

# --- SECCIÓN 2: OBJETIVO ESPECÍFICO 1 ---
add_h2("2. Objetivo Específico 1: Canales y medios de comunicación interna y percepción del clima laboral")
add_p(
    "En lo atinente al primer objetivo específico, se buscó evaluar la correspondencia entre la provisión de canales comunicacionales "
    "y el entorno socioemocional percibido por los colaboradores. En la Tabla 7 se presentan las distribuciones de frecuencia relativas."
)

# TABLA 7: FRECUENCIAS OE1
add_p("Tabla 7", bold=True, space_after=1)
add_p("Distribución de Frecuencias de Canales y Medios de Comunicación Interna y Percepción del Clima Laboral", italic=True, space_after=4)

headers_f2 = ["Niveles de Valoración", "Rango (Canales)", "Frecuencia (Canales)", "Porcentaje (Canales)", "Rango (Clima)", "Frecuencia (Clima)", "Porcentaje (Clima)"]
data_f2 = [
    ["Totalmente de acuerdo", "56 - 65", "7", "4,43%", "14 - 15", "11", "6,96%"],
    ["De acuerdo", "45 - 55", "53", "33,54%", "11 - 13", "31", "19,62%"],
    ["Indiferente", "35 - 44", "37", "23,42%", "9 - 10", "45", "28,48%"],
    ["Poco de acuerdo", "24 - 34", "51", "32,28%", "6 - 8", "53", "33,54%"],
    ["Nada de acuerdo", "13 - 23", "10", "6,33%", "3 - 5", "18", "11,39%"],
    ["TOTAL", "--", "158", "100,00%", "--", "158", "100,00%"]
]
add_custom_table(headers_f2, data_f2, [1.6, 1.0, 0.8, 1.0, 1.0, 0.8, 1.0])
add_p("Nota. Canales = Canales y Medios de Comunicación Interna; Clima = Percepción del Clima Laboral.", italic=True, size=9.5, space_after=6)

add_p(
    "La evidencia de la Tabla 7 denota que un porcentaje relevante de trabajadores (32,28% en canales y 33,54% en clima) manifiesta "
    "reparos o distanciamiento respecto a los medios informativos actuales. En la Tabla 8 se presenta la contrastación inferencial de Spearman."
)

# TABLA 8: CORRELACIÓN OE1
add_p("Tabla 8", bold=True, space_after=1)
add_p("Matriz de Correlación de Spearman entre Canales y Medios y Percepción del Clima Laboral", italic=True, space_after=4)

headers_c2 = ["Estadístico", "Variables", "Canales y Medios Comunicativos", "Percepción del Clima Laboral"]
data_c2 = [
    ["Rho de Spearman", "Canales y Medios Comunicativos", "1,000", ",840**"],
    ["", "Sig. bilateral (p-valor)", "--", ",000"],
    ["", "N", "158", "158"],
    ["", "Percepción del Clima Laboral", ",840**", "1,000"],
    ["", "Sig. bilateral (p-valor)", ",000", "--"],
    ["", "N", "158", "158"]
]
add_custom_table(headers_c2, data_c2, [1.5, 1.9, 1.8, 1.8])
add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral).", italic=True, size=9.5, space_after=6)

# FIGURA 2
add_p(
    "En la Figura 2 se grafica el diagrama de dispersión que sintetiza la asociación directa observada entre la disponibilidad "
    "de canales formales y la percepción del clima de trabajo."
)
add_figure_image(
    "extracted_figures/image3.png", 2,
    "Diagrama de dispersión bivariado entre Canales y Medios de Comunicación Interna y Percepción del Clima Laboral",
    "Diagrama de correlación de Spearman (rs = 0,840; p = 0,000) generado en SPSS v.26, ratificando la relación positiva muy fuerte entre la accesibilidad a canales informativos y el microclima laboral en las dependencias eléctricas."
)

p_disc2 = add_p(
    "Discusión del Objetivo Específico 1:\n"
    "El valor del coeficiente de correlación obtenido (rs = 0,840; p = 0,000) revela una interdependencia positiva muy fuerte "
    "y estadísticamente robusta. Dicho resultado indica que en una organización de suministro eléctrico donde coexisten dependencias "
    "desconcentradas en provincias andinas alejadas de la sede central, los canales informativos operan como el pegamento institucional "
    "que preserva la confianza y la certidumbre funcional. Este hallazgo valida empíricamente las conclusiones de "
)
add_csl_field(p_disc2, "espinoza_2023", "Espinoza y Suxe (2023)")
p_disc2.add_run(
    ", quienes comprobaron que los canales descendentes y horizontales mitigan la fragmentación entre áreas operativas y directivas. Asimismo, concuerda con las observaciones de "
)
add_csl_field(p_disc2, "hinostroza_2025", "Hinostroza (2025)")
p_disc2.add_run(
    ", quien puntualizó en el sector de energía que la modernización de los medios digitales institucionales fortalece la cohesión del clima laboral."
)

# --- SECCIÓN 3: OBJETIVO ESPECÍFICO 2 ---
add_h2("3. Objetivo Específico 2: Calidad del flujo comunicacional y nivel de cohesión y satisfacción laboral")
add_p(
    "El segundo objetivo específico evaluó la calidad intrínseca del flujo informativo (claridad, inmediatez y retroalimentación) "
    "frente a la cohesión de los equipos y la satisfacción en el trabajo. En la Tabla 9 se detallan las frecuencias observadas."
)

# TABLA 9: FRECUENCIAS OE2
add_p("Tabla 9", bold=True, space_after=1)
add_p("Distribución de Frecuencias de Calidad del Flujo Comunicacional y Satisfacción y Cohesión Laboral", italic=True, space_after=4)

headers_f3 = ["Niveles de Escala", "Rango (Calidad Flujo)", "Frecuencia (Flujo)", "Porcentaje (Flujo)", "Rango (Satisfacción)", "Frecuencia (Satisf.)", "Porcentaje (Satisf.)"]
data_f3 = [
    ["Totalmente de acuerdo", "30 - 35", "15", "9,49%", "9 - 10", "13", "8,23%"],
    ["De acuerdo", "25 - 29", "32", "20,25%", "8", "45", "28,48%"],
    ["Indiferente", "19 - 24", "54", "34,18%", "6 - 7", "49", "31,01%"],
    ["Poco de acuerdo", "14 - 18", "57", "36,08%", "5", "3", "1,90%"],
    ["Nada de acuerdo", "7 - 13", "0", "0,00%", "2 - 4", "48", "30,38%"],
    ["TOTAL", "--", "158", "100,00%", "--", "158", "100,00%"]
]
add_custom_table(headers_f3, data_f3, [1.6, 1.0, 0.8, 1.0, 1.0, 0.8, 1.0])
add_p("Nota. Flujo = Calidad del Flujo Comunicacional; Satisf. = Satisfacción y Cohesión Laboral.", italic=True, size=9.5, space_after=6)

add_p(
    "Se advierte en la Tabla 9 que un 30,38% de los técnicos se manifiesta 'nada de acuerdo' con los niveles de cohesión y satisfacción actual, "
    "lo cual guarda estrecha vinculación con el 36,08% que encuentra deficiencias en el flujo comunicacional. La Tabla 10 exhibe la correlación estadística."
)

# TABLA 10: CORRELACIÓN OE2
add_p("Tabla 10", bold=True, space_after=1)
add_p("Matriz de Correlación de Spearman entre Calidad del Flujo Comunicacional y Satisfacción y Cohesión Laboral", italic=True, space_after=4)

headers_c3 = ["Estadístico", "Variables", "Calidad del Flujo Comunicacional", "Satisfacción y Cohesión Laboral"]
data_c3 = [
    ["Rho de Spearman", "Calidad del Flujo Comunicacional", "1,000", ",567**"],
    ["", "Sig. bilateral (p-valor)", "--", ",000"],
    ["", "N", "158", "158"],
    ["", "Satisfacción y Cohesión Laboral", ",567**", "1,000"],
    ["", "Sig. bilateral (p-valor)", ",000", "--"],
    ["", "N", "158", "158"]
]
add_custom_table(headers_c3, data_c3, [1.5, 1.9, 1.8, 1.8])
add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral).", italic=True, size=9.5, space_after=6)

# FIGURA 3
add_p(
    "La correspondencia bivariada entre la calidad informativa y la satisfacción de los colaboradores se visualiza en la nube de dispersión de la Figura 3."
)
add_figure_image(
    "extracted_figures/image4.png", 3,
    "Diagrama de dispersión bivariado entre Calidad del Flujo Comunicacional y Satisfacción y Cohesión Laboral",
    "Gráfico de dispersión generado en IBM SPSS v.26 que ilustra el coeficiente positivo moderado (rs = 0,567; p = 0,000) entre la claridad informativa y la satisfacción laboral en Electro Puno S.A.A."
)

p_disc3 = add_p(
    "Discusión del Objetivo Específico 2:\n"
    "La prueba de hipótesis corroboró una correlación positiva moderada-considerable (rs = 0,567; p = 0,000). Este resultado evidencia "
    "que la ausencia de mensajes ambiguos, la oportunidad cronológica de las directivas técnicas y la presencia de retroalimentación "
    "constructiva actúan como inductores directos de la satisfacción y el espíritu de cuerpo entre los técnicos de campo. Lo hallado concuerda con "
)
add_csl_field(p_disc3, "rodriguez_2025", "Rodríguez Delgado (2025)")
p_disc3.add_run(" y ")
add_csl_field(p_disc3, "ramirez_2025", "Ramírez (2025)")
p_disc3.add_run(
    ", quienes comprobaron que la retroalimentación asertiva disminuye el estrés laboral y consolida equipos multidisciplinarios de alto rendimiento. Del mismo modo, convalida la teoría de "
)
add_csl_field(p_disc3, "oyarvide_2017", "Oyarvide-Ramírez et al. (2017)")
p_disc3.add_run(
    ", en el sentido de que la información de calidad evita malentendidos procedimentales en tareas de alta peligrosidad técnica como el mantenimiento de redes de alta y media tensión."
)

# --- SECCIÓN 4: OBJETIVO ESPECÍFICO 3 ---
add_h2("4. Objetivo Específico 3: Rol del clima laboral sobre el desempeño individual y colectivo")
add_p(
    "El tercer objetivo específico analizó la incidencia del microclima laboral sobre los niveles de rendimiento y motivación individual "
    "y colectiva de los profesionales y técnicos. En la Tabla 11 se presentan las frecuencias relativas."
)

# TABLA 11: FRECUENCIAS OE3
add_p("Tabla 11", bold=True, space_after=1)
add_p("Distribución de Frecuencias de Clima Laboral y Desempeño Laboral (Objetivo Específico 3)", italic=True, space_after=4)

headers_f4 = ["Niveles de Escala", "Rango (Clima)", "Frecuencia (Clima)", "Porcentaje (Clima)", "Rango (Desempeño)", "Frecuencia (Des.)", "Porcentaje (Des.)"]
data_f4 = [
    ["Totalmente de acuerdo", "18 - 20", "19", "12,03%", "47 - 55", "23", "14,56%"],
    ["De acuerdo", "15 - 17", "35", "22,15%", "39 - 46", "61", "38,61%"],
    ["Indiferente", "11 - 14", "36", "22,78%", "29 - 38", "74", "46,84%"],
    ["Poco de acuerdo", "8 - 10", "64", "40,51%", "21 - 28", "0", "0,00%"],
    ["Nada de acuerdo", "4 - 7", "4", "2,53%", "11 - 20", "0", "0,00%"],
    ["TOTAL", "--", "158", "100,00%", "--", "158", "100,00%"]
]
add_custom_table(headers_f4, data_f4, [1.6, 1.0, 0.8, 1.0, 1.0, 0.8, 1.0])
add_p("Nota. Clima = Clima Laboral como Variable Mediadora; Des. = Desempeño Laboral Individual y Colectivo.", italic=True, size=9.5, space_after=6)

add_p(
    "De la Tabla 11 se desprende que el 40,51% califica su clima laboral como 'poco de acuerdo', mientras que el desempeño se concentra "
    "entre 'de acuerdo' (38,61%) e 'indiferente' (46,84%). En la Tabla 12 se exhibe el coeficiente de correlación inferencial."
)

# TABLA 12: CORRELACIÓN OE3
add_p("Tabla 12", bold=True, space_after=1)
add_p("Matriz de Correlación de Spearman entre Clima Laboral y Desempeño Laboral", italic=True, space_after=4)

headers_c4 = ["Estadístico", "Variables", "Clima Laboral", "Desempeño Laboral"]
data_c4 = [
    ["Rho de Spearman", "Clima Laboral", "1,000", ",806**"],
    ["", "Sig. bilateral (p-valor)", "--", ",000"],
    ["", "N", "158", "158"],
    ["", "Desempeño Laboral", ",806**", "1,000"],
    ["", "Sig. bilateral (p-valor)", ",000", "--"],
    ["", "N", "158", "158"]
]
add_custom_table(headers_c4, data_c4, [1.5, 1.9, 1.8, 1.8])
add_p("Nota. ** La correlación es significativa en el nivel 0,01 (bilateral).", italic=True, size=9.5, space_after=6)

# FIGURA 4
add_p(
    "El vínculo estadístico no paramétrico muy fuerte entre el clima laboral y el desempeño se comprueba gráficamente en la Figura 4."
)
add_figure_image(
    "extracted_figures/image5.png", 4,
    "Diagrama de dispersión bivariado entre Clima Laboral y Desempeño Laboral",
    "Diagrama de dispersión en IBM SPSS v.26 que comprueba la correlación positiva muy fuerte (rs = 0,806; p = 0,000) entre la atmósfera psicosocial de trabajo y el rendimiento de profesionales y técnicos en Electro Puno S.A.A."
)

p_disc4 = add_p(
    "Discusión del Objetivo Específico 3:\n"
    "El coeficiente de correlación de Spearman rs = 0,806 con p = 0,000 demuestra una relación positiva muy fuerte y estadísticamente "
    "robusta entre el clima organizacional y el desempeño laboral. Este hallazgo es de suma trascendencia, pues confirma que el clima "
    "laboral actúa como una variable mediadora indispensable: cuando el entorno de trabajo ofrece reconocimiento, equidad en la supervisión "
    "y soporte socioemocional, las capacidades cognitivas y destrezas de los técnicos se transforman en mayor productividad, continuidad del "
    "suministro y reducción de contingencias eléctricas. Dicha constatación empírica concuerda plenamente con los estudios desarrollados en Puno por "
)
add_csl_field(p_disc4, "berolatti_2023", "Berolatti Obando (2023)")
p_disc4.add_run(", ")
add_csl_field(p_disc4, "tacuri_2023", "Tacuri Pinto (2023)")
p_disc4.add_run(", ")
add_csl_field(p_disc4, "choquejahua_2023", "Choquejahua Olivera (2023)")
p_disc4.add_run(" y ")
add_csl_field(p_disc4, "llanos_2023", "Llanos Torres (2023)")
p_disc4.add_run(
    ", quienes concluyeron unánimemente que el microclima laboral explica de forma determinante el rendimiento y la calidad del servicio en organizaciones públicas del sur del Perú."
)

# --- SECCIÓN 5: IMPLICANCIAS TEÓRICAS Y PRÁCTICAS ---
add_h2("5. Implicancias Teóricas, Organizacionales y Propuesta de Gestión")
p_imp = add_p(
    "En términos teóricos, la investigación valida la aplicabilidad de los modelos integradores de relaciones humanas y teoría de sistemas "
)
add_csl_field(p_imp, ["capriotti_2021", "chiavenato_2009"], "(Capriotti, 2021; Chiavenato, 2009)")
p_imp.add_run(
    " en el sector de servicios energéticos regulados. Desde la perspectiva práctica, los resultados evidencian la imperiosa necesidad "
    "de que Electro Puno S.A.A. implemente un Plan Estratégico de Comunicación Interna enfocado en modernizar los canales digitales "
    "(plataformas colaborativas móviles accesibles desde tablets de campo), estandarizar protocolos de retroalimentación constructiva "
    "e institucionalizar programas semestrales de evaluación del desempeño por competencias asociados a incentivos al mérito profesional."
)

# ==============================================================================
# CONCLUSIONES
# ==============================================================================
add_h1("CONCLUSIONES")

add_p(
    "1. Se determinó una relación directa, positiva y en grado considerable (rs = 0,743; p = 0,000) entre la comunicación interna y el "
    "desempeño laboral de los profesionales y técnicos de Electro Puno S.A.A. durante el ciclo 2022-2024. Se concluye que un sistema comunicacional "
    "institucional bien articulado, ágil y transparente actúa como un catalizador decisivo en el cumplimiento oportuno y eficiente de las metas corporativas."
)

add_p(
    "2. Se constató una asociación positiva muy fuerte y altamente significativa (rs = 0,840; p = 0,000) entre la provisión idónea de canales "
    "y medios comunicativos internos y la percepción favorable del clima laboral. La diversificación y modernización de plataformas oficiales "
    "mitiga el aislamiento informativo de los servicios eléctricos provinciales y cohesiona la identidad organizacional."
)

add_p(
    "3. Se evidenció una correlación positiva moderada-considerable (rs = 0,567; p = 0,000) entre los atributos cualitativos del flujo informativo "
    "(claridad, inmediatez y retroalimentación bidireccional) y los índices de cohesión y satisfacción de los trabajadores. La transparencia y "
    "certeza procedimental reducen fricciones interdepartamentales y promueven la colaboración interdisciplinaria en el terreno operativo."
)

add_p(
    "4. Se comprobó que el clima organizacional ejerce un rol modulador y mediador preponderante, reflejado en una correlación positiva muy "
    "fuerte (rs = 0,806; p = 0,000) con el desempeño laboral individual y colectivo. Un ambiente de trabajo fundamentado en la equidad, el respeto "
    "mutuo y la supervisión formativa maximiza el talento técnico y eleva los estándares de calidad en el suministro de energía eléctrica regional."
)

# ==============================================================================
# AGRADECIMIENTOS
# ==============================================================================
add_h1("AGRADECIMIENTOS")
add_p(
    "A la Empresa Regional de Servicios Públicos de Electricidad de Puno S.A.A. (Electro Puno S.A.A.), a sus autoridades y muy especialmente "
    "al equipo de profesionales y técnicos de las distintas sedes y servicios eléctricos del departamento de Puno, por su valiosa apertura, "
    "tiempo y compromiso en la administración de los instrumentos psicométricos que permitieron la concreción de este trabajo de investigación."
)

# ==============================================================================
# CONFLICTO DE INTERÉS
# ==============================================================================
add_h1("CONFLICTO DE INTERÉS")
add_p(
    "La autora declara formalmente que no existe ningún conflicto de interés económico, institucional ni personal que haya condicionado "
    "el diseño, análisis o interpretación de los resultados expuestos en el presente artículo científico (C.K.S.Q.)."
)

# ==============================================================================
# REFERENCIAS (Normas APA 7ma Edición con DOIs y Enlaces)
# ==============================================================================
add_h1("REFERENCIAS")

def add_ref_item(text, url="", doi=""):
    p = add_p()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_run_font(r, size_pt=10.5)
    if doi:
        r_doi = p.add_run(f" https://doi.org/{doi}")
        set_run_font(r_doi, size_pt=10.5, color_rgb=RGBColor(0, 0, 255))
    elif url:
        r_url = p.add_run(f" {url}")
        set_run_font(r_url, size_pt=10.5, color_rgb=RGBColor(0, 0, 255))

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
print(f"Complete advanced article generated successfully at: {article_path}")
