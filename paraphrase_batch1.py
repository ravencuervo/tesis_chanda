import docx
import re

# Load the thesis document
doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

# Dictionary of paragraph replacements: index -> new_text
replacements = {}

# ==============================================================================
# 1. RESUMEN Y ABSTRACT (P139, P142)
# ==============================================================================
replacements[139] = (
    "El presente trabajo investigativo examinó el nexo existente entre la dinámica de la comunicación interna, "
    "el clima organizacional y el desempeño laboral exhibido por el personal técnico y profesional vinculado a "
    "Electro Puno S.A.A. durante el ciclo 2022-2024. Se planteó como interrogante central elucidar cómo incide "
    "la comunicación intraorganizacional sobre el entorno de trabajo y la productividad laboral en dicha corporación, "
    "desglosándose en metas orientadas a valorar los canales de difusión, la calidad y fluidez informativa, "
    "así como el rol modulador del ambiente institucional. La metodología adoptó una ruta cuantitativa de alcance "
    "descriptivo-correlacional, sustentada en un diseño no experimental y de corte transversal, administrando un cuestionario "
    "validado por juicio pericial a una muestra censal de 158 trabajadores. El análisis estadístico evidenció un "
    "coeficiente Rho de Spearman igual a 0,743 con un nivel de significancia bilateral p = 0,000, demostrando una "
    "asociación positiva considerable entre la comunicación corporativa y el rendimiento del capital humano. Se infiere "
    "que optimizar los esquemas comunicativos internos propicia un microclima laboral armónico y cataliza la consecución de las metas operativas."
)

replacements[142] = (
    "This research investigated the relationship between internal communication dynamics, organizational climate, "
    "and job performance among professional and technical staff at Electro Puno S.A.A. during the 2022–2024 period. "
    "The core inquiry centered on determining how internal communications impact the labor environment and workforce "
    "productivity, addressing specific dimensions such as media accessibility, information flow quality, and the mediating "
    "function of organizational climate. Methodologically, a quantitative paradigm with descriptive-correlational scope "
    "and a non-experimental cross-sectional design was executed, administering validated Likert-scale questionnaires to a "
    "census sample of 158 employees. Statistical computation through Spearman's Rho yielded a coefficient of 0.743 "
    "(p = 0.000), proving a substantial positive correlation between internal communication efficiency and employee "
    "performance. In conclusion, solidifying institutional communication channels systematically enhances workplace "
    "morale, fosters professional engagement, and reinforces strategic organizational achievements."
)

# ==============================================================================
# 2. INTRODUCCIÓN - PROBLEMÁTICA GLOBAL, REGIONAL Y LOCAL (P149 to P156, P161)
# ==============================================================================
replacements[149] = (
    "En el contexto latinoamericano, las fallas comunicacionales representan una constante que debilita el rendimiento "
    "laboral, eleva las tasas de rotación y exacerba los episodios de estrés ocupacional producto de la ambigüedad en los mensajes; "
    "de ahí la urgencia de que las instituciones configuren mecanismos de intercambio informativo directos, comprensibles y "
    "homogéneos (Damián et al., 2021). Evaluaciones empíricas efectuadas en más de 500 corporaciones de la región constatan que "
    "apenas un 23% de las entidades acompaña decididamente a sus colaboradores en la formulación de iniciativas y contenidos, un "
    "35% permanece en la inacción y un 25% circunscribe el diálogo exclusivamente a la cúpula directiva, lo cual denota una "
    "marcada centralización (Talent & Empresa, 2023). A pesar de ello, se reconoce que estructurar una adecuada política comunicativa "
    "favorece la articulación estratégica y cohesiona las operaciones del negocio mediante plataformas digitales y canales corporativos formales."
)

replacements[150] = (
    "Al aterrizar en la realidad operativa de Electro Puno S.A.A., emergen manifestaciones concretas de apatía y "
    "desmotivación entre el talento humano, generadas en gran medida por la carencia de retroalimentación oportuna y "
    "la asimetría informativa entre las dependencias técnicas y administrativas. Dicha disonancia repercute de manera "
    "desfavorable sobre la satisfacción laboral y el clima institucional, limitando la integración de los profesionales "
    "en los objetivos estratégicos corporativos."
)

replacements[151] = (
    "A partir de este diagnóstico empírico, se erigió el problema general de indagación: ¿De qué modo repercute la comunicación "
    "interna en la configuración del clima institucional y en el desempeño del colectivo técnico y profesional de Electro Puno S.A.A., "
    "a lo largo del bienio 2022-2024?"
)

replacements[152] = (
    "En el espectro internacional, se examinó la contribución de Jibaja (2020), quien analizó la efectividad de los canales "
    "comunicacionales en el sector salud de Quito, reportando correlaciones directas y estadísticamente significativas entre la "
    "fluidez informativa y los indicadores de eficacia del personal asistencial."
)

replacements[153] = (
    "Asimismo, en el escenario ecuatoriano se contrastaron estudios orientados a cuantificar la relevancia de las redes "
    "informativas internas, comprobándose que los esquemas basados en diálogo multidireccional incrementan notablemente el sentido "
    "de pertenencia y mitigan conflictos interpersonales."
)

replacements[154] = (
    "En el plano nacional, Rodríguez Delgado (2025) comprobó que una adecuada gestión de los canales organizacionales reduce "
    "sustancialmente los índices de insatisfacción y fortalece el compromiso de los colaboradores en corporaciones prestadoras de servicios."
)

replacements[155] = (
    "De igual forma, las evidencias aportadas por Ramírez (2025) corroboran que el intercambio comunicativo formal constituye "
    "un predictor directo de la estabilidad emocional y laboral dentro de las entidades públicas y privadas del país."
)

replacements[156] = (
    "A escala regional, destaca la pesquisa de Berolatti Obando (2023), desarrollada en Puno, donde se concluyó que la percepción "
    "favorable del entorno laboral está estrechamente ligada a la transparencia institucional y a la calidad de la retroalimentación "
    "que proporcionan los mandos directivos."
)

replacements[161] = (
    "Bajo estas premisas, el propósito primordial de la investigación radicó en determinar la incidencia de la comunicación "
    "interna en el clima organizacional y en el rendimiento laboral exhibido por profesionales y técnicos en Electro Puno S.A.A. "
    "durante el periodo 2022 al 2024, generando pautas sólidas para optimizar la gestión del talento en el sector eléctrico regional."
)

print("Batch 1 (Resumen & Introduccion) defined.")
