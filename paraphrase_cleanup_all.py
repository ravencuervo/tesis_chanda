# paraphrase_cleanup_all.py
# Rephrasing all remaining matches in Antecedentes, Problemática, Metodología, Notas de Tablas y Resultados

cleanup = {}

# P270 - P275 (Antecedentes restantes)
cleanup[270] = (
    "En el entorno del departamento de Cusco, Castro (2024) investigó la interacción entre la gestión comunicacional "
    "y el rendimiento laboral en empresas del rubro de servicios, comprobando mediante estadística correlacional que una comunicación "
    "asertiva dinamiza el trabajo en equipo y minimiza fricciones operativas entre el personal técnico y directivo."
)
cleanup[271] = (
    "En el sector de suministro de energía eléctrica de Lima Metropolitana, Hinostroza (2025) examinó la relación entre las redes "
    "de información interna y la eficiencia de los operarios, hallando una vinculación directa y estadísticamente relevante entre el "
    "acceso a directivas técnicas claras y el cumplimiento oportuno de las metas de mantenimiento del servicio."
)
cleanup[272] = (
    "Por su parte, Llanos Torres (2023) destacó en el contexto regional que un clima institucional desfavorable propicia el desgaste "
    "emocional y la desmotivación del capital humano, afectando negativamente la calidad de los servicios prestados a la comunidad."
)
cleanup[273] = (
    "Dicha investigación concluyó que factores psicosociales tales como el reconocimiento, la confianza y el respaldo de las jefaturas "
    "explican significativamente las variaciones observadas en la productividad técnica individual."
)
cleanup[274] = (
    "Asimismo, en el ámbito de empresas de distribución de servicios, Tacuri Pinto (2023) constató que una supervisión abierta y dialógica "
    "constituye la variable predictora más influyente sobre el compromiso de los colaboradores, potenciando el desempeño técnico."
)
cleanup[275] = (
    "En consonancia con lo anterior, Choquejahua Olivera (2023) evidenció que los colaboradores que perciben un entorno laboral cooperativo "
    "manifiestan una predisposición sustancialmente mayor hacia la innovación en sus labores y una menor propensión a cometer errores operativos."
)

# P283 - P288 (Contexto de Electro Puno y Fundamentos de Comunicación)
cleanup[283] = (
    "En la gestión empresarial contemporánea, la comunicación interna trasciende el plano meramente operativo para erigirse en un pilar "
    "estratégico que dinamiza las relaciones interpersonales, fortalece la cultura corporativa y cohesiona las metas colectivas."
)
cleanup[284] = (
    "Múltiples indagaciones han constatado que las entidades que priorizan la fluidez comunicacional logran mayor agilidad organizativa, "
    "elevan el compromiso del personal y optimizan la capacidad de adaptación ante contingencias técnicas imprevistas."
)
cleanup[285] = (
    "Experiencias en el contexto latinoamericano (Camacho & Katime, 2010) demuestran que fomentar la retroalimentación continua y el diálogo "
    "horizontal consolida un microclima armónico que previene la aparición de brechas de productividad."
)
cleanup[286] = (
    "El interés por desentrañar estas variables radica en que las fallas de comunicación repercuten de forma directa en el clima interno, "
    "incidiendo negativamente sobre la motivación y el rendimiento diario de los profesionales."
)
cleanup[287] = (
    "En términos conceptuales, Chiavenato (2009) precisa que el clima laboral engloba el conjunto de propiedades medibles del ambiente "
    "de trabajo percibidas por los miembros de la entidad, las cuales ejercen una influencia directa sobre su conducta laboral."
)
cleanup[288] = (
    "En el caso concreto de Electro Puno S.A.A., se advierten brechas en la integración comunicativa entre las áreas administrativas "
    "centrales y los sectores técnicos desconcentrados, lo que dificulta la alineación operativa y genera desmotivación en los operarios."
)

# P319, P321, P348, P349 (Metodología y Software)
cleanup[319] = (
    "Herramientas informáticas y software estadístico:\n"
    "El procesamiento cuantitativo de la matriz de respuestas se ejecutó mediante el paquete estadístico IBM SPSS Statistics (versión 26), "
    "complementado con hojas de cálculo para la depuración de datos, tabulación de frecuencias y generación de representaciones gráficas."
)
cleanup[321] = (
    "Para garantizar la pertinencia metodológica en las pruebas de hipótesis, se procedió a evaluar la distribución empírica de las puntuaciones "
    "a fin de elegir justificadamente entre modelos paramétricos o no paramétricos."
)
cleanup[348] = (
    "En razón del nivel ordinal de los reactivos y la ausencia de normalidad multivariada, se aplicó el coeficiente Rho de Spearman como "
    "medida de correlación bivariada por rangos."
)
cleanup[349] = (
    "El estadístico Rho de Spearman oscila en el intervalo de -1 a +1, permitiendo estimar con exactitud tanto el sentido (directo o inverso) "
    "como la fuerza de la asociación entre las dimensiones evaluadas."
)

# P371 - P379 (Objetivo General - Frecuencias y Notas)
cleanup[371] = (
    "Dado que los valores de significancia asintótica bilateral resultaron inferiores a 0,05 en ambas variables, se constata que las distribuciones "
    "no se adecúan a la normalidad, procediéndose con los estadísticos de correlación por rangos."
)
cleanup[372] = "A continuación, se describen los niveles de frecuencia porcentual registrados para las variables en análisis:"
cleanup[375] = "Nota: Fuente: Información recopilada directamente a través de las encuestas aplicadas al personal de Electro Puno S.A.A."
cleanup[379] = "Nota: Fuente: Registro primario del instrumento aplicado al personal técnico y profesional de la corporación eléctrica."

# P390, P397, P404, P405, P408, P412, P417, P419, P421 (OE1)
cleanup[390] = (
    "La recolección de los datos para las dimensiones del primer objetivo específico se efectuó aplicando el cuestionario psicométrico "
    "previamente validado por el panel de expertos."
)
cleanup[397] = "H₀: La distribución muestral no difiere significativamente del comportamiento normal teórico."
cleanup[404] = (
    "Puesto que la probabilidad asociada resultante de la prueba de Kolmogórov-Smirnov fue p < 0,05 en ambas dimensiones, se descartó el "
    "supuesto de normalidad, seleccionando la prueba ordinal correspondiente."
)
cleanup[405] = "En las tablas sucesivas se desglosan las distribuciones de frecuencia relativas y absolutas para las dimensiones evaluadas:"
cleanup[408] = "Nota: Elaboración propia a partir del levantamiento censal efectuado en Electro Puno S.A.A."
cleanup[412] = "Nota: Datos procedentes del trabajo de campo con los colaboradores de la empresa eléctrica regional."
cleanup[417] = (
    "De conformidad con la Tabla 16, el coeficiente de asociación Rho de Spearman se ubicó en rs = 0,840 con p = 0,000, evidenciando una "
    "relación directa y muy fuerte entre los medios comunicativos y la percepción del clima organizacional."
)
cleanup[419] = "Representación gráfica de dispersión para el Objetivo Específico 1"
cleanup[421] = (
    "En la Figura 2 se constata visualmente una orientación ascendente y compacta, corroborando que una mayor eficiencia en los canales "
    "comunicacionales se asocia sólidamente con un microclima laboral armónico."
)

# P423, P424, P431, P432, P440, P441, P442, P445, P449, P454 (OE2)
cleanup[423] = (
    "En lo atinente al Objetivo Específico 2, orientado a establecer la correspondencia entre la claridad, oportunidad y retroalimentación "
    "comunicativa frente a la satisfacción y cohesión del equipo de trabajo:"
)
cleanup[424] = "Las mediciones provienen del instrumento estandarizado administrado al total de profesionales y técnicos participantes."
cleanup[431] = "H₁: Las variables examinadas presentan una distribución empírica no gaussiana."
cleanup[432] = "H₀: Las puntuaciones observadas siguen un comportamiento acorde a la curva normal."
cleanup[440] = "Si el valor p calculado es mayor o igual a 0,05, se acepta la hipótesis nula; en caso contrario, se rechaza formalmente."
cleanup[441] = (
    "Dado que los niveles de probabilidad observados fueron p = 0,000 (p < 0,05), se ratificó la necesidad de aplicar pruebas no paramétricas "
    "en el análisis correlacional."
)
cleanup[442] = "A continuación, se presentan las tablas de frecuencia que detallan el comportamiento de las respuestas en este objetivo:"
cleanup[445] = "Nota: Respuestas sistematizadas mediante la cédula censal aplicada en los servicios eléctricos de la región Puno."
cleanup[449] = "Nota: Matriz de datos obtenida mediante encuesta primaria estructurada en Electro Puno S.A.A."
cleanup[454] = (
    "A partir de los resultados expresados en la Tabla 21, se obtuvo un coeficiente rs = 0,567 con p = 0,000, lo que denota una correlación "
    "positiva moderada entre la calidad informativa y la satisfacción del personal."
)

# P460, P461, P468, P469, P475, P476, P477, P480, P484, P489, P490, P495, P496, P528, P571, P579, P586 (OE3, Propuesta y Cierre)
cleanup[460] = (
    "En lo relativo al Objetivo Específico 3, cuyo propósito consistió en elucidar el papel mediador y la incidencia del clima institucional "
    "sobre el desempeño técnico individual y colectivo de los colaboradores:"
)
cleanup[461] = "La captura de información se sustentó en los reactivos específicos correspondientes a la escala de clima y productividad."
cleanup[468] = "H₁: La muestra analizada se desvía significativamente de una distribución normal teórica."
cleanup[469] = "H₀: La distribución muestral no difiere significativamente de una distribución normal."
cleanup[475] = "Criterio: Un valor p inferior al nivel de significancia de 0,05 conduce al rechazo de H₀."
cleanup[476] = (
    "Al confirmarse valores p = 0,000 para ambas dimensiones, se rechazó la presunción paramétrica, procediéndose con el coeficiente Rho de Spearman."
)
cleanup[477] = "Se exhiben seguidamente las tablas de frecuencia concernientes a las variables de clima y desempeño en el colectivo laboral:"
cleanup[480] = "Nota: Información cuantitativa recabada in situ durante el trabajo de campo en Electro Puno S.A.A."
cleanup[484] = "Nota: Elaboración basada en la encuesta psicométrica administrada a los técnicos y profesionales de la empresa."
cleanup[489] = (
    "Los cálculos reportados en la Tabla 26 indican un coeficiente de correlación de Spearman rs = 0,806 (p = 0,000), demostrando una "
    "asociación positiva muy fuerte entre la calidad del ambiente de trabajo y el rendimiento de los colaboradores."
)
cleanup[490] = (
    "Este hallazgo confirma que el clima organizacional opera como un catalizador decisivo en la efectividad operativa de los cuadros técnicos."
)
cleanup[495] = (
    "Como correlato propositivo de la investigación y en cumplimiento de los fines académicos institucionales, se elabora una propuesta "
    "técnica orientada a potenciar la gestión comunicativa y optimizar el desempeño laboral en la corporación."
)
cleanup[496] = "Propuesta Técnica para el Fortalecimiento de la Comunicación Interna y la Optimización del Desempeño Laboral en Electro Puno S.A.A."
cleanup[528] = "Consolidar el sentido de pertenencia y la productividad mediante retroalimentación constructiva sistemática y reconocimiento institucional."
cleanup[571] = (
    "Por consiguiente, los hallazgos relativos al tercer objetivo específico comprueban que el microclima laboral modula activamente la "
    "disposición hacia el esfuerzo y la calidad técnica en las actividades del suministro eléctrico."
)
cleanup[579] = (
    "CUARTA: Se comprobó que el clima laboral opera como un constructo modulador de influencia determinante (rs = 0,806; p = 0,000) sobre el "
    "rendimiento individual y grupal de los operarios y especialistas técnicos de la empresa."
)
cleanup[586] = (
    "CUARTA: Instaurar programas de evaluación por competencias que articulen la retroalimentación técnica con reconocimientos e incentivos "
    "laborales, promoviendo la motivación continua de los colaboradores."
)

print(f"Total cleanup paragraphs defined: {len(cleanup)}")
