#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Caso 86 — Catálogo de Esquemas Tributarios SII 2025
Informe escrito — Análisis tributario y defensa de posición
Magíster en Dirección Tributaria UVM
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT = "/home/user/Prueba/Caso86_Informe.docx"

# ── Colores institucionales ───────────────────────────────────────────────────
AZUL    = RGBColor(0x1A, 0x3A, 0x5C)
CELESTE = RGBColor(0x2E, 0x6D, 0xA4)
ROJO    = RGBColor(0xB0, 0x30, 0x30)
GRIS    = RGBColor(0x55, 0x55, 0x55)
NEGRO   = RGBColor(0x1C, 0x1C, 0x1C)

doc = Document()

# Configuración de página
for section in doc.sections:
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# Configurar estilo base
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.line_spacing = 1.15

def set_cell_bg(cell, color_hex):
    """Aplica color de fondo a una celda."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def add_title(text, size=16, color=AZUL, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = doc.add_paragraph()
    p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    return p

def add_subtitle(text, size=11, color=CELESTE):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    return p

def add_section(text, size=12, color=AZUL):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '2E6DA4')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

def add_para(text, bold=False, justify=True, color=None, size=10.5):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return p

def add_bullet(text, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.5 + 0.5*level)
    p.paragraph_format.space_after = Pt(2)
    run = p.runs[0] if p.runs else p.add_run("")
    run.text = text
    run.font.size = Pt(10)
    return p

def add_quote(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(9.5)
    run.font.italic = True
    run.font.color.rgb = GRIS
    return p

# ═══════════════════════════════════════════════════════════════════════════════
# PORTADA / ENCABEZADO
# ═══════════════════════════════════════════════════════════════════════════════

# Encabezado institucional
encabezado = doc.add_paragraph()
encabezado.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = encabezado.add_run("MAGÍSTER EN DIRECCIÓN TRIBUTARIA — UNIVERSIDAD VIÑA DEL MAR")
run.font.size = Pt(9)
run.font.bold = True
run.font.color.rgb = GRIS

add_title("CASO 86 — CATÁLOGO DE ESQUEMAS TRIBUTARIOS SII 2025", size=14, color=AZUL)
add_title("Prestación de servicios profesionales mediante interposición de sociedad de profesionales",
          size=11, color=CELESTE)
add_title("Informe de Análisis y Defensa de Posición", size=10, color=GRIS)

# Línea divisoria
linea = doc.add_paragraph()
linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
linea_run = linea.add_run("_" * 80)
linea_run.font.size = Pt(8)
linea_run.font.color.rgb = CELESTE

# ═══════════════════════════════════════════════════════════════════════════════
# I. DESCRIPCIÓN DEL ESQUEMA
# ═══════════════════════════════════════════════════════════════════════════════
add_section("I.  DESCRIPCIÓN DEL ESQUEMA OBJETO DE ANÁLISIS")

add_para(
    "El contribuyente A, persona natural, constituye junto a otros dos socios (B y C) una "
    "sociedad de profesionales D, destinada formalmente a la prestación de servicios. "
    "La participación social se distribuye en partes iguales (33,33% cada uno), tanto en "
    "capital como en utilidades. El capital social aportado no resulta significativo en "
    "comparación con los ingresos anuales generados por la compañía."
)

add_para(
    "El elemento crítico del esquema es que los servicios profesionales son prestados "
    "casi exclusivamente por el socio A a diversas instituciones y personas naturales, "
    "mientras que B y C no prestan servicios profesionales para la sociedad ni "
    "desempeñan algún rol estratégico dentro de ella. Sin embargo, al efectuarse los "
    "retiros de utilidades en proporción al porcentaje de capital aportado (33,33% cada uno), "
    "los tres socios reciben en partes iguales las utilidades generadas íntegramente por A."
)

add_para(
    "Como resultado, los ingresos generados por la actividad profesional de A son "
    "diluidos en tres tramos de Impuesto Global Complementario (IGC) independientes, "
    "reduciendo significativamente la carga tributaria del contribuyente A en relación "
    "a la que correspondería si dichos ingresos los percibiera directamente (sea como "
    "trabajador independiente del Art. 42 N°2 LIR, o como único socio prestador de servicios)."
)

# ═══════════════════════════════════════════════════════════════════════════════
# II. NORMATIVA APLICABLE
# ═══════════════════════════════════════════════════════════════════════════════
add_section("II.  NORMATIVA APLICABLE")

add_subtitle("1. Artículo 42 N°2 de la Ley sobre Impuesto a la Renta")
add_para(
    "El Art. 42 N°2 LIR califica como rentas de segunda categoría los ingresos provenientes "
    "del ejercicio de profesiones liberales u otras ocupaciones lucrativas no comprendidas "
    "en la primera categoría. El inciso tercero de la norma permite a las sociedades de "
    "profesionales que prestan exclusivamente servicios o asesorías profesionales optar "
    "por declarar sus rentas conforme a las normas de la primera categoría, sujetándose "
    "a sus disposiciones para todos los efectos de la ley."
)

add_subtitle("2. Circular N°21 de 1991 y Circular N°50 de 2020 del SII")
add_para(
    "Ambas circulares fijan los requisitos copulativos que deben cumplirse para calificar "
    "como sociedad de profesionales:"
)
add_bullet("Debe tratarse de una sociedad de personas.")
add_bullet("Su objeto exclusivo debe ser la prestación de servicios o asesorías profesionales.")
add_bullet("Los servicios deben ser prestados por intermedio de sus socios, asociados o "
           "con la colaboración de dependientes que coadyuven a la prestación del servicio.")
add_bullet("TODOS sus socios (sean personas naturales u otras sociedades de profesionales) "
           "deben ejercer sus profesiones para la sociedad. No es aceptable que uno o más "
           "de ellos solo aporte capital.")
add_bullet("No es relevante la forma de remuneración a los socios (contrato de trabajo, "
           "sueldo patronal, retiro de utilidades), sino la realización efectiva de "
           "labores profesionales en beneficio de la sociedad.")

add_subtitle("3. Normas Generales Antielusivas — Arts. 4 bis, 4 ter, 4 quáter y 4 quinquies del CT")
add_para(
    "Incorporadas por la Ley N°20.780 (2014) y modificadas por la Ley N°21.210 (2020), "
    "establecen el marco para que la administración tributaria pueda recalificar actos "
    "o contratos cuando se configura abuso de las formas jurídicas o simulación:"
)
add_bullet(
    "Art. 4 bis CT: Principio de buena fe. Las obligaciones tributarias nacen y se exigen "
    "con arreglo a la naturaleza jurídica de los hechos, actos o negocios realizados, "
    "cualquiera sea la forma o denominación que los interesados les hubieren dado."
)
add_bullet(
    "Art. 4 ter CT (Abuso): Existe abuso cuando se evita total o parcialmente la realización "
    "del hecho gravado, o se disminuye la base imponible o la obligación tributaria, "
    "mediante actos o negocios que, individualmente considerados o en su conjunto, no "
    "produzcan resultados o efectos jurídicos o económicos relevantes para el contribuyente "
    "o un tercero, que sean distintos de los meramente tributarios."
)
add_bullet(
    "Art. 4 quáter CT (Simulación): Habrá simulación cuando los actos o negocios disimulen "
    "la configuración del hecho gravado o la naturaleza de los elementos constitutivos "
    "de la obligación tributaria, o su verdadero monto o data de nacimiento."
)
add_bullet(
    "Art. 4 quinquies CT: Procedimiento. La existencia de abuso o simulación será "
    "declarada por el Tribunal Tributario y Aduanero, a requerimiento del Director del SII."
)

add_subtitle("4. Circular N°65 de 2015 del SII")
add_para(
    "Imparte instrucciones sobre la aplicación práctica de las Normas Generales Antielusivas. "
    "Precisa que el abuso requiere acreditar la ausencia de efectos jurídicos o económicos "
    "relevantes distintos de los tributarios, y que la simulación supone una divergencia "
    "consciente entre la voluntad real y la declarada."
)

add_subtitle("5. Oficio N°2359 de 2022 del SII")
add_para(
    "Pronunciamiento que aclara que la sociedad de profesionales que optó por tributar "
    "en primera categoría no puede transitar nuevamente a segunda categoría, reforzando "
    "el carácter estable de la opción del Art. 42 N°2 inciso 3°."
)

# ═══════════════════════════════════════════════════════════════════════════════
# III. ANÁLISIS JURÍDICO
# ═══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_section("III.  ANÁLISIS JURÍDICO DEL CASO")

add_subtitle("1. Calificación de la sociedad D como 'sociedad de profesionales'")

add_para(
    "El primer punto a dilucidar es si la sociedad D efectivamente califica como "
    "'sociedad de profesionales' en los términos del Art. 42 N°2 inciso 3° LIR y "
    "de la Circular N°21 de 1991. La respuesta es negativa, por la siguiente razón:"
)
add_para(
    "Uno de los requisitos copulativos exigidos por la administración tributaria es que "
    "TODOS los socios ejerzan sus profesiones para la sociedad. En el caso analizado, "
    "solo el socio A presta efectivamente los servicios profesionales; B y C no prestan "
    "servicios ni tienen rol estratégico en la entidad. En consecuencia, D no reúne los "
    "requisitos formales para ser considerada sociedad de profesionales y, por tanto, "
    "tampoco puede acogerse válidamente a la opción de tributar en primera categoría "
    "que contempla la norma especial.", bold=False
)

add_subtitle("2. Análisis bajo el principio de realidad económica (Art. 4 bis CT)")
add_para(
    "Conforme al Art. 4 bis del Código Tributario, las obligaciones tributarias nacen y "
    "se exigen con arreglo a la naturaleza jurídica de los hechos, cualquiera sea la "
    "forma o denominación. En el caso, la realidad económica es que A es el único "
    "profesional que ejerce labor lucrativa, generando los ingresos que se atribuyen "
    "indiscriminadamente a los tres socios."
)

add_subtitle("3. ¿Abuso o Simulación?")

add_para(
    "Corresponde determinar si la conducta descrita configura un caso de abuso de las "
    "formas jurídicas (Art. 4 ter CT) o de simulación (Art. 4 quáter CT). El análisis "
    "del grupo concluye que se está frente a un caso de ABUSO, no de simulación, por "
    "las siguientes razones:"
)

add_bullet(
    "La sociedad D existe formalmente y está válidamente constituida. Los actos "
    "societarios son reales: hay un contrato de sociedad, un capital social aportado, "
    "una distribución de utilidades efectiva. No hay un acto disimulado tras otro: "
    "lo que se observa es exactamente lo que se declara (una sociedad con tres socios "
    "que reparten utilidades en proporción al capital)."
)
add_bullet(
    "El abuso radica en la utilización de una FORMA JURÍDICA inapropiada para el sustrato "
    "económico que se quiere amparar. La sociedad de profesionales es una figura prevista "
    "para que profesionales que trabajan en conjunto puedan organizarse con personalidad "
    "jurídica común. No está prevista para canalizar el trabajo individual de uno solo, "
    "diluyendo artificialmente sus utilidades en otros socios que no aportan trabajo."
)
add_bullet(
    "Los actos del esquema no producen efectos jurídicos o económicos relevantes "
    "distintos del meramente tributario: la existencia de B y C como socios no agrega "
    "valor económico al negocio (no aportan capital significativo, no prestan servicios, "
    "no aportan clientela, no tienen rol estratégico). Su única función es servir de "
    "vehículo para 'absorber' parte de las utilidades de A."
)

add_subtitle("4. Efecto tributario del esquema — perjuicio fiscal")

add_para(
    "El esquema afecta indebidamente la progresión del IGC del socio A. A modo "
    "ilustrativo, supóngase que la sociedad genera utilidades anuales por $300 millones, "
    "íntegramente atribuibles al trabajo profesional de A:"
)

# Tabla comparativa
tabla = doc.add_table(rows=4, cols=3)
tabla.style = 'Light Grid Accent 1'
tabla.alignment = WD_ALIGN_PARAGRAPH.CENTER

hdr = tabla.rows[0].cells
hdr[0].text = "Escenario"
hdr[1].text = "Atribución de renta"
hdr[2].text = "Impacto IGC aproximado"
for c in hdr:
    set_cell_bg(c, "1A3A5C")
    for p in c.paragraphs:
        for r in p.runs:
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            r.font.bold = True
            r.font.size = Pt(9.5)

filas_data = [
    ["Sin esquema (A solo)", "A recibe $300 MM",
     "IGC en tramo marginal alto (40%) sobre un único contribuyente"],
    ["Con esquema D (3 socios)", "A: $100 MM · B: $100 MM · C: $100 MM",
     "Cada socio queda en tramo más bajo: ahorro fiscal significativo"],
    ["Diferencial", "Mismo ingreso económico real",
     "Reducción artificial de carga tributaria — ABUSO"],
]
for i, fila in enumerate(filas_data):
    row = tabla.rows[i+1].cells
    for j, val in enumerate(fila):
        row[j].text = val
        for p in row[j].paragraphs:
            for r in p.runs:
                r.font.size = Pt(9)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════════════════════════
# IV. POSICIÓN DEL GRUPO
# ═══════════════════════════════════════════════════════════════════════════════
add_section("IV.  POSICIÓN DEL GRUPO")

add_para(
    "Considerando el análisis jurídico precedente, el grupo sostiene la siguiente posición:",
    bold=True
)

add_para(
    "El esquema descrito en el Caso 86 configura un ABUSO de las formas jurídicas en los "
    "términos del Art. 4 ter del Código Tributario. La constitución de la sociedad de "
    "profesionales D, sin que B y C ejerzan efectivamente labor profesional para ella, "
    "constituye una utilización inapropiada de una figura societaria especial, cuyo único "
    "efecto relevante es la reducción artificial de la carga tributaria de A en el IGC.", bold=True
)

add_para("Fundamentos centrales de la posición:")
add_bullet(
    "Incumplimiento de requisitos formales: D no califica como sociedad de profesionales "
    "según Circular N°21 de 1991 y N°50 de 2020 (no todos los socios ejercen su profesión "
    "para la sociedad)."
)
add_bullet(
    "Inexistencia de motivo económico válido: la participación de B y C no tiene justificación "
    "económica fuera de la dilución del IGC de A. El capital aportado es insignificante "
    "frente a los ingresos."
)
add_bullet(
    "Configuración del Art. 4 ter CT: se evita parcialmente el hecho gravado del IGC en su "
    "tramo marginal más alto, mediante un acto que no produce efectos jurídicos ni económicos "
    "relevantes distintos del meramente tributario."
)
add_bullet(
    "Descarte de simulación (Art. 4 quáter): no existe divergencia entre voluntad real y "
    "declarada. Los actos son reales pero la forma es abusiva, no simulada."
)

# ═══════════════════════════════════════════════════════════════════════════════
# V. ALTERNATIVAS DE SOLUCIÓN
# ═══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_section("V.  ALTERNATIVAS DE SOLUCIÓN AL CASO")

add_para(
    "El grupo propone las siguientes alternativas legítimas de estructuración tributaria "
    "que permitirían al contribuyente A organizar su actividad profesional sin incurrir "
    "en abuso de formas jurídicas:"
)

add_subtitle("Alternativa 1: Tributación directa como profesional independiente (Art. 42 N°2 LIR)")
add_para(
    "El socio A puede ejercer su actividad profesional como persona natural independiente, "
    "emitiendo boletas de honorarios y declarando sus ingresos en segunda categoría. "
    "Tendrá derecho a deducir gastos efectivos o a optar por gastos presuntos (30% con "
    "tope de 15 UTA). Pagará IGC sobre la totalidad de los ingresos según tabla progresiva. "
    "Esta opción es la más simple y libre de cuestionamientos."
)

add_subtitle("Alternativa 2: Constitución de sociedad de profesionales con socios que efectivamente trabajen")
add_para(
    "Si A tiene voluntad de asociarse con otros profesionales, debe hacerlo con personas "
    "que efectivamente ejerzan su profesión para la sociedad. Cada socio debe coadyuvar "
    "a la prestación del servicio profesional. En tal caso, D calificaría legítimamente "
    "como sociedad de profesionales, podrá optar por tributar en primera o segunda "
    "categoría, y la distribución de utilidades reflejará el trabajo aportado por cada uno."
)

add_subtitle("Alternativa 3: Distribución de utilidades proporcional al trabajo realmente prestado")
add_para(
    "Si la sociedad se mantiene con A, B y C como socios, los estatutos pueden contemplar "
    "una distribución diferenciada de utilidades que reconozca la mayor contribución de A. "
    "Esto puede instrumentarse mediante: (i) participación social diferenciada (ej.: A 90%, "
    "B y C 5% cada uno) o (ii) sueldo patronal a A por el trabajo prestado, antes de "
    "repartir utilidades. La distribución debe reflejar la realidad económica del aporte."
)

add_subtitle("Alternativa 4: Sociedad por acciones (SpA) con régimen Pro Pyme 14D N°3")
add_para(
    "Constituir una SpA acogida al régimen 14D N°3 (Pro Pyme General) con tasa IDPC del 25%, "
    "siempre que A sea el único accionista o que los demás accionistas tengan rol "
    "efectivamente económico. La SpA no tiene la restricción de la sociedad de profesionales "
    "pero exige sustancia económica real. Permite planificación de retiros y reinversión."
)

add_subtitle("Alternativa 5: Régimen Pro Pyme Transparente 14D N°8")
add_para(
    "Si los socios son únicamente personas naturales, evaluar el régimen 14D N°8 "
    "(Transparencia Tributaria). La empresa no paga IDPC; los socios tributan directamente "
    "con IGC sobre su participación. Sin embargo, ATENCIÓN: este régimen no soluciona el "
    "problema de fondo si los socios B y C siguen siendo ficticios — el SII puede igualmente "
    "aplicar NGA. Solo es válida si B y C efectivamente aportan trabajo o sustancia económica."
)

# ═══════════════════════════════════════════════════════════════════════════════
# VI. CONCLUSIONES
# ═══════════════════════════════════════════════════════════════════════════════
add_section("VI.  CONCLUSIONES")

add_para(
    "1. El esquema descrito en el Caso 86 del Catálogo SII 2025 corresponde a un ABUSO "
    "de las formas jurídicas conforme al Art. 4 ter del Código Tributario. NO se trata "
    "de simulación, ya que los actos jurídicos son reales aunque inapropiados."
)
add_para(
    "2. La sociedad D no califica como sociedad de profesionales conforme a las Circulares "
    "N°21 de 1991 y N°50 de 2020, dado que B y C no ejercen sus profesiones para ella. "
    "Esta sola circunstancia ya genera un problema formal que cuestiona la opción de "
    "tributar en primera categoría."
)
add_para(
    "3. El SII, a requerimiento del Director, puede solicitar al Tribunal Tributario y "
    "Aduanero la declaración de abuso conforme al Art. 4 quinquies CT, con el efecto de "
    "recalificar el esquema y atribuir la totalidad de los ingresos a A para el cálculo "
    "del IGC, con los correspondientes intereses y multas."
)
add_para(
    "4. Existen alternativas legítimas de organización tributaria que el contribuyente A "
    "puede adoptar, según se ha expuesto en la sección V. La elección dependerá de la "
    "estrategia de negocio, las proyecciones de ingresos, la composición real del equipo "
    "profesional y los objetivos de planificación de retiros."
)
add_para(
    "5. Recomendación final: en una eventual fiscalización, el contribuyente debe estar "
    "preparado para acreditar la sustancia económica real del esquema. La sola existencia "
    "formal de la sociedad y de los socios B y C no será suficiente; el SII y los tribunales "
    "exigirán demostración del aporte efectivo de cada socio.", bold=True
)

# Línea final
doc.add_paragraph()
linea2 = doc.add_paragraph()
linea2.alignment = WD_ALIGN_PARAGRAPH.CENTER
linea2_run = linea2.add_run("_" * 80)
linea2_run.font.size = Pt(8)
linea2_run.font.color.rgb = CELESTE

# Pie de página
pie = doc.add_paragraph()
pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
pie_run = pie.add_run(
    "Magíster en Dirección Tributaria — UVM · Caso 86 · Catálogo SII 2025\n"
    "Fuentes: DL 824 (LIR Art. 42 N°2); Código Tributario Arts. 4 bis, 4 ter, 4 quáter, 4 quinquies; "
    "Circular N°21 de 1991; Circular N°50 de 2020; Circular N°65 de 2015; Oficio N°2359 de 2022 SII."
)
pie_run.font.size = Pt(8)
pie_run.font.italic = True
pie_run.font.color.rgb = GRIS

doc.save(OUTPUT)
print(f"✔ Informe Word generado: {OUTPUT}")
