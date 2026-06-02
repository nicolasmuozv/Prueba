#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Caso 2 Individual - Tributación Sectorial
Nicolás Muñoz Valdebenito - Magíster en Dirección Tributaria UVM
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT = "/home/user/Prueba/Caso2_Tributacion_Sectorial_Nicolas_Munoz.docx"

NEGRO = RGBColor(0x00, 0x00, 0x00)

doc = Document()

# Márgenes
for sec in doc.sections:
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin   = Cm(3.0)
    sec.right_margin  = Cm(2.5)

# ── helpers ───────────────────────────────────────────────────────────────────
def parrafo(text, size=12, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.alignment = align
    r = p.add_run(text)
    r.font.size  = Pt(size)
    r.font.bold  = bold
    r.font.italic = italic
    r.font.color.rgb = NEGRO
    r.font.name  = "Times New Roman"
    return p

def titulo_seccion(text, level=1):
    sizes = {1: 14, 2: 13, 3: 12}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run(text)
    r.font.size  = Pt(sizes.get(level, 12))
    r.font.bold  = True
    r.font.color.rgb = NEGRO
    r.font.name  = "Times New Roman"
    return p

def linea_separadora():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    from docx.oxml import OxmlElement
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)

def parrafo_mixto(partes, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  space_before=0, space_after=6):
    """partes: lista de (texto, bold, italic)"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.alignment = align
    for texto, bold, italic in partes:
        r = p.add_run(texto)
        r.font.size  = Pt(12)
        r.font.bold  = bold
        r.font.italic = italic
        r.font.color.rgb = NEGRO
        r.font.name  = "Times New Roman"
    return p

# ══════════════════════════════════════════════════════════════════
# PORTADA / ENCABEZADO
# ══════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("MAGÍSTER EN DIRECCIÓN TRIBUTARIA")
r.font.size = Pt(13); r.font.bold = True
r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Universidad Viña del Mar")
r.font.size = Pt(12); r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Tributación Sectorial")
r.font.size = Pt(12); r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Profesora: Ximena Niño Roa")
r.font.size = Pt(12); r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("CASO 2 — TRABAJO INDIVIDUAL")
r.font.size = Pt(15); r.font.bold = True
r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Concesiones de Obras de Uso Público — Tratamiento Tributario de Pagos Especiales")
r.font.size = Pt(12)
r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Alumno: Nicolás Muñoz Valdebenito")
r.font.size = Pt(12); r.font.bold = True
r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Enero 2026")
r.font.size = Pt(12); r.font.color.rgb = NEGRO; r.font.name = "Times New Roman"

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════
# INTRODUCCIÓN
# ══════════════════════════════════════════════════════════════════
titulo_seccion("INTRODUCCIÓN", 1)
linea_separadora()

parrafo(
    "Las concesiones de obras de uso público constituyen una modalidad de financiamiento "
    "de infraestructura mediante la cual el Estado encomienda a privados la construcción, "
    "conservación, reparación y explotación de obras públicas, a cambio de otorgarles el "
    "derecho temporal a explotar dichas obras. En Chile, este sistema se rige por la "
    "Ley de Concesiones N°19.460 de 1996, el Decreto MOP N°900 y su Reglamento contenido "
    "en el D.S. N°956 de 1997.", space_before=6
)

parrafo(
    "Desde el punto de vista tributario, la normativa aplicable comprende principalmente "
    "el artículo 15 de la Ley sobre Impuesto a la Renta (LIR), la Circular N°49 de 2006 "
    "del SII para efectos del impuesto a la renta, y los artículos 16 letras c) y h) y 55 "
    "del Decreto Ley N°825 (LIVS), en conjunto con la Circular N°44 de 2006, para efectos "
    "del Impuesto al Valor Agregado (IVA)."
)

parrafo(
    "El Concesionario desarrolla fundamentalmente dos tipos de actividades: (i) la "
    "construcción de una obra civil en terrenos de propiedad fiscal, que constituye un "
    "contrato general de construcción; y (ii) las labores de conservación, reparación "
    "y explotación de la obra, que configuran un contrato de servicios mediante la "
    "utilización de bienes ajenos. Ambas actividades tienen tratamientos tributarios "
    "distintos que la normativa regula en detalle."
)

parrafo(
    "El presente trabajo analiza el tratamiento frente al impuesto a la renta y al IVA "
    "de pagos especiales que surgen en el marco de los contratos de concesión, "
    "específicamente: los premios, las indemnizaciones, el ingreso mínimo garantizado, "
    "y los pagos por concepto de expropiaciones."
)

# ══════════════════════════════════════════════════════════════════
# PREGUNTA 1
# ══════════════════════════════════════════════════════════════════
titulo_seccion("1. TRATAMIENTO TRIBUTARIO DE LOS PAGOS QUE SURGEN EN LAS CONCESIONES DE USO PÚBLICO", 1)
linea_separadora()

parrafo(
    "Como norma general, reconocida en la jurisprudencia administrativa del SII y en la "
    "presentación de la profesora Ximena Niño (Concesiones de Obras de Uso Público, "
    "enero 2026), cualquier pago que el Fisco deba hacer al Concesionario para cumplir "
    "lo convenido en el contrato de concesión, conformará también el precio de la "
    "convención afecta a impuesto. La única excepción a esta regla general corresponde "
    "a los pagos que el Fisco efectúe para solucionar indemnizaciones de perjuicios "
    "pactadas en una cláusula penal, cuyo tratamiento se analiza por separado.", space_before=6
)

# ── 1a ──
titulo_seccion("a) Premios que paga la autoridad concedente al concesionario", 2)

parrafo("Tratamiento frente al Impuesto a la Renta:", bold=True, space_before=6)

parrafo(
    "Las cantidades que el Fisco paga al Concesionario por concepto de premios constituyen "
    "para éste un mayor ingreso bruto afecto a los impuestos generales de la Ley sobre "
    "Impuesto a la Renta. Así lo establece expresamente la Circular N°49 de 2006 del SII "
    "y lo instruye la presentación de la relatora del curso."
)

parrafo(
    "En consecuencia, dichas sumas deben incorporarse a los ingresos brutos del ejercicio "
    "en que sean percibidos o devengados por el Concesionario, quedando afectos al "
    "Impuesto de Primera Categoría (IDPC) conforme a las normas generales del artículo "
    "20 de la LIR. Lo anterior se funda en que los premios implican un incremento "
    "patrimonial que no se encuentra expresamente excluido de renta por el artículo 17 "
    "de la LIR, y que tiene su origen directo en la ejecución del contrato de concesión."
)

parrafo(
    "Cabe destacar que estos ingresos forman parte de los ingresos brutos por servicios "
    "de conservación, reparación y explotación de la obra, cuya base imponible para "
    "efectos del impuesto a la renta se determina conforme al artículo 15 de la LIR y "
    "las instrucciones de la Circular N°49 de 2006, deduciendo del total percibido la "
    "cuota de amortización del costo de construcción imputable al ejercicio respectivo."
)

parrafo("Tratamiento frente al IVA:", bold=True, space_before=4)

parrafo(
    "Desde el punto de vista del IVA, en la medida en que los premios pagados por la "
    "autoridad concedente constituyan una contraprestación asociada a la prestación de "
    "los servicios comprendidos en el contrato de concesión —vale decir, un mayor precio "
    "por los servicios de conservación, reparación y explotación de la obra—, dichos "
    "montos integran la base imponible del IVA conforme a lo dispuesto en los artículos "
    "16 letras c) y h) del DL 825, y las instrucciones de la Circular N°44 de 2006 del SII."
)

parrafo(
    "En este esquema, la base imponible del IVA de los servicios de explotación está "
    "constituida por el total de los ingresos obtenidos, deduciendo la proporción "
    "que deba imputarse al pago del precio de la construcción de la obra —que ya fue "
    "facturado y tributado en la etapa de construcción—. Por lo tanto, si el premio "
    "integra el precio total de los servicios concesionados, el concesionario deberá "
    "emitir la factura correspondiente incluyendo dicho monto en la base imponible "
    "del tributo."
)

# ── 1b ──
titulo_seccion("b) Indemnizaciones que paga la autoridad concedente al concesionario", 2)

parrafo("Tratamiento frente al Impuesto a la Renta:", bold=True, space_before=6)

parrafo(
    "Las indemnizaciones que el Fisco pague al Concesionario —en especial aquellas "
    "derivadas de la resolución o terminación anticipada del contrato de concesión— "
    "constituyen renta para efectos tributarios, en la medida en que impliquen un "
    "incremento patrimonial para su beneficiario. La Circular N°49 de 2006 del SII y "
    "la presentación de la profesora instructora del curso son explícitas al respecto: "
    "\"Las indemnizaciones que el Fisco pague al Concesionario por resolución anticipada "
    "de la concesión, teniendo presente que los Concesionarios son empresas que declaran "
    "la renta efectiva en la Primera Categoría, constituyen renta para los efectos "
    "tributarios.\"", italic=False
)

parrafo(
    "Lo anterior se sustenta en que el artículo 17 de la LIR, que excluye ciertas "
    "indemnizaciones del concepto de renta, no contempla a este tipo de pagos dentro "
    "de las exclusiones taxativas que establece. En consecuencia, estos montos deben "
    "incorporarse a los ingresos brutos del concesionario y quedan afectos al IDPC."
)

parrafo(
    "Sin perjuicio de lo anterior, el Concesionario podrá rebajar como gasto, conforme "
    "al artículo 31 de la LIR, aquella parte de la indemnización que corresponda al "
    "daño emergente efectivamente sufrido en su patrimonio, en la medida en que se "
    "acredite que dicho monto tiene por único objeto restituir una pérdida patrimonial "
    "real y debidamente documentada. El lucro cesante, en cambio, no tiene este "
    "tratamiento y queda gravado íntegramente."
)

parrafo("Tratamiento frente al IVA:", bold=True, space_before=4)

parrafo(
    "Las indemnizaciones que la autoridad concedente pague al Concesionario no se "
    "encuentran afectas a IVA, por cuanto no constituyen una remuneración o precio "
    "por la prestación de un servicio, sino que corresponden a una compensación por "
    "el perjuicio patrimonial sufrido con motivo de la terminación anticipada o "
    "modificación del contrato de concesión. En tal sentido, no se configura el hecho "
    "gravado básico del IVA contemplado en los artículos 2° y 8° del DL 825, ya que "
    "no existe una convención que transfiera bienes ni una remuneración por servicios "
    "prestados a la que se asocie dicho pago."
)

parrafo(
    "Esta conclusión es consistente con la jurisprudencia administrativa del SII, que "
    "distingue entre pagos que forman parte del precio de la convención —afectos a "
    "IVA— y pagos que tienen por objeto solucionar indemnizaciones de perjuicios, "
    "que quedan fuera del hecho gravado. La excepción, que confirma la regla, es "
    "precisamente la cláusula penal: el pago efectuado en virtud de ella no forma "
    "parte del precio ni queda afecto al tributo."
)

# ── 1c ──
titulo_seccion("c) Ingreso mínimo garantizado", 2)

parrafo("Tratamiento frente al Impuesto a la Renta:", bold=True, space_before=6)

parrafo(
    "El ingreso mínimo garantizado (IMG) es una garantía otorgada por el Estado al "
    "Concesionario consistente en que, cuando los ingresos efectivos de la explotación "
    "no alcancen un umbral mínimo previamente establecido en el contrato de concesión, "
    "el Fisco se obliga a complementar la diferencia hasta completar dicho nivel mínimo. "
    "Su propósito es reducir el riesgo de demanda que asume el inversionista privado."
)

parrafo(
    "Desde el punto de vista tributario, las sumas que el Fisco se obliga a pagar al "
    "Concesionario en virtud del ingreso mínimo garantizado constituyen un mayor precio "
    "o mayor pago de los servicios que comprende el contrato de concesión. En consecuencia, "
    "deben formar parte de los ingresos brutos del Concesionario para el cumplimiento "
    "de sus obligaciones tributarias frente al IDPC. Así lo instruye expresamente la "
    "Circular N°49 de 2006 del SII: \"Las sumas que el Fisco se obliga a pagar al "
    "Concesionario por alcanzar niveles de ingresos mínimos establecidos previamente "
    "por la explotación de la obra, constituyen para éste un mayor pago o precio de "
    "los servicios que comprende el contrato de concesión, y por lo tanto, deben "
    "formar parte de los ingresos brutos para el cumplimiento de sus obligaciones "
    "tributarias.\""
)

parrafo(
    "Estos ingresos se devengan en el ejercicio en que el derecho al cobro se hace "
    "exigible, conforme a los términos del contrato de concesión. El tratamiento "
    "tributario es análogo al de cualquier otro ingreso proveniente de la explotación "
    "de la concesión, siendo la base imponible del IDPC la diferencia entre el total "
    "percibido y la cuota de amortización del costo de construcción que corresponda "
    "al período."
)

parrafo("Tratamiento frente al IVA:", bold=True, space_before=4)

parrafo(
    "Al constituir el ingreso mínimo garantizado un mayor precio de los servicios "
    "concesionados, este pago integra la base imponible del IVA conforme a los "
    "artículos 16 letras c) y h) del DL 825, en los mismos términos que los demás "
    "ingresos provenientes de la explotación de la obra, de acuerdo con las "
    "instrucciones de la Circular N°44 de 2006 del SII."
)

parrafo(
    "En consecuencia, el Concesionario deberá emitir la factura correspondiente "
    "al Ministerio de Obras Públicas o a la autoridad concedente dentro del mes en "
    "que perciba dicho ingreso, aplicando la base imponible correspondiente a la "
    "proporción de los ingresos asignada a los servicios de conservación, reparación "
    "y explotación —habitualmente el 20% del total percibido, descontada la parte "
    "que se imputa al pago del precio de la construcción—, sobre la cual se aplicará "
    "la tasa de IVA vigente."
)

parrafo(
    "Es importante tener en cuenta que la fracción del ingreso mínimo garantizado "
    "que corresponda a la amortización del costo de construcción —el porcentaje "
    "habitualmente acordado es el 80%— no se encontrará gravada con IVA, por haberse "
    "ya tributado con dicho impuesto en la etapa de construcción al momento de emitirse "
    "los estados de pago correspondientes."
)

# ══════════════════════════════════════════════════════════════════
# PREGUNTA 2
# ══════════════════════════════════════════════════════════════════
titulo_seccion("2. TRATAMIENTO FRENTE AL IMPUESTO A LA RENTA DE LOS PAGOS POR EXPROPIACIONES", 1)
linea_separadora()

parrafo(
    "En el marco de los contratos de concesión de obras de uso público, el Concesionario "
    "puede verse en la necesidad de incurrir en desembolsos por concepto de expropiaciones, "
    "uso de infraestructura preexistente, supervisión y control, estudios efectuados por "
    "el Fisco, y administración del contrato. El tratamiento tributario de estos pagos "
    "frente al impuesto a la renta depende de la etapa del proyecto en que se efectúen.", space_before=6
)

titulo_seccion("a) Pagos efectuados durante la etapa de construcción", 3)

parrafo(
    "Los pagos que el Concesionario realiza al Estado —incluidas las expropiaciones— "
    "con anterioridad a la explotación de la obra y que son necesarios para su "
    "construcción, forman parte del costo de construcción de la obra. Ello es así "
    "porque el artículo 15 de la LIR establece que el ingreso tributario del "
    "Concesionario en la etapa de construcción es equivalente al costo total incurrido "
    "en dicha etapa —esto es, Ingreso Tributario = Costo Tributario—, el que se "
    "devenga en el ejercicio en que se inicia la explotación efectiva de la obra.", space_before=6
)

parrafo(
    "Dentro de los elementos que conforman el costo de construcción, la Circular N°49 "
    "de 2006 del SII incluye expresamente, entre otros: la mano de obra, los materiales, "
    "los servicios utilizados, los gastos de financiamiento y la subcontratación total "
    "o parcial de la obra. Los pagos por expropiaciones y por uso de infraestructura "
    "preexistente que el Concesionario deba efectuar al MOP en esta etapa integran "
    "dicho costo, tal como lo refleja la práctica contable y tributaria de las "
    "sociedades concesionarias, en que estas partidas aparecen registradas directamente "
    "en la cuenta de costo de construcción."
)

parrafo(
    "En consecuencia, estos desembolsos no son deducibles como gasto en el período "
    "en que se incurran, sino que se activan como parte del costo de la obra y se "
    "imputan al resultado tributario en el ejercicio en que se devenga el ingreso "
    "por el contrato de construcción —es decir, al inicio de la explotación—. "
    "Posteriormente, el costo total de construcción así determinado se amortiza "
    "durante el plazo de duración de la concesión, como elemento del gasto diferido "
    "de los ingresos de explotación."
)

titulo_seccion("b) Pagos efectuados durante la etapa de explotación", 3)

parrafo(
    "Los pagos que el Concesionario efectúe al Estado durante la etapa de explotación "
    "de la concesión —por concepto de expropiaciones asociadas a ampliaciones o mejoras "
    "de la obra, uso de infraestructura preexistente, supervisión y control, estudios "
    "efectuados por el Fisco y administración del contrato— tienen la naturaleza de "
    "gastos necesarios para producir la renta de explotación.", space_before=6
)

parrafo(
    "En consecuencia, estos desembolsos son deducibles de los ingresos brutos percibidos "
    "o devengados por la explotación de la concesión, para efectos de la determinación "
    "de la base imponible del IDPC, en la medida en que cumplan con los requisitos "
    "generales del artículo 31 de la LIR, esto es: (i) que sean gastos necesarios para "
    "producir la renta —con aptitud para generar o mantener ingresos—; (ii) que no "
    "hayan sido deducidos como parte del costo de la obra; y (iii) que se encuentren "
    "debidamente documentados y acreditados."
)

parrafo(
    "Para el Estado, estos pagos constituyen un reembolso de gastos o recuperación de "
    "desembolsos incurridos, sin que ello genere un ingreso tributable para el Fisco "
    "en términos del impuesto a la renta, dado que la contraparte es el Estado y no "
    "un contribuyente en el sentido convencional."
)

titulo_seccion("c) Síntesis del tratamiento", 3)

parrafo(
    "A modo de resumen, el tratamiento tributario de los pagos por expropiaciones "
    "efectuados por el Concesionario puede esquematizarse de la siguiente manera:", space_before=6
)

# Tabla resumen
table = doc.add_table(rows=3, cols=3)
table.style = "Table Grid"
table.autofit = True

headers = ["Etapa", "Naturaleza del pago", "Tratamiento Tributario (Renta)"]
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(h)
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.name = "Times New Roman"
    r.font.color.rgb = NEGRO

rows_data = [
    ("Construcción\n(antes de la PSP)",
     "Expropiaciones y costos necesarios para construir la obra",
     "Forman parte del costo de construcción. Se activan y se devengan al inicio de la explotación (Art. 15 LIR y Circular N°49/2006)."),
    ("Explotación\n(después de la PSP)",
     "Expropiaciones por ampliaciones, uso de infraestructura, supervisión y administración",
     "Gastos necesarios para producir la renta. Deducibles de los ingresos brutos de explotación conforme al Art. 31 LIR."),
]
for i, (e, n, t) in enumerate(rows_data):
    for j, txt in enumerate([e, n, t]):
        cell = table.cell(i + 1, j)
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(txt)
        r.font.size = Pt(10)
        r.font.name = "Times New Roman"
        r.font.color.rgb = NEGRO

doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════
# CONCLUSIONES
# ══════════════════════════════════════════════════════════════════
titulo_seccion("CONCLUSIONES", 1)
linea_separadora()

parrafo(
    "Del análisis efectuado se desprenden las siguientes conclusiones principales:", space_before=6
)

conclusiones = [
    ("Premios: ", "Constituyen un mayor ingreso bruto afecto al IDPC y a IVA cuando "
     "sean contraprestación por los servicios concesionados. Se incorporan a los "
     "ingresos de explotación conforme a la Circular N°49 y N°44 de 2006 del SII."),
    ("Indemnizaciones: ", "Constituyen renta para el Concesionario y quedan afectas al "
     "IDPC. El daño emergente debidamente acreditado puede rebajarse como gasto (Art. 31 LIR). "
     "No están afectas a IVA, pues no son precio de servicios sino compensación por perjuicios."),
    ("Ingreso Mínimo Garantizado: ", "Es un mayor precio de los servicios concesionados. "
     "Afecto al IDPC como ingreso bruto de explotación y afecto a IVA en la proporción "
     "correspondiente a los servicios de explotación (típicamente el 20%), conforme a "
     "los Arts. 16 letras c) y h) del DL 825 y Circular N°44 de 2006."),
    ("Expropiaciones (etapa construcción): ", "Forman parte del costo de la obra. Se "
     "activan y se imputan al resultado tributario al inicio de la explotación, "
     "amortizándose durante el plazo de la concesión (Art. 15 LIR, Circular N°49/2006)."),
    ("Expropiaciones (etapa explotación): ", "Son gastos necesarios para producir la "
     "renta, deducibles de los ingresos brutos de explotación conforme al Art. 31 LIR, "
     "siempre que cumplan los requisitos generales de necesidad, vinculación al giro "
     "y debida documentación."),
]

for i, (titulo, texto) in enumerate(conclusiones):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r1 = p.add_run(f"{i+1}. {titulo}")
    r1.font.size = Pt(12); r1.font.bold = True
    r1.font.color.rgb = NEGRO; r1.font.name = "Times New Roman"
    r2 = p.add_run(texto)
    r2.font.size = Pt(12); r2.font.bold = False
    r2.font.color.rgb = NEGRO; r2.font.name = "Times New Roman"

# ══════════════════════════════════════════════════════════════════
# NORMATIVA CONSULTADA
# ══════════════════════════════════════════════════════════════════
titulo_seccion("NORMATIVA Y FUENTES CONSULTADAS", 1)
linea_separadora()

fuentes = [
    "Decreto Ley N°824 de 1974, Ley sobre Impuesto a la Renta (LIR), en especial Arts. 15, 17, 20 y 31.",
    "Decreto Ley N°825 de 1974, Ley sobre Impuesto a las Ventas y Servicios (LIVS), Arts. 16 letras c) y h), y Art. 55.",
    "Decreto MOP N°900 de 1996, que fija el Texto Refundido, Coordinado y Sistematizado del DFL MOP N°164 de 1991, Ley de Concesiones de Obras Públicas.",
    "Decreto MOP N°956 de 1997, Reglamento de la Ley de Concesiones de Obras Públicas.",
    "Ley N°19.460 de 13 de julio de 1996.",
    "Circular N°49 de 2006, SII — Imparte instrucciones sobre tributación aplicable a los contratos de concesión de obras de uso público (Impuesto a la Renta).",
    "Circular N°44 de 2006, SII — Imparte instrucciones sobre tributación aplicable a los contratos de concesión de obras de uso público (IVA).",
    "Niño Roa, Ximena. \"Concesiones de Obras de Uso Público\", presentación Tributación Sectorial, Magíster en Dirección Tributaria UVM, enero 2026.",
]

for fuente in fuentes:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(fuente)
    r.font.size = Pt(11)
    r.font.color.rgb = NEGRO
    r.font.name = "Times New Roman"

doc.save(OUTPUT)
print(f"✔  Documento generado: {OUTPUT}")
