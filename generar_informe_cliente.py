#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
import datetime

OUTPUT = "/home/user/Prueba/Asesoria_Tributaria_Cliente.pdf"

AZUL    = colors.HexColor("#1A3A5C")
CELESTE = colors.HexColor("#2E6DA4")
GRIS    = colors.HexColor("#F4F6F9")
NARANJA = colors.HexColor("#C0531A")
VERDE   = colors.HexColor("#1A6B3C")
ROJO    = colors.HexColor("#B03030")
NEGRO   = colors.HexColor("#1C1C1C")
PURPURA = colors.HexColor("#6B3A8C")
TEAL    = colors.HexColor("#1A6B5C")

doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=letter,
    rightMargin=2.2*cm, leftMargin=2.2*cm,
    topMargin=2.5*cm,   bottomMargin=2.2*cm,
    title="Asesoría Tributaria – Cliente – Portafolio Inmobiliario"
)

st = getSampleStyleSheet()

def S(name, parent="Normal", **kw):
    return ParagraphStyle(name, parent=st[parent], **kw)

titulo_doc = S("TD", fontSize=20, textColor=colors.white, alignment=TA_CENTER,
               leading=26, fontName="Helvetica-Bold")
sub_doc    = S("SD", fontSize=11, textColor=colors.HexColor("#BDD7EE"),
               alignment=TA_CENTER, leading=16, fontName="Helvetica")
fecha_sty  = S("FS", fontSize=9, textColor=colors.HexColor("#9EC5E8"),
               alignment=TA_CENTER, fontName="Helvetica-Oblique")
sec_title  = S("ST", fontSize=13, textColor=AZUL, fontName="Helvetica-Bold",
               spaceBefore=10, spaceAfter=4, leading=17)
sub_title  = S("SbT", fontSize=10.5, textColor=CELESTE, fontName="Helvetica-Bold",
               spaceBefore=7, spaceAfter=3, leading=14)
cuerpo     = S("Cu", fontSize=9.5, textColor=NEGRO, leading=14,
               spaceAfter=5, alignment=TA_JUSTIFY)
bala       = S("Ba", fontSize=9.5, textColor=NEGRO, leading=13,
               leftIndent=14, spaceAfter=3)
alerta     = S("Al", fontSize=9.5, textColor=ROJO, fontName="Helvetica-Bold",
               leading=13, leftIndent=12, spaceAfter=4)
positivo   = S("Po", fontSize=9.5, textColor=VERDE, fontName="Helvetica-Bold",
               leading=13, leftIndent=12, spaceAfter=4)
atencion   = S("At", fontSize=9.5, textColor=NARANJA, fontName="Helvetica-Bold",
               leading=13, leftIndent=12, spaceAfter=4)
nota       = S("No", fontSize=8.5, textColor=colors.HexColor("#555555"),
               leading=12, leftIndent=10, fontName="Helvetica-Oblique")
tc         = S("Tc", fontSize=8.5, textColor=NEGRO, leading=12, alignment=TA_LEFT)
th         = S("Th", fontSize=8.5, textColor=colors.white, fontName="Helvetica-Bold",
               alignment=TA_CENTER, leading=12)
disclaimers= S("Di", fontSize=7.5, textColor=colors.HexColor("#888888"),
               leading=11, alignment=TA_JUSTIFY, fontName="Helvetica-Oblique")
prop_num   = S("PN", fontSize=14, textColor=colors.white, fontName="Helvetica-Bold",
               alignment=TA_CENTER, leading=18)
prop_name  = S("PNm", fontSize=10, textColor=colors.white, fontName="Helvetica",
               alignment=TA_CENTER, leading=14)

def hr(color=CELESTE, w=0.8):
    return HRFlowable(width="100%", thickness=w, color=color, spaceAfter=5, spaceBefore=5)

def section(texto, color=AZUL):
    return [
        Spacer(1, 0.25*cm),
        Table([[Paragraph(texto, sec_title)]],
              colWidths=["100%"],
              style=TableStyle([
                  ("BACKGROUND",    (0,0),(-1,-1), GRIS),
                  ("LEFTPADDING",   (0,0),(-1,-1), 10),
                  ("RIGHTPADDING",  (0,0),(-1,-1), 10),
                  ("TOPPADDING",    (0,0),(-1,-1), 7),
                  ("BOTTOMPADDING", (0,0),(-1,-1), 7),
                  ("LINEBELOW",     (0,0),(-1,-1), 2, color),
              ])),
        Spacer(1, 0.15*cm),
    ]

def tabla_kv(titulo, filas, col1=5.5*cm, col2=12.3*cm, hcolor=AZUL):
    data = [[Paragraph(titulo, th), ""]]
    for k, v in filas:
        data.append([Paragraph(k, S("k", fontSize=8.5, fontName="Helvetica-Bold",
                                     textColor=NEGRO, leading=12)),
                     Paragraph(v, tc)])
    t = Table(data, colWidths=[col1, col2])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), hcolor),
        ("SPAN",          (0,0),(-1,0)),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, GRIS]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#CCCCCC")),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("RIGHTPADDING",  (0,0),(-1,-1), 8),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
    ]))
    return [t, Spacer(1, 0.25*cm)]

def tabla_comp(headers, filas, anchos=None):
    data = [[Paragraph(h, th) for h in headers]]
    for fila in filas:
        data.append([Paragraph(c, tc) for c in fila])
    n = len(headers)
    if anchos is None:
        anchos = [17.8*cm / n] * n
    t = Table(data, colWidths=anchos)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), AZUL),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, GRIS]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#CCCCCC")),
        ("LEFTPADDING",   (0,0),(-1,-1), 7),
        ("RIGHTPADDING",  (0,0),(-1,-1), 7),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
    ]))
    return [t, Spacer(1, 0.25*cm)]

def badge_box(texto, bg=CELESTE):
    t = Table([[Paragraph(texto, S("bg", fontSize=9.5, textColor=colors.white,
                                    fontName="Helvetica-Bold", leading=13,
                                    alignment=TA_LEFT))]],
              colWidths=["100%"])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), bg),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
    ]))
    return [t, Spacer(1, 0.1*cm)]

def prop_header(numero, titulo, subtitulo, color):
    data = [
        [Paragraph(f"PROPIEDAD {numero}", prop_num)],
        [Paragraph(titulo, prop_name)],
        [Paragraph(subtitulo, S("psub", fontSize=8.5, textColor=colors.HexColor("#BDD7EE"),
                                 alignment=TA_CENTER, fontName="Helvetica-Oblique"))],
    ]
    t = Table(data, colWidths=["100%"])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), color),
        ("TOPPADDING",    (0,0),(-1,-1), 10),
        ("BOTTOMPADDING", (0,0),(-1,-1), 10),
        ("LEFTPADDING",   (0,0),(-1,-1), 15),
        ("RIGHTPADDING",  (0,0),(-1,-1), 15),
    ]))
    return [t, Spacer(1, 0.2*cm)]

# ══════════════════════════════════════════════════════════════════════════════
story = []

# ─── PORTADA ─────────────────────────────────────────────────────────────────
portada = Table([
    [Paragraph("INFORME DE ASESORÍA TRIBUTARIA", titulo_doc)],
    [Spacer(1, 0.3*cm)],
    [Paragraph("Portafolio Inmobiliario — 4 Propiedades", sub_doc)],
    [Spacer(1, 0.4*cm)],
    [Paragraph("DFL2 · Airbnb · Promesa de Compra · Venta a Mediano Plazo", sub_doc)],
    [Spacer(1, 0.7*cm)],
    [Paragraph(f"Fecha de emisión: {datetime.date.today().strftime('%d de %B de %Y')}", fecha_sty)],
    [Paragraph("Confidencial — Uso exclusivo del cliente", fecha_sty)],
], colWidths=["100%"])
portada.setStyle(TableStyle([
    ("BACKGROUND",    (0,0),(-1,-1), AZUL),
    ("TOPPADDING",    (0,0),(-1,-1), 18),
    ("BOTTOMPADDING", (0,0),(-1,-1), 18),
    ("LEFTPADDING",   (0,0),(-1,-1), 20),
    ("RIGHTPADDING",  (0,0),(-1,-1), 20),
]))
story += [portada, Spacer(1, 0.5*cm)]

# ─── RESUMEN PORTAFOLIO ───────────────────────────────────────────────────────
story += section("RESUMEN DEL PORTAFOLIO — 4 PROPIEDADES")
story.append(Paragraph(
    "El cliente persona natural posee 3 departamentos y tiene una promesa de compra sobre "
    "un cuarto. Su objetivo de mediano plazo es vender todas las propiedades para adquirir "
    "un departamento de uso personal.", cuerpo))

data_port = [
    [Paragraph("#", th), Paragraph("Descripción", th), Paragraph("Costo", th),
     Paragraph("Financiamiento", th), Paragraph("Uso / Estado", th),
     Paragraph("DFL2", th), Paragraph("Adquisición", th)],
    [Paragraph("1", tc),
     Paragraph("Dpto. DFL2 — Arriendo Airbnb", tc),
     Paragraph("2.100 UF", tc),
     Paragraph("Mutuaria", tc),
     Paragraph("Arriendo amoblado (Airbnb)", tc),
     Paragraph("✔ DFL2\n1ª de 2", tc),
     Paragraph("Actual", tc)],
    [Paragraph("2", tc),
     Paragraph("Dpto. DFL2 — Ñuñoa", tc),
     Paragraph("Por determinar", tc),
     Paragraph("Crédito hipotecario", tc),
     Paragraph("A definir", tc),
     Paragraph("✔ DFL2\n2ª de 2\n(límite alcanzado)", tc),
     Paragraph("Actual", tc)],
    [Paragraph("3", tc),
     Paragraph("Dpto. — Ñuñoa\n(sin DFL2)", tc),
     Paragraph("Por determinar", tc),
     Paragraph("Por determinar", tc),
     Paragraph("A definir", tc),
     Paragraph("✖ Sin DFL2\n(3ª propiedad)", tc),
     Paragraph("2026", tc)],
    [Paragraph("4", tc),
     Paragraph("Dpto. — Promesa de compra\n(entrega 1er sem. 2029)", tc),
     Paragraph("Por determinar", tc),
     Paragraph("Por definir", tc),
     Paragraph("Sin uso — pendiente entrega", tc),
     Paragraph("✖ Sin DFL2\n(4ª propiedad)", tc),
     Paragraph("Promesa\nfirmada", tc)],
]
t_port = Table(data_port, colWidths=[0.6*cm, 3.8*cm, 2*cm, 2.8*cm, 3.2*cm, 2.5*cm, 2.9*cm])
t_port.setStyle(TableStyle([
    ("BACKGROUND",    (0,0),(-1,0), AZUL),
    ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, GRIS,
                                      colors.HexColor("#FFF8F0"),
                                      colors.HexColor("#F5F0FF")]),
    ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#CCCCCC")),
    ("LEFTPADDING",   (0,0),(-1,-1), 5),
    ("RIGHTPADDING",  (0,0),(-1,-1), 5),
    ("TOPPADDING",    (0,0),(-1,-1), 5),
    ("BOTTOMPADDING", (0,0),(-1,-1), 5),
    ("VALIGN",        (0,0),(-1,-1), "TOP"),
    ("FONTSIZE",      (0,0),(-1,-1), 8),
]))
story += [t_port, Spacer(1, 0.3*cm)]

story.append(Paragraph(
    "⚠ Alerta DFL2: La Ley N°21.420 de 2022 limita el beneficio DFL2 a máximo 2 propiedades "
    "por persona natural. Las propiedades 3 y 4 no gozan de este beneficio. "
    "Sus arriendos siempre tributarán como renta de 1ª Categoría.", alerta))
story.append(Paragraph(
    "⚠ Alerta habitualidad: La venta de 4 propiedades en un período relativamente corto "
    "eleva el riesgo de que el SII califique al cliente como vendedor habitual de inmuebles, "
    "con potencial aplicación de IVA en las ventas y pérdida del INR de 8.000 UF.", alerta))

# ══════════════════════════════════════════════════════════════════════════════
# P1
story.append(PageBreak())
story += prop_header("1", "Departamento DFL2 — Arriendo Airbnb",
                     "Costo: 2.100 UF · Financiamiento: Mutuaria · DFL2 N°1", CELESTE)

story += section("P1 — FICHA")
story += tabla_kv("DATOS CLAVE — PROPIEDAD 1", [
    ("Clasificación",   "DFL2 — Propiedad N°1 de las 2 permitidas (Ley N°21.420)"),
    ("Costo",           "2.100 UF (supera el tope de 2.000 UF para CEEC — no aplica crédito constructora)"),
    ("Financiamiento",  "Mutuaria (mutuo hipotecario). Intereses deducibles como gasto necesario."),
    ("Uso actual",      "Arriendo temporal vía Airbnb → inmueble amoblado"),
    ("Adquisición",     "Vigente — ya es propietario"),
])

story += section("P1 — IVA: AIRBNB = AMOBLADO = AFECTO A IVA", color=NARANJA)
story += badge_box("⚠ Arriendo amoblado (Airbnb) constituye hecho gravado Art. 8 g) DL 825 → IVA 19%", NARANJA)
for item in [
    "<b>Obligación de ser contribuyente de IVA</b> e inscribirse en el SII.",
    "<b>Base imponible:</b> renta mensual MENOS 11% del avalúo fiscal anual ÷ 12 (deducción obligatoria, Ley N°21.210).",
    "<b>Documento:</b> debe emitir boleta o factura por cada período de arriendo.",
    "<b>Riesgo si no ha declarado IVA:</b> impuesto adeudado + intereses (1,5% mensual) + multas. "
    "Se recomienda revisar y regularizar con urgencia.",
]:
    story.append(Paragraph(f"• {item}", bala))

story += section("P1 — RENTA: NO ES INR (AMOBLADO)", color=CELESTE)
story += tabla_comp(
    ["Uso", "INR DFL2", "Impuesto Renta", "IVA"],
    [
        ["Airbnb — amoblado (situación actual)", "✖ No aplica", "Sí — 1ª Cat. IDPC", "Sí — Art. 8 g)"],
        ["Si se arrendara sin amoblar", "✔ Sí (DFL2 N°1)", "No", "No"],
    ],
    anchos=[5*cm, 2.5*cm, 4.8*cm, 5.5*cm]
)
story.append(Paragraph(
    "✔ Los intereses del mutuo hipotecario (Mutuaria) son deducibles como gasto "
    "necesario contra la renta de 1ª Categoría. Reducen la base imponible del IDPC.", positivo))

story += section("P1 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor — Propiedad 1 (PN sin contabilidad completa)", [
    ("Régimen", "Art. 17 N°8 b) LIR"),
    ("INR", "Hasta 8.000 UF acumuladas globales (límite vitalicio compartido con P2, P3 y P4)"),
    ("Condición plazo", "Mínimo 1 año desde adquisición + venta a no relacionado → aplica INR"),
    ("Si vende antes de 1 año", "Impuesto Único Sustitutivo 10%, sin INR"),
    ("Exceso sobre 8.000 UF", "IUS 10% o IGC/IA a elección"),
    ("IVA en la venta", "No aplica si el cliente no es vendedor habitual. El SII califica."),
    ("Costo tributario", "Precio adquisición reajustado por IPC + mejoras declaradas al SII"),
])

# ══════════════════════════════════════════════════════════════════════════════
# P2
story.append(PageBreak())
story += prop_header("2", "Departamento DFL2 — Ñuñoa",
                     "Crédito Hipotecario · DFL2 N°2 (última permitida)", TEAL)

story += section("P2 — FICHA")
story += tabla_kv("DATOS CLAVE — PROPIEDAD 2", [
    ("Ubicación",      "Ñuñoa"),
    ("Clasificación",  "DFL2 — Propiedad N°2 de 2 (con esta se agota el límite Ley N°21.420)"),
    ("Financiamiento", "Crédito hipotecario (banco)"),
    ("Uso actual",     "Por definir"),
    ("Adquisición",    "Vigente — ya es propietario"),
])

story += section("P2 — ESCENARIOS DE ARRIENDO", color=CELESTE)
story += tabla_comp(
    ["Escenario", "IVA", "Renta", "Recomendación"],
    [
        ["No amoblado", "No", "INR — DFL2 (exento)", "✔ Opción más eficiente: ingreso libre de impuesto"],
        ["Amoblado / Airbnb", "Sí — 19%", "1ª Cat. IDPC", "⚠ Pierde INR. Requiere inscripción IVA."],
        ["Sin arrendar", "No", "No", "Sin impacto corriente"],
    ],
    anchos=[3.5*cm, 2*cm, 3.3*cm, 9*cm]
)
story.append(Paragraph(
    "✔ Si se arrienda sin amoblar, los ingresos son 100% INR bajo DFL2. "
    "Es la segunda y última propiedad con este beneficio.", positivo))
story.append(Paragraph(
    "✔ Intereses del crédito hipotecario son deducibles solo si genera renta de 1ª Categoría "
    "(arriendo amoblado). Si el ingreso es INR, no hay deducción de gastos aplicable.", positivo))

story += section("P2 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor — Propiedad 2", [
    ("Régimen", "Art. 17 N°8 b) LIR — PN sin contabilidad completa"),
    ("INR global", "Compartido con P1, P3 y P4. Límite acumulado total: 8.000 UF de por vida."),
    ("Condición", "Mínimo 1 año desde adquisición, venta a no relacionado"),
    ("IVA en venta", "No aplica si no es vendedor habitual"),
])

# ══════════════════════════════════════════════════════════════════════════════
# P3
story.append(PageBreak())
story += prop_header("3", "Departamento — Ñuñoa (sin DFL2)",
                     "Adquirido en 2026 · 3ª Propiedad — sin beneficio DFL2", NARANJA)

story += section("P3 — FICHA")
story += tabla_kv("DATOS CLAVE — PROPIEDAD 3", [
    ("Ubicación",      "Ñuñoa"),
    ("Clasificación",  "SIN DFL2 — es la 3ª propiedad del cliente (Ley N°21.420 permite solo 2)"),
    ("Adquisición",    "Año 2026 — ya es propietario"),
    ("IVA en compra",  "Si compró a constructora/inmobiliaria habitual, pagó IVA. "
                        "Como PN sin empresa, ese IVA fue costo (no recuperable)."),
    ("Financiamiento", "Por determinar"),
    ("Uso actual",     "Por definir"),
])

story += section("P3 — TRIBUTACIÓN DEL ARRIENDO", color=CELESTE)
story.append(Paragraph(
    "Al no tener beneficio DFL2, cualquier arriendo de P3 tributa como renta de 1ª Categoría, "
    "independientemente de si está amoblada o no:", cuerpo))
story += tabla_comp(
    ["Escenario", "IVA", "Renta", "Observación"],
    [
        ["No amoblado", "No", "1ª Cat. IDPC — Art. 20 N°1", "Sin INR (no es DFL2)"],
        ["Amoblado / Airbnb", "Sí — 19%", "1ª Cat. IDPC", "IVA + Renta. Requiere inscripción IVA."],
    ],
    anchos=[3.5*cm, 2*cm, 4.3*cm, 8*cm]
)

story += section("P3 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor — Propiedad 3", [
    ("Régimen",        "Art. 17 N°8 b) LIR — PN sin contabilidad completa"),
    ("Plazo para INR", "Adquirida en 2026 → desde 2027 cumple el mínimo de 1 año. "
                        "Venta a no relacionado. INR hasta el saldo disponible de las 8.000 UF globales."),
    ("Costo tributario","Precio adquisición 2026 reajustado por IPC + mejoras declaradas al SII"),
    ("INR disponible", "Depende de cuánto ya se haya consumido con las ventas de P1 y P2. "
                        "Si se vende P3 después de P1 y P2, podría quedar saldo reducido."),
    ("IVA en venta",   "Riesgo de habitualidad crece al ser la 3ª venta. Ver sección de alertas."),
])

# ══════════════════════════════════════════════════════════════════════════════
# P4
story.append(PageBreak())
story += prop_header("4", "Departamento — Promesa de Compra",
                     "Entrega 1er Semestre 2029 · Sin DFL2 · 4ª Propiedad", PURPURA)

story += section("P4 — FICHA")
story += tabla_kv("DATOS CLAVE — PROPIEDAD 4", [
    ("Estado",         "Promesa de compra firmada — aún no se ha escriturado ni entregado"),
    ("Entrega",        "1er semestre 2029"),
    ("Clasificación",  "SIN DFL2 — es la 4ª propiedad (muy por sobre el límite de 2)"),
    ("Financiamiento", "Por definir (se definirá antes de la escrituración en 2029)"),
    ("Uso proyectado", "A definir al momento de la entrega"),
])

story += section("P4 — IVA EN LA COMPRA (PROMESA Y ESCRITURA)", color=CELESTE)
story.append(Paragraph(
    "La promesa de compra no transfiere el dominio del inmueble. El IVA se devenga "
    "normalmente al momento de la <b>escrituración y entrega del inmueble en 2029</b>. "
    "Sin embargo, si en la promesa se pactan anticipos o cuotas, el SII puede exigir "
    "IVA en la medida que se perciban dichos pagos.", cuerpo))

story += tabla_kv("IVA en P4", [
    ("Vendedor habitual",
     "La inmobiliaria o constructora es vendedor habitual → la compra estará afecta a IVA 19%."),
    ("Base imponible",
     "Precio de venta MENOS valor de adquisición del terreno (reajustado IPC, sin tope desde 2020). "
     "Se debe deducir obligatoriamente el valor comercial del terreno."),
    ("IVA pagado como PN",
     "Como persona natural sin empresa, el IVA de la compra es un COSTO directo. No se puede usar "
     "como crédito fiscal porque el cliente no es contribuyente de IVA (salvo si tiene empresa)."),
    ("CEEC",
     "Verificar precio y permiso de edificación. Si el permiso fue solicitado antes del 30.04.2023 "
     "y la construcción inició antes del 01.01.2027, podría aplicar CEEC parcial (25%) que reduce "
     "el precio final al comprador. Confirmar con la inmobiliaria."),
    ("Cuotas de promesa",
     "Si se pagan cuotas anticipadas en 2026–2028, revisar si la inmobiliaria aplica IVA "
     "en esas cuotas. Solicitar desglose precio terreno/construcción en el contrato de promesa."),
])

story += section("P4 — TRIBUTACIÓN DEL ARRIENDO (DESDE 2029)", color=CELESTE)
story.append(Paragraph(
    "Sin beneficio DFL2 (4ª propiedad), cualquier arriendo tributa como renta de 1ª Categoría:", cuerpo))
story += tabla_comp(
    ["Escenario", "IVA", "Renta"],
    [
        ["No amoblado", "No", "1ª Cat. IDPC — Art. 20 N°1 (no es INR)"],
        ["Amoblado / Airbnb", "Sí — 19%", "1ª Cat. IDPC"],
    ],
    anchos=[4*cm, 2*cm, 11.8*cm]
)

story += section("P4 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor — Propiedad 4", [
    ("Plazo mínimo para INR",
     "La escritura se realiza en 2029 → el plazo de 1 año se cumple desde 2030. "
     "Si se vende antes de 2030 (antes de 1 año): IUS 10% sin INR."),
    ("INR remanente",
     "Si P1, P2 y P3 ya consumieron parte de las 8.000 UF, el saldo disponible para P4 podría "
     "ser mínimo o inexistente. Muy probable que el mayor valor de P4 tribute con IUS 10%."),
    ("Riesgo habitualidad",
     "⚠ Alta alerta: 4 ventas de inmuebles activa fuertemente el criterio de habitualidad del SII. "
     "Se recomienda documentar en cada venta la motivación de financiar vivienda personal."),
    ("IVA en la venta",
     "Si el SII califica habitualidad en las 4 ventas, aplicaría IVA. Planificar tiempos."),
])

story += section("P4 — OPORTUNIDAD: EMPRESA PARA RECUPERAR IVA", color=CELESTE)
story.append(Paragraph(
    "Dado el IVA que pagará en la compra de P4 (2029), el cliente podría evaluar "
    "adquirirla mediante una <b>sociedad inmobiliaria</b> para recuperar ese IVA:", cuerpo))
for item in [
    "La empresa como contribuyente de IVA convierte el IVA de la compra en <b>crédito fiscal</b>.",
    "Si acumula remanente por 2+ meses, puede solicitar devolución vía <b>Art. 27 bis</b> "
    "(Formulario N°3280, plazo SII 20 días hábiles, silencio positivo).",
    "Alternativa: imputar el remanente a PPM u otros impuestos fiscales.",
    "<b>Contrapeso:</b> si P4 está en empresa con contabilidad completa, el mayor valor "
    "en su venta tributa con IDPC + IGC, sin INR de 8.000 UF.",
    "Decisión depende del monto del IVA a recuperar vs. la ganancia esperada. "
    "Recomendable análisis de flujo de caja comparativo antes de la escrituración en 2029.",
]:
    story.append(Paragraph(f"• {item}", bala))

# ══════════════════════════════════════════════════════════════════════════════
# VENTA GLOBAL
story.append(PageBreak())
story += section("ESTRATEGIA DE VENTA — 4 PROPIEDADES")
story.append(Paragraph(
    "El objetivo de vender las 4 propiedades para comprar un departamento personal "
    "requiere una planificación cuidadosa para maximizar el uso del INR de 8.000 UF "
    "y minimizar el riesgo de habitualidad.", cuerpo))

story += section("LÍMITE GLOBAL DE 8.000 UF — INR", color=CELESTE)
story += tabla_kv("Cómo funciona el límite de 8.000 UF", [
    ("¿Qué es?",
     "Límite acumulado de por vida por persona natural: el mayor valor total en ventas de bienes "
     "raíces que puede quedar como INR es 8.000 UF. Se acumula a lo largo de todas las ventas."),
    ("Hasta 8.000 UF acumuladas",
     "INR — sin impuesto. No se declara."),
    ("Exceso sobre 8.000 UF",
     "Opción A: Impuesto Único Sustitutivo (IUS) 10% — conveniente si la tasa marginal IGC es alta.\n"
     "Opción B: IGC/IA según tasas progresivas — puede reliquidar por años. Elegir según contexto."),
    ("Condición de plazo",
     "≥ 1 año desde adquisición (no agrícola) + venta a no relacionado. "
     "Si P3 fue adquirida en 2026, cumple el plazo desde 2027."),
    ("P4 y el plazo",
     "Escriturada en 2029 → plazo de 1 año se cumple en 2030. Si se vende antes: IUS 10% sin INR."),
])

story += section("RIESGO DE HABITUALIDAD — 4 VENTAS", color=ROJO)
story += badge_box("⚠ Vender 4 propiedades activa fuertemente el criterio de habitualidad del SII", ROJO)
for item in [
    "<b>El SII puede calificar al cliente como vendedor habitual</b> si vende múltiples "
    "inmuebles en períodos cercanos, especialmente si son de naturaleza similar.",
    "Consecuencias de la habitualidad: IVA en las ventas (19%) + pérdida del INR "
    "(mayor valor tributaría como renta Art. 20 N°5: IDPC + IGC).",
    "<b>Cómo mitigar el riesgo:</b> (a) espaciar las ventas en el tiempo; "
    "(b) documentar en cada escritura que la motivación es financiar vivienda personal; "
    "(c) no vender las 4 en el mismo año calendario.",
]:
    story.append(Paragraph(f"⚠ {item}", alerta))

story += section("ORDEN DE VENTA SUGERIDO", color=VERDE)
story += tabla_comp(
    ["Orden", "Propiedad", "Motivo estratégico"],
    [
        ["1°", "P2 — DFL2 Ñuñoa",
         "Si tiene menor ganancia, usar parte del INR con menor costo. "
         "Libera el crédito hipotecario. Verificar plazo mínimo 1 año."],
        ["2°", "P3 — Ñuñoa (adq. 2026)",
         "Plazo de 1 año cumplido desde 2027. Sin DFL2 pero INR disponible. "
         "Vender con suficiente distancia temporal de P2."],
        ["3°", "P1 — DFL2 Airbnb",
         "Si tiene mayor valorización, conviene vender después de consumir parte del INR "
         "en P2/P3. Evaluar si el exceso va a IUS 10% o IGC."],
        ["4°", "P4 — Promesa (2029)",
         "Escriturada en 2029. Esperar al menos hasta 2030 para cumplir plazo de 1 año. "
         "Probable que supere el INR restante → prever IUS 10%."],
    ],
    anchos=[1.5*cm, 4*cm, 12.3*cm]
)
story.append(Paragraph(
    "✔ Recomendación adicional: llevar registro tributario del costo de cada propiedad "
    "(precio adquisición + mejoras declaradas al SII) para calcular con precisión el mayor "
    "valor tributable al momento de cada venta.", positivo))

# CUADRO RESUMEN 4 PROPS
story.append(PageBreak())
story += section("CUADRO RESUMEN — LAS 4 PROPIEDADES")

story += tabla_comp(
    ["Aspecto", "P1 — DFL2 Airbnb", "P2 — DFL2 Ñuñoa", "P3 — Ñuñoa 2026", "P4 — Promesa 2029"],
    [
        ["DFL2", "✔ Sí (1ª)", "✔ Sí (2ª)", "✖ No (3ª)", "✖ No (4ª)"],
        ["IVA en compra", "Según vendedor", "Según vendedor",
         "Sí si compró a habitual (2026)", "Sí — inmobiliaria (2029)"],
        ["IVA en arriendo", "✔ Sí — amoblado", "Solo si amoblado", "Solo si amoblado", "Solo si amoblado"],
        ["Renta arriendo", "1ª Cat. (no INR — amoblado)", "INR si no amoblado",
         "1ª Cat. (sin DFL2)", "1ª Cat. (sin DFL2)"],
        ["INR en venta", "✔ Hasta saldo 8.000 UF", "✔ Hasta saldo 8.000 UF",
         "✔ Hasta saldo (adq. 2026)", "✔ Saldo remanente (desde 2030)"],
        ["Plazo mínimo venta", "≥ 1 año desde adquisición", "≥ 1 año desde adquisición",
         "≥ 1 año desde 2026 (cumple desde 2027)", "≥ 1 año desde escritura 2029 (desde 2030)"],
        ["Riesgo habitualidad", "Bajo (aislada)", "Bajo", "Medio (3ª venta)", "Alto (4ª venta)"],
        ["IVA en venta", "No si PN no habitual", "No si PN no habitual",
         "Riesgo medio", "Riesgo alto"],
    ],
    anchos=[3.2*cm, 3.6*cm, 3.4*cm, 3.4*cm, 4.2*cm]
)

# ALERTAS
story += section("ALERTAS Y ACCIONES PRIORITARIAS")

alertas = [
    (ROJO,    "URGENTE",    "1",
     "IVA Airbnb P1 no declarado",
     "Si el cliente arrienda P1 vía Airbnb sin emitir boleta/factura ni inscripción IVA, "
     "está en incumplimiento. Regularizar con el SII a la brevedad para evitar multas e intereses."),
    (ROJO,    "URGENTE",    "2",
     "Verificar si P3 generó IVA en la compra (2026)",
     "Si P3 fue adquirida a una constructora/inmobiliaria en 2026, se pagó IVA. "
     "Confirmar si ese IVA fue correctamente documentado y si hay algún saldo recuperable."),
    (NARANJA, "IMPORTANTE", "3",
     "DFL2 agotado en P3 y P4",
     "P3 y P4 no tienen beneficio DFL2. Sus arriendos siempre tributan como renta 1ª Categoría. "
     "Considerar esto en el análisis de rentabilidad de cada propiedad."),
    (NARANJA, "IMPORTANTE", "4",
     "Promesa de compra P4 — revisar IVA en cuotas",
     "Si se pagan anticipos o cuotas por P4 antes de la escritura, revisar si la inmobiliaria "
     "aplica IVA en esos pagos. Solicitar desglose terreno/construcción en el contrato de promesa."),
    (NARANJA, "IMPORTANTE", "5",
     "Evaluar empresa para P4 (Art. 27 bis)",
     "Antes de escriturar P4 en 2029, analizar si conviene adquirirla mediante sociedad "
     "para recuperar el IVA de la compra vía Art. 27 bis (devolución en 2 meses)."),
    (CELESTE, "PLANIFICAR", "6",
     "Registro de costos y mejoras — 4 propiedades",
     "Llevar registro del precio de adquisición reajustado por IPC y mejoras declaradas al SII "
     "para cada propiedad. Esto reduce el mayor valor tributable al vender."),
    (CELESTE, "PLANIFICAR", "7",
     "Espaciar las 4 ventas — riesgo de habitualidad",
     "No vender las 4 propiedades en el mismo año ni en períodos muy cortos. "
     "Documentar en cada escritura la motivación: financiar compra de vivienda personal."),
    (VERDE,   "OPORTUNIDAD","8",
     "Optimizar uso del INR de 8.000 UF",
     "Calcular la ganancia esperada de cada propiedad. Comenzar por las de mayor ganancia "
     "para usar el INR donde el ahorro fiscal sea máximo. El límite es acumulado y vitalicio."),
]

for color, badge_txt, num, titulo_a, texto_a in alertas:
    data = [
        [Paragraph(badge_txt, S("bt", fontSize=7.5, textColor=colors.white,
                                 fontName="Helvetica-Bold", alignment=TA_CENTER)),
         Paragraph(f"{num}. {titulo_a}", S("at", fontSize=9.5, textColor=AZUL,
                                             fontName="Helvetica-Bold", leading=13)),
         Paragraph(texto_a, S("ax", fontSize=8.5, textColor=NEGRO, leading=12,
                               alignment=TA_JUSTIFY))],
    ]
    t = Table(data, colWidths=[2*cm, 4.5*cm, 11.3*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(0,-1), color),
        ("BACKGROUND",    (1,0),(-1,-1), GRIS),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
        ("LEFTPADDING",   (0,0),(-1,-1), 7),
        ("RIGHTPADDING",  (0,0),(-1,-1), 7),
        ("LINEBELOW",     (0,0),(-1,-1), 0.5, colors.HexColor("#DDDDDD")),
    ]))
    story += [t, Spacer(1, 0.1*cm)]

story += [Spacer(1, 0.4*cm), hr(color=colors.HexColor("#AAAAAA"), w=0.5)]
story.append(Paragraph(
    "NOTA LEGAL: Informe elaborado con fines de asesoría tributaria, basado en el material del "
    "Magíster en Dirección Tributaria UVM y la legislación vigente (DL 825, DL 824, "
    "Ley N°21.210/2020, Ley N°21.420/2022, DFL N°2/1959, Art. 27 bis DL 825). "
    "No constituye opinión legal formal. Verificar circunstancias específicas con asesor "
    "tributario antes de tomar decisiones. La habitualidad es facultad exclusiva del SII.",
    disclaimers))

doc.build(story)
print(f"✔ PDF generado: {OUTPUT}")
