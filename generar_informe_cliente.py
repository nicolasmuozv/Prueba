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

# ── Paleta de colores ──────────────────────────────────────────────────────────
AZUL    = colors.HexColor("#1A3A5C")
CELESTE = colors.HexColor("#2E6DA4")
GRIS    = colors.HexColor("#F4F6F9")
NARANJA = colors.HexColor("#C0531A")
VERDE   = colors.HexColor("#1A6B3C")
ROJO    = colors.HexColor("#B03030")
AMARILLO= colors.HexColor("#F5A623")
NEGRO   = colors.HexColor("#1C1C1C")
GRIS2   = colors.HexColor("#EEEEEE")

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

# Estilos
titulo_doc  = S("TD", fontSize=20, textColor=colors.white, alignment=TA_CENTER,
                leading=26, fontName="Helvetica-Bold")
sub_doc     = S("SD", fontSize=11, textColor=colors.HexColor("#BDD7EE"),
                alignment=TA_CENTER, leading=16, fontName="Helvetica")
fecha_sty   = S("FS", fontSize=9, textColor=colors.HexColor("#9EC5E8"),
                alignment=TA_CENTER, fontName="Helvetica-Oblique")
sec_title   = S("ST", fontSize=13, textColor=AZUL, fontName="Helvetica-Bold",
                spaceBefore=10, spaceAfter=4, leading=17)
sub_title   = S("SbT", fontSize=10.5, textColor=CELESTE, fontName="Helvetica-Bold",
                spaceBefore=7, spaceAfter=3, leading=14)
cuerpo      = S("Cu", fontSize=9.5, textColor=NEGRO, leading=14,
                spaceAfter=5, alignment=TA_JUSTIFY)
bala        = S("Ba", fontSize=9.5, textColor=NEGRO, leading=13,
                leftIndent=14, spaceAfter=3)
alerta      = S("Al", fontSize=9.5, textColor=ROJO, fontName="Helvetica-Bold",
                leading=13, leftIndent=12, spaceAfter=4)
positivo    = S("Po", fontSize=9.5, textColor=VERDE, fontName="Helvetica-Bold",
                leading=13, leftIndent=12, spaceAfter=4)
atencion    = S("At", fontSize=9.5, textColor=NARANJA, fontName="Helvetica-Bold",
                leading=13, leftIndent=12, spaceAfter=4)
nota        = S("No", fontSize=8.5, textColor=colors.HexColor("#555555"),
                leading=12, leftIndent=10, fontName="Helvetica-Oblique")
tc          = S("Tc", fontSize=8.5, textColor=NEGRO, leading=12, alignment=TA_LEFT)
th          = S("Th", fontSize=8.5, textColor=colors.white, fontName="Helvetica-Bold",
                alignment=TA_CENTER, leading=12)
disclaimers = S("Di", fontSize=7.5, textColor=colors.HexColor("#888888"),
                leading=11, alignment=TA_JUSTIFY, fontName="Helvetica-Oblique")
prop_num    = S("PN", fontSize=14, textColor=colors.white, fontName="Helvetica-Bold",
                alignment=TA_CENTER, leading=18)
prop_name   = S("PNm", fontSize=10, textColor=colors.white, fontName="Helvetica",
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

def tabla_kv(titulo, filas, col1=5.5*cm, col2=12.3*cm, header_color=AZUL):
    data = [[Paragraph(titulo, th), ""]]
    for k, v in filas:
        data.append([Paragraph(k, S("k", fontSize=8.5, fontName="Helvetica-Bold",
                                     textColor=NEGRO, leading=12)),
                     Paragraph(v, tc)])
    t = Table(data, colWidths=[col1, col2])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), header_color),
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

def badge(texto, bg=CELESTE, fg=colors.white, width="100%"):
    t = Table([[Paragraph(texto, S("bg", fontSize=9.5, textColor=fg,
                                    fontName="Helvetica-Bold", leading=13,
                                    alignment=TA_LEFT))]],
              colWidths=["100%"])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), bg),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),[4,4,4,4]),
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

# ─── PORTADA ──────────────────────────────────────────────────────────────────
portada = Table([
    [Paragraph("INFORME DE ASESORÍA TRIBUTARIA", titulo_doc)],
    [Spacer(1, 0.3*cm)],
    [Paragraph("Portafolio Inmobiliario — Análisis por Propiedad", sub_doc)],
    [Spacer(1, 0.4*cm)],
    [Paragraph("Persona Natural · DFL2 · Airbnb · Venta a Mediano Plazo", sub_doc)],
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

# ─── RESUMEN DEL PORTAFOLIO ───────────────────────────────────────────────────
story += section("RESUMEN DEL PORTAFOLIO DEL CLIENTE")
story.append(Paragraph(
    "El cliente posee actualmente dos departamentos acogidos al régimen DFL2 y proyecta "
    "adquirir un tercero con entrega en 2029. Su objetivo de mediano plazo es vender las "
    "tres propiedades para financiar la compra de un departamento de uso personal.", cuerpo))

story.append(Spacer(1, 0.1*cm))

# Tabla resumen portafolio
data_port = [
    [Paragraph("#", th), Paragraph("Propiedad", th), Paragraph("Costo", th),
     Paragraph("Financiamiento", th), Paragraph("Uso Actual", th), Paragraph("Estado DFL2", th)],
    [Paragraph("1", tc), Paragraph("Dpto. DFL2 – Ubicación actual", tc),
     Paragraph("2.100 UF", tc), Paragraph("Mutuaria (hipotecario)", tc),
     Paragraph("Arriendo Airbnb (amoblado)", tc),
     Paragraph("✔ DFL2\n(1ª de 2 permitidas)", tc)],
    [Paragraph("2", tc), Paragraph("Dpto. DFL2 – Ñuñoa", tc),
     Paragraph("No indicado", tc), Paragraph("Crédito hipotecario", tc),
     Paragraph("A definir", tc),
     Paragraph("✔ DFL2\n(2ª de 2 permitidas)", tc)],
    [Paragraph("3", tc), Paragraph("Dpto. futuro – entrega 1er sem. 2029", tc),
     Paragraph("Por determinar", tc), Paragraph("Por definir", tc),
     Paragraph("Sin uso aún", tc),
     Paragraph("⚠ Sin beneficio DFL2\n(límite ya alcanzado)", tc)],
]
t_port = Table(data_port, colWidths=[0.7*cm, 4.5*cm, 2*cm, 3.5*cm, 3.5*cm, 3.6*cm])
t_port.setStyle(TableStyle([
    ("BACKGROUND",    (0,0),(-1,0), AZUL),
    ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, GRIS, colors.HexColor("#FFF8F0")]),
    ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#CCCCCC")),
    ("LEFTPADDING",   (0,0),(-1,-1), 6),
    ("RIGHTPADDING",  (0,0),(-1,-1), 6),
    ("TOPPADDING",    (0,0),(-1,-1), 5),
    ("BOTTOMPADDING", (0,0),(-1,-1), 5),
    ("VALIGN",        (0,0),(-1,-1), "TOP"),
    ("FONTSIZE",      (0,0),(-1,-1), 8.5),
    ("BACKGROUND",    (5,3),(5,3), colors.HexColor("#FFF3E0")),
]))
story += [t_port, Spacer(1, 0.3*cm)]

story.append(Paragraph(
    "⚠ Alerta relevante: La Ley N°21.420 de 2022 limitó el beneficio DFL2 a un máximo de "
    "<b>2 propiedades por persona natural</b>. El tercer departamento NO gozará del beneficio "
    "de Ingreso No Renta por arrendamiento ni de las franquicias DFL2.", alerta))

# ══════════════════════════════════════════════════════════════════════════════
# ─── PROPIEDAD 1 ──────────────────────────────────────────────────────────────
story.append(PageBreak())
story += prop_header("1", "Departamento DFL2 — Arriendo Airbnb",
                     "Costo: 2.100 UF · Financiamiento: Mutuaria", CELESTE)

story += section("P1 — FICHA DE LA PROPIEDAD")
story += tabla_kv("DATOS CLAVE — PROPIEDAD 1", [
    ("Clasificación",        "DFL2 (Propiedad N°1 de 2 permitidas bajo Ley N°21.420)"),
    ("Costo de adquisición", "2.100 UF"),
    ("Financiamiento",       "Mutuaria (mutuo hipotecario)"),
    ("Uso actual",           "Arriendo temporal vía Airbnb (inmueble amoblado)"),
    ("CEEC",                 "No aplica: costo de 2.100 UF supera el tope de 2.000 UF para CEEC"),
])

# IVA Airbnb
story += section("P1 — ANÁLISIS IVA: AIRBNB = INMUEBLE AMOBLADO", color=NARANJA)
story.append(Paragraph(
    "El arrendamiento vía Airbnb implica que el departamento se entrega <b>amoblado</b>, "
    "configurando el <b>hecho gravado especial del Art. 8 letra g) del DL 825</b>. "
    "Esto tiene consecuencias tributarias importantes que el cliente debe conocer:", cuerpo))

story += badge("⚠ El arriendo amoblado (Airbnb) está AFECTO A IVA — El cliente debe ser contribuyente de IVA", NARANJA)

for item in [
    "<b>IVA aplicable:</b> 19% sobre la base imponible mensual.",
    "<b>Base imponible:</b> Renta de arriendo mensual <b>menos</b> el 11% del avalúo fiscal anual "
    "del inmueble (deducción obligatoria desde Ley N°21.210).",
    "<b>Ejemplo orientativo:</b> Si el ingreso mensual Airbnb es $800.000 y el avalúo fiscal es "
    "$25.000.000 → Deducción: $229.167 → Base IVA: $570.833 → IVA: $108.458.",
    "<b>Obligación:</b> Emitir factura o boleta por los servicios de arrendamiento. Declarar en F29 mensual.",
    "<b>Riesgo actual:</b> Si el cliente no está inscrito como contribuyente de IVA, "
    "puede estar en incumplimiento tributario expuesto a multas, intereses y recargos.",
]:
    story.append(Paragraph(f"• {item}", bala))

# Renta Airbnb
story += section("P1 — ANÁLISIS RENTA: ¿ES INGRESO NO RENTA (INR)?", color=CELESTE)
story.append(Paragraph(
    "Aunque el departamento es DFL2, el arriendo amoblado (Airbnb) <b>pierde el beneficio "
    "de Ingreso No Renta</b>. El INR del DFL2 solo aplica a arriendos de inmuebles "
    "<b>no amoblados</b>. En este caso:", cuerpo))

story += tabla_comp(
    ["Situación", "INR DFL2", "Impuesto Renta", "IVA"],
    [
        ["Arriendo amoblado (Airbnb) — situación actual",
         "✖ NO aplica", "Sí — 1ª Categoría (IDPC)", "Sí — Art. 8 g) DL 825"],
        ["Si se arrendara sin amoblar",
         "✔ SÍ aplica (DFL2)", "No (INR)", "No"],
    ],
    anchos=[5.5*cm, 2.5*cm, 4.5*cm, 5.3*cm]
)

story.append(Paragraph(
    "✔ Deducción de gastos necesarios: los intereses del mutuo hipotecario (Mutuaria) "
    "son deducibles como gasto necesario para producir la renta, reduciendo la base imponible "
    "del Impuesto de Primera Categoría.", positivo))

story += section("P1 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor en venta — Propiedad 1", [
    ("Régimen aplicable",
     "Art. 17 N°8 b) LIR — Persona Natural sin contabilidad completa"),
    ("INR disponible",
     "Hasta 8.000 UF acumuladas (límite vitalicio global por contribuyente). "
     "Aplica si vende a no relacionado y han transcurrido al menos 1 año desde la adquisición."),
    ("Tasa si excede INR",
     "Impuesto Único Sustitutivo del 10% sobre el exceso, o IGC/IA a elección."),
    ("Costo tributario",
     "Valor de adquisición (2.100 UF) reajustado por IPC + mejoras declaradas al SII."),
    ("Cuidado",
     "Si la venta ocurre antes de 1 año desde la adquisición: no aplica INR. "
     "Tributa con Impuesto Único 10% sin excepción."),
    ("IVA en la venta",
     "En principio NO aplica IVA si el cliente como persona natural no es vendedor habitual "
     "de inmuebles. El SII califica la habitualidad caso a caso."),
])

# ══════════════════════════════════════════════════════════════════════════════
# ─── PROPIEDAD 2 ──────────────────────────────────────────────────────────────
story.append(PageBreak())
story += prop_header("2", "Departamento DFL2 — Ñuñoa",
                     "Financiamiento: Crédito Hipotecario Banco", colors.HexColor("#2E7D5C"))

story += section("P2 — FICHA DE LA PROPIEDAD")
story += tabla_kv("DATOS CLAVE — PROPIEDAD 2", [
    ("Ubicación",            "Ñuñoa"),
    ("Clasificación",        "DFL2 (Propiedad N°2 de 2 permitidas — límite alcanzado)"),
    ("Financiamiento",       "Crédito hipotecario (banco)"),
    ("Uso actual",           "Por definir (no se indica si está arrendada)"),
])

story += section("P2 — ANÁLISIS DE ESCENARIOS DE ARRIENDO", color=CELESTE)
story += tabla_comp(
    ["Escenario", "IVA", "Impuesto Renta", "Recomendación"],
    [
        ["Arriendo NO amoblado (departamento vacío o con mínimo mobiliario)",
         "No aplica", "INR — DFL2 (exento de impuesto a la renta)",
         "✔ Opción más eficiente tributariamente si el objetivo es generar flujo sin carga impositiva."],
        ["Arriendo amoblado (como Airbnb o similar)",
         "Sí — 19% s/ base reducida",
         "1ª Categoría (IDPC 25%-27%)",
         "⚠ Pierde INR. Mayor carga tributaria. Requiere inscripción IVA."],
        ["Uso personal (sin arrendar)",
         "No aplica", "No aplica",
         "Sin impacto corriente. Solo relevante al vender."],
    ],
    anchos=[4.5*cm, 2.5*cm, 3.5*cm, 7.3*cm]
)

story.append(Paragraph(
    "✔ Si el cliente arrienda esta propiedad sin amoblar, los ingresos son 100% INR "
    "(exentos de impuesto a la renta) bajo el beneficio DFL2. Es la segunda y última "
    "propiedad DFL2 con este beneficio.", positivo))

story.append(Paragraph(
    "✔ Los intereses del crédito hipotecario son deducibles como gasto solo si la propiedad "
    "genera rentas de 1ª Categoría (arriendo amoblado o no DFL2). Si es INR, no hay "
    "deducción aplicable ya que el ingreso tampoco está gravado.", positivo))

story += section("P2 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor en venta — Propiedad 2", [
    ("Régimen aplicable",
     "Art. 17 N°8 b) LIR — Persona Natural sin contabilidad completa"),
    ("INR compartido con P1",
     "El límite de 8.000 UF es GLOBAL para el contribuyente, compartido entre todas sus "
     "enajenaciones de bienes raíces. Si la ganancia de P1 + P2 + P3 supera 8.000 UF, "
     "el exceso paga Impuesto Único 10% o IGC."),
    ("Orden de venta",
     "Se recomienda planificar el orden de venta de las propiedades para optimizar el "
     "uso del límite de 8.000 UF disponible."),
    ("IVA en la venta",
     "No aplica si el cliente como PN no es vendedor habitual. El SII califica caso a caso."),
    ("Intereses hipotecarios",
     "Los intereses pagados durante la tenencia NO aumentan el costo tributario del inmueble. "
     "Solo aumentan el costo: precio de adquisición + mejoras declaradas al SII, reajustados por IPC."),
])

# ══════════════════════════════════════════════════════════════════════════════
# ─── PROPIEDAD 3 ──────────────────────════════════════════════════════════════
story.append(PageBreak())
story += prop_header("3", "Departamento Futuro — Entrega 1er Semestre 2029",
                     "Sin DFL2 · Por adquirir · Compra a inmobiliaria", colors.HexColor("#6B3A8C"))

story += section("P3 — ANÁLISIS PREVIO A LA COMPRA")
story.append(Paragraph(
    "Este departamento aún no ha sido adquirido. La planificación previa a la compra es "
    "fundamental para optimizar la carga tributaria desde el origen.", cuerpo))

story += tabla_kv("SITUACIÓN TRIBUTARIA — PROPIEDAD 3", [
    ("Estado DFL2",
     "⚠ SIN beneficio DFL2: el cliente ya tiene 2 propiedades DFL2 (límite máximo Ley N°21.420 de 2022). "
     "Los ingresos de arrendamiento de esta propiedad NO serán INR."),
    ("IVA en la compra",
     "La inmobiliaria es vendedor habitual → la compra estará afecta a IVA (19% sobre base "
     "imponible = precio venta − valor de adquisición del terreno reajustado, sin tope desde 2020). "
     "El cliente pagará IVA en la escritura."),
    ("Crédito fiscal IVA",
     "Como persona natural SIN empresa, el cliente NO puede usar el IVA pagado como crédito fiscal. "
     "El IVA es un costo directo que se suma al precio del inmueble."),
    ("CEEC",
     "Verificar si aplica según permiso de edificación. Desde Ley 21.558: CEEC parcial 25% para "
     "obras con permiso posterior a 30.04.2023 e inicio antes de 01.01.2027."),
    ("Entrega 2029",
     "El vendedor emitirá la factura con IVA al momento de la escrituración (o según contrato de promesa). "
     "Verificar si hay anticipo de IVA en cuotas de promesa."),
])

story += section("P3 — OPORTUNIDAD: EMPRESA PARA RECUPERAR EL IVA", color=CELESTE)
story.append(Paragraph(
    "Dado que la compra de P3 tendrá un IVA significativo (dependiendo del precio), "
    "el cliente podría evaluar si conviene adquirirla a través de una <b>empresa inmobiliaria</b> "
    "para aprovechar el <b>Art. 27 bis</b>:", cuerpo))

story += tabla_kv("¿Conviene empresa para P3?", [
    ("Ventaja principal",
     "Si P3 se compra mediante empresa contribuyente de IVA, el IVA pagado en la compra "
     "es crédito fiscal. Si se acumula remanente por 2+ meses, puede recuperarse mediante "
     "Art. 27 bis (devolución en efectivo desde Tesorería) o imputarse a otros impuestos."),
    ("Condición Art. 27 bis",
     "El inmueble debe destinarse al activo fijo de la empresa. Se considera activo fijo "
     "desde que la obra/etapa es recibida conforme. Mínimo 2 períodos de remanente acumulado."),
    ("Contrapeso (pérdida INR venta)",
     "Si P3 está en activo de empresa con contabilidad completa, al venderla el mayor "
     "valor tributa en régimen general (IDPC + IGC), sin el beneficio de INR de 8.000 UF."),
    ("Pérdida DFL2",
     "Las personas jurídicas no tienen beneficio DFL2. P3 como empresa tampoco tendría INR."),
    ("Conclusión preliminar",
     "La decisión empresa vs PN para P3 depende del monto del IVA a recuperar versus "
     "la ganancia esperada en la venta futura. Se recomienda análisis de flujo de caja "
     "comparativo antes de la compra."),
])

story += section("P3 — VENTA A MEDIANO PLAZO", color=VERDE)
story += tabla_kv("Mayor valor en venta — Propiedad 3 (como PN)", [
    ("Régimen aplicable",
     "Art. 17 N°8 b) LIR — si el cliente la mantiene como PN sin contabilidad completa"),
    ("INR remanente",
     "El límite de 8.000 UF es global. Si P1 y P2 ya consumieron parte, el saldo disponible "
     "para P3 podría ser menor o incluso agotado. Planificación de orden de venta es clave."),
    ("IVA en la venta de P3",
     "Si el cliente vende P3 como PN no habitual: sin IVA. Pero si vende las 3 propiedades "
     "en corto tiempo, el SII podría calificar habitualidad y gravar con IVA las ventas."),
    ("Riesgo de habitualidad",
     "⚠ Vender 3 propiedades en un período corto puede ser calificado por el SII como actividad "
     "habitual de venta de inmuebles, lo que generaría IVA en las ventas y tributación "
     "ordinaria sobre el mayor valor (sin INR)."),
])

# ══════════════════════════════════════════════════════════════════════════════
# ─── VENTA GLOBAL ─────────────────────────────────────────────────────────────
story.append(PageBreak())
story += section("OBJETIVO FINAL: VENTA DE LAS 3 PROPIEDADES")
story.append(Paragraph(
    "El cliente tiene como objetivo de mediano plazo vender los tres departamentos para "
    "financiar la compra de un departamento de uso personal. A continuación se analizan "
    "los aspectos críticos de esta estrategia.", cuerpo))

story += section("INR DE 8.000 UF — LÍMITE GLOBAL", color=CELESTE)
story.append(Paragraph(
    "El Art. 17 N°8 b) LIR establece un límite de <b>8.000 UF acumuladas de por vida</b> "
    "para el beneficio de INR en la venta de bienes raíces por personas naturales. "
    "Este límite es compartido entre todas las enajenaciones que realice el contribuyente.", cuerpo))

story += tabla_kv("Planificación del límite de 8.000 UF", [
    ("¿Cómo se calcula el mayor valor?",
     "Precio de venta − Costo tributario (precio adquisición + mejoras, reajustado por IPC) "
     "− Pérdidas del ejercicio por operaciones similares."),
    ("Hasta 8.000 UF de mayor valor acumulado",
     "INR — sin impuesto. No se declara en F22."),
    ("Mayor valor que excede las 8.000 UF",
     "Opción A: Impuesto Único Sustitutivo del 10% sobre el exceso (base percibida o devengada).\n"
     "Opción B: Impuesto Global Complementario (IGC) según tasas progresivas — puede reliquidar.\n"
     "Conviene comparar ambas opciones según nivel de renta total del año."),
    ("Condición plazo para INR",
     "Venta a no relacionado + al menos 1 año transcurrido desde adquisición (inmueble no agrícola). "
     "Si vende antes de 1 año: paga IUS 10% sin INR."),
    ("Venta a relacionado",
     "Sin INR ni IUS. Tributa con IGC/IA (percibido o devengado, lo que ocurra primero)."),
])

story += section("RIESGO DE HABITUALIDAD EN LAS 3 VENTAS", color=ROJO)
story.append(Paragraph(
    "Si el cliente vende las 3 propiedades en un período relativamente corto, "
    "el SII podría <b>calificar la actividad como habitual en la venta de inmuebles</b>, "
    "con las siguientes consecuencias:", cuerpo))

for item in [
    "<b>IVA en las ventas:</b> las enajenaciones podrían quedar afectas a IVA 19% "
    "(sobre diferencia precio venta − precio compra, excluyendo terreno).",
    "<b>Pérdida del INR:</b> el mayor valor ya no aplicaría el beneficio del Art. 17 N°8 b). "
    "Tributaría como renta del Art. 20 N°5 (régimen general IDPC + IGC).",
    "<b>Obligación de llevar contabilidad</b> si el SII califica la actividad como empresarial.",
]:
    story.append(Paragraph(f"⚠ {item}", alerta))

story.append(Paragraph(
    "Recomendación: Se sugiere que las 3 ventas se espacien en el tiempo y se pueda "
    "acreditar que cada operación responde a la necesidad de financiar una vivienda "
    "personal, no a una actividad lucrativa de compraventa inmobiliaria.", atencion))

story += section("ESTRATEGIA DE VENTA RECOMENDADA", color=VERDE)
story += tabla_comp(
    ["Orden sugerido", "Propiedad", "Motivo"],
    [
        ["1°", "Propiedad 2 — DFL2 Ñuñoa",
         "Si tiene menor valorización, conviene venderla primero para 'usar' parte del INR "
         "de 8.000 UF con menor impacto. Verificar plazo de 1 año desde adquisición."],
        ["2°", "Propiedad 1 — DFL2 Airbnb",
         "Vendida después de P2, permite optimizar el saldo del INR restante. "
         "Evaluar si convendrá IUS 10% o IGC según situación del año."],
        ["3°", "Propiedad 3 — Entrega 2029",
         "Vender con suficiente tiempo de distancia de las anteriores para evitar "
         "calificación de habitualidad. Probablemente exceda el INR de 8.000 UF si las "
         "anteriores ya lo consumieron — prever pago de IUS 10%."],
    ],
    anchos=[2.5*cm, 4.5*cm, 10.8*cm]
)

story.append(Paragraph(
    "✔ Se recomienda llevar un registro tributario del costo de adquisición de cada propiedad "
    "y sus mejoras (debidamente declaradas al SII para ser consideradas en el costo), "
    "lo que minimizará el mayor valor tributable al momento de la venta.", positivo))

# ─── CUADRO RESUMEN FINAL ─────────────────────────────────────────────────────
story.append(PageBreak())
story += section("CUADRO RESUMEN — SITUACIÓN TRIBUTARIA POR PROPIEDAD")

story += tabla_comp(
    ["", "P1 — DFL2 Airbnb", "P2 — DFL2 Ñuñoa", "P3 — Futuro 2029"],
    [
        ["DFL2", "✔ Sí (1ª prop.)", "✔ Sí (2ª prop.)", "✖ No (límite alcanzado)"],
        ["IVA en compra", "Según vendedor al momento de compra", "Según vendedor",
         "✔ Sí — inmobiliaria habitual"],
        ["IVA en arriendo", "✔ Sí — amoblado Airbnb", "Solo si amoblado", "Solo si amoblado"],
        ["Renta arriendo", "1ª Cat. IDPC (NO es INR — amoblado)", "INR si no amoblado",
         "1ª Cat. IDPC (no DFL2)"],
        ["Intereses deducibles", "✔ Sí (Mutuaria) — vs renta 1ª Cat.", "Solo si renta 1ª Cat.",
         "Si arriendo gravado"],
        ["INR venta (8.000 UF)", "✔ Potencial (compartido global)", "✔ Potencial",
         "✔ Si queda saldo disponible"],
        ["IVA en venta", "No si PN no habitual", "No si PN no habitual",
         "Riesgo si vende muy rápido"],
        ["Riesgo habitualidad", "Bajo (1 prop.)", "Bajo (si aislada)", "Medio (3ª venta)"],
    ],
    anchos=[3.8*cm, 4.7*cm, 4.7*cm, 4.6*cm]
)

# ─── ALERTAS PRIORITARIAS ────────────────────────────────────────────────────
story += section("ALERTAS Y ACCIONES PRIORITARIAS")

alertas = [
    (ROJO,    "URGENTE",    "1",
     "IVA Airbnb no declarado",
     "Si el cliente no está inscrito como contribuyente de IVA y viene cobrando arriendos "
     "vía Airbnb sin emitir boleta/factura, está en incumplimiento. Se recomienda revisar "
     "la situación y regularizar ante el SII a la brevedad."),
    (NARANJA, "IMPORTANTE", "2",
     "Límite DFL2 agotado con P3",
     "La adquisición de P3 no tendrá beneficio DFL2. Los arriendos de P3 siempre "
     "tributarán como renta de 1ª Categoría. Considerar esto en el análisis de rentabilidad."),
    (NARANJA, "IMPORTANTE", "3",
     "Planificar compra P3: ¿PN o empresa?",
     "Antes de comprar P3 en 2029, evaluar si conviene adquirirla mediante empresa para "
     "recuperar el IVA de la compra vía Art. 27 bis. Requiere análisis comparativo "
     "con la pérdida del INR de 8.000 UF en la futura venta."),
    (CELESTE, "PLANIFICAR", "4",
     "Registro de costos y mejoras",
     "Llevar registro cronológico del costo de adquisición de cada propiedad, "
     "mejoras realizadas (declaradas al SII en su momento) y reajustes por IPC. "
     "Esto reduce el mayor valor tributable al vender."),
    (CELESTE, "PLANIFICAR", "5",
     "Espaciar las 3 ventas para evitar habitualidad",
     "Planificar que las 3 ventas no ocurran en el mismo año o en períodos muy cercanos. "
     "Documentar que cada venta responde a la necesidad de financiar vivienda personal, "
     "no a actividad habitual de compraventa inmobiliaria."),
    (VERDE,   "OPORTUNIDAD","6",
     "INR de 8.000 UF — planificación de uso",
     "Calcular las ganancias esperadas de cada propiedad y optimizar el orden de venta "
     "para usar el INR de 8.000 UF donde genere mayor ahorro fiscal "
     "(comenzar por las propiedades con mayor ganancia)."),
]

for color, badge_txt, num, titulo_a, texto_a in alertas:
    data = [
        [Paragraph(badge_txt, S("bt", fontSize=8, textColor=colors.white,
                                 fontName="Helvetica-Bold", alignment=TA_CENTER)),
         Paragraph(f"{num}. {titulo_a}", S("at", fontSize=10, textColor=AZUL,
                                             fontName="Helvetica-Bold", leading=14)),
         Paragraph(texto_a, S("ax", fontSize=9, textColor=NEGRO, leading=13,
                               alignment=TA_JUSTIFY))],
    ]
    t = Table(data, colWidths=[2*cm, 4.5*cm, 11.3*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(0,-1), color),
        ("BACKGROUND",    (1,0),(-1,-1), GRIS),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("RIGHTPADDING",  (0,0),(-1,-1), 8),
        ("LINEBELOW",     (0,0),(-1,-1), 0.5, colors.HexColor("#DDDDDD")),
    ]))
    story += [t, Spacer(1, 0.1*cm)]

# ─── DISCLAIMER ──────────────────────────────────────────────────────────────
story += [Spacer(1, 0.4*cm), hr(color=colors.HexColor("#AAAAAA"), w=0.5)]
story.append(Paragraph(
    "NOTA LEGAL: El presente informe ha sido elaborado con fines de asesoría tributaria, "
    "basado en el material del Magíster en Dirección Tributaria UVM y la legislación vigente "
    "(DL 825, DL 824, Ley N°21.210/2020, Ley N°21.420/2022, DFL N°2/1959). "
    "No constituye una opinión legal formal. Se recomienda revisar las circunstancias "
    "específicas de cada operación con el asesor tributario antes de tomar decisiones. "
    "El análisis de habitualidad es facultad exclusiva del SII.",
    disclaimers))

# ─── BUILD ───────────────────────────────────────────────────────────────────
doc.build(story)
print(f"✔ PDF generado: {OUTPUT}")
