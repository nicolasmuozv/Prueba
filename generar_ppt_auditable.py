#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT Auditable — Boleta vs Sociedad para profesionales de la salud
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

OUTPUT = "/home/user/Prueba/Auditable_Infografia.pptx"

# Paleta Auditable
CORAL    = RGBColor(0xF2, 0x5C, 0x4F)
CORAL_LT = RGBColor(0xFC, 0xE4, 0xE1)
GRAPH    = RGBColor(0x2C, 0x2C, 0x2C)
GRAY_DK  = RGBColor(0x55, 0x55, 0x55)
GRAY     = RGBColor(0x88, 0x88, 0x88)
GRAY_BG  = RGBColor(0xFA, 0xFA, 0xFA)
GRAY_LT  = RGBColor(0xF4, 0xF4, 0xF4)
GREEN    = RGBColor(0x2E, 0x8B, 0x57)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BORDER   = RGBColor(0xE0, 0xE0, 0xE0)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ── helpers ───────────────────────────────────────────────────────────────────
def rect(slide, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line:
        sh.line.color.rgb = line; sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def rrect(slide, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line:
        sh.line.color.rgb = line; sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def oval(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background(); sh.shadow.inherit = False
    return sh

def tb(slide, text, x, y, w, h, size, color=GRAPH, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, italic=False, font="Calibri"):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = font
    return box

def box_filled(slide, text, x, y, w, h, fill, fc=WHITE, size=12, bold=False,
               align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE, rounded=False,
               line_color=None):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line_color:
        sh.line.color.rgb = line_color; sh.line.width = Pt(0.5)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.12)
    tf.margin_top = tf.margin_bottom = Inches(0.06)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = fc; r.font.name = "Calibri"
    return sh

def new_slide(bg=WHITE):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, 13.333, 7.5, bg)
    return s

def header(s, num, title, sub=None):
    """Section header with coral square + title."""
    rrect(s, 0.5, 0.4, 0.7, 0.7, CORAL)
    tb(s, str(num), 0.5, 0.42, 0.7, 0.7, 22, color=WHITE, bold=True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tb(s, title, 1.4, 0.42, 11, 0.42, 22, color=GRAPH, bold=True,
       valign=MSO_ANCHOR.MIDDLE)
    if sub:
        tb(s, sub, 1.4, 0.82, 11, 0.32, 11, color=GRAY_DK,
           valign=MSO_ANCHOR.TOP)
    rect(s, 0.5, 1.25, 12.333, 0.04, CORAL)

def footer(s):
    rect(s, 0, 7.15, 13.333, 0.35, GRAPH)
    rect(s, 0, 7.13, 13.333, 0.03, CORAL)
    tb(s, "AUDITABLE  ·  Soluciones Empresariales", 0.5, 7.21, 6, 0.25,
       9, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    tb(s, "Guía Tributaria 2026  ·  Profesionales de la Salud",
       6.8, 7.21, 6, 0.25, 8.5, color=CORAL_LT, align=PP_ALIGN.RIGHT,
       valign=MSO_ANCHOR.MIDDLE)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — PORTADA
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide(bg=CORAL)

# Decorative
rect(s, 0, 6.95, 13.333, 0.55, GRAPH)

# Brand top
tb(s, "AUDITABLE", 0.5, 0.4, 12, 0.5, 24, color=WHITE, bold=True)
tb(s, "SOLUCIONES EMPRESARIALES", 0.5, 0.9, 12, 0.3, 10, color=CORAL_LT)

rect(s, 0.5, 1.35, 1.5, 0.05, WHITE)

# Main title
tb(s, "Boleta de honorarios", 0.5, 2.2, 12.5, 1.0, 44, color=WHITE, bold=True)
tb(s, "o sociedad:", 0.5, 3.2, 12.5, 1.0, 44, color=WHITE, bold=True)
tb(s, "¿cuál te conviene?", 0.5, 4.2, 12.5, 1.0, 44, color=GRAPH, bold=True)

# Subtitle
tb(s, "Guía práctica para profesionales de la salud",
   0.5, 5.5, 12.5, 0.4, 16, color=WHITE)
tb(s, "que quieren ordenar su tributación y pagar lo justo.",
   0.5, 5.9, 12.5, 0.4, 16, color=WHITE)

# Footer
tb(s, "AUDITABLE  ·  Soluciones Empresariales",
   0.5, 7.22, 6, 0.25, 10, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
tb(s, "Guía Tributaria 2026", 6.8, 7.22, 6, 0.25, 10,
   color=CORAL_LT, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — LAS DOS RUTAS
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 1, "Las dos rutas", "¿Cómo cobras tus servicios profesionales?")

CW, CH = 6.0, 5.55
CY = 1.55

# Card A
ax = 0.7
rrect(s, ax, CY, CW, CH, WHITE, line=BORDER)
rrect(s, ax, CY, CW, 1.0, GRAPH)
tb(s, "RUTA A", ax + 0.3, CY + 0.1, 3, 0.3, 11, color=CORAL_LT, bold=True)
tb(s, "Persona Natural", ax + 0.3, CY + 0.42, 5, 0.5, 20,
   color=WHITE, bold=True)

items_a = [
    "Emite boletas de honorarios electrónicas",
    "Tributa en 2ª Categoría (Art. 42 N°2 LIR)",
    "Retención 15,25% año 2026 (imputable al IGC)",
    "Paga IGC sobre el total anual (hasta 40%)",
    "Deduce gastos reales o presunción 30% (tope 15 UTA)",
    "Exenta de IVA (Art. 12 letra E N°8 DL 825)",
]
ya = CY + 1.25
for it in items_a:
    oval(s, ax + 0.3, ya + 0.18, 0.12, 0.12, CORAL)
    tb(s, it, ax + 0.55, ya, CW - 0.7, 0.5, 11.5, color=GRAPH,
       valign=MSO_ANCHOR.MIDDLE)
    ya += 0.65

# Card B
bx = 7.0
rrect(s, bx, CY, CW, CH, WHITE, line=BORDER)
rrect(s, bx, CY, CW, 1.0, CORAL)
tb(s, "RUTA B", bx + 0.3, CY + 0.1, 3, 0.3, 11, color=WHITE, bold=True)
tb(s, "Sociedad (SpA / EIRL)", bx + 0.3, CY + 0.42, 5.5, 0.5, 20,
   color=WHITE, bold=True)

items_b = [
    "Emite facturas exentas por servicios de salud ambulatorios",
    "Tributa en 1ª Categoría (Art. 20 LIR)",
    "IDPC 12,5% régimen Pro Pyme (ejemplo aplicable)",
    "Socios pagan IGC solo sobre lo que retiran",
    "Deduce gastos reales del giro (sin tope)",
    "Exenta de IVA por prestaciones de salud ambulatoria",
]
yb = CY + 1.25
for it in items_b:
    oval(s, bx + 0.3, yb + 0.18, 0.12, 0.12, CORAL)
    tb(s, it, bx + 0.55, yb, CW - 0.7, 0.5, 11.5, color=GRAPH,
       valign=MSO_ANCHOR.MIDDLE)
    yb += 0.65

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — COMPARATIVA RÁPIDA
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 2, "Comparativa rápida", "Lo esencial de cada estructura en una mirada")

TY = 1.55
TW = 12.333
C1, C2, C3 = 4.8, 3.77, 3.77
TX = 0.5
RH = 0.52

# Header row
rect(s, TX, TY, TW, RH, GRAPH)
tb(s, "CONCEPTO", TX + 0.2, TY, C1, RH, 11, color=WHITE, bold=True,
   valign=MSO_ANCHOR.MIDDLE)
tb(s, "PERSONA NATURAL", TX + C1, TY, C2, RH, 11, color=WHITE, bold=True,
   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
tb(s, "SOCIEDAD (SpA / EIRL)", TX + C1 + C2, TY, C3, RH, 11, color=WHITE,
   bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
TY += RH

rows = [
    ("Impuesto a la renta", "IGC (hasta 40%)", "IDPC 12,5% + IGC sobre retiros"),
    ("¿Cuándo paga?", "100% del ingreso anual", "Solo sobre lo retirado"),
    ("Deducción de gastos", "30% presunto o reales", "Reales sin tope"),
    ("IVA", "Exento", "Exento (salud ambulatoria)"),
    ("Retención mensual", "15,25% (año 2026)", "No aplica"),
    ("Contabilidad", "Simple", "Completa obligatoria"),
    ("Costo de operación", "Bajo", "Medio"),
    ("Planificación tributaria", "Limitada", "Amplia"),
]
for i, (a, b, cc) in enumerate(rows):
    bg = WHITE if i % 2 == 0 else GRAY_BG
    rect(s, TX, TY, TW, RH, bg)
    tb(s, a, TX + 0.2, TY, C1, RH, 10.5, color=GRAPH, bold=True,
       valign=MSO_ANCHOR.MIDDLE)
    tb(s, b, TX + C1, TY, C2, RH, 10.5, color=GRAY_DK,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tb(s, cc, TX + C1 + C2, TY, C3, RH, 10.5, color=GRAY_DK,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    TY += RH

# Border
sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                        Inches(0.5), Inches(1.55), Inches(TW), Inches(TY-1.55))
sh.fill.background()
sh.line.color.rgb = BORDER; sh.line.width = Pt(0.5)
sh.shadow.inherit = False

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — 3 FORMAS DE REMUNERARSE
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 3, "3 formas de remunerarte desde tu sociedad",
       "Cada vía tiene un tratamiento tributario distinto")

forms = [
    ("Retiro de utilidades",
     "Tributas solo cuando retiras. El IDPC pagado por la sociedad se imputa "
     "como crédito al IGC. Permite diferir impuestos reinvirtiendo en la empresa."),
    ("Sueldo patronal",
     "Te pagas un sueldo como administrador. Es gasto para la sociedad y "
     "tributa como renta del trabajo. Sujeto a cotizaciones. Útil para flujo estable."),
    ("Honorarios a tu propia sociedad",
     "Emites boleta a tu sociedad. Tributa en 2ª categoría. Atención: el SII "
     "puede objetar la operación si no existe sustancia económica real."),
]
fy = 1.55
fh = 1.65
for i, (t, d) in enumerate(forms):
    rrect(s, 0.5, fy, 12.333, fh, GRAY_BG, line=BORDER)
    rect(s, 0.5, fy, 0.15, fh, CORAL)
    oval(s, 0.95, fy + 0.45, 0.75, 0.75, CORAL)
    tb(s, str(i+1), 0.95, fy + 0.45, 0.75, 0.75, 22, color=WHITE, bold=True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tb(s, t, 1.95, fy + 0.2, 10.5, 0.45, 17, color=GRAPH, bold=True,
       valign=MSO_ANCHOR.MIDDLE)
    tb(s, d, 1.95, fy + 0.75, 10.3, 0.85, 12, color=GRAY_DK,
       valign=MSO_ANCHOR.TOP)
    fy += fh + 0.15

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — EJEMPLO IMPACTO IGC ($40M)
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 4, "Impacto tributario — ejemplo con $40.000.000 anuales",
       "Comparativo referencial para un profesional de la salud · Año 2026")

# Caso contexto
rrect(s, 0.5, 1.55, 12.333, 0.6, CORAL_LT)
tb(s, "Ingresos profesionales anuales: $40.000.000  ·  Cifras referenciales 2026",
   0.5, 1.55, 12.333, 0.6, 13, color=GRAPH, bold=True,
   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# Dos columnas comparativas
CY = 2.35
CW = 6.0
CH = 4.3

# Columna A — Persona Natural
ax = 0.5
rrect(s, ax, CY, CW, CH, WHITE, line=BORDER)
rect(s, ax, CY, CW, 0.6, GRAPH)
tb(s, "PERSONA NATURAL  ·  Boleta de Honorarios", ax + 0.2, CY, CW - 0.2,
   0.6, 12.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)

rows_a = [
    ("Ingreso bruto anual", "$40.000.000"),
    ("Gastos presuntos (30%)", "($12.000.000)"),
    ("Base imponible", "$28.000.000"),
    ("Retención BHE 15,25% (anual)", "$6.100.000"),
    ("IGC estimado tramo 30-50 UTA", "≈ $1.280.000"),
    ("Devolución probable en F22", "Sí (exceso retención)"),
]
ya = CY + 0.85
for label, val in rows_a:
    tb(s, label, ax + 0.3, ya, CW - 3.0, 0.4, 11, color=GRAY_DK,
       valign=MSO_ANCHOR.MIDDLE)
    tb(s, val, ax + CW - 2.8, ya, 2.6, 0.4, 11, color=GRAPH, bold=True,
       align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)
    rect(s, ax + 0.3, ya + 0.42, CW - 0.6, 0.01, BORDER)
    ya += 0.52

# Resultado A
rrect(s, ax + 0.3, CY + CH - 0.8, CW - 0.6, 0.6, GRAPH)
tb(s, "Carga efectiva ≈ $1.280.000", ax + 0.3, CY + CH - 0.8, CW - 0.6, 0.6,
   13, color=WHITE, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# Columna B — Sociedad
bx = 6.83
rrect(s, bx, CY, CW, CH, WHITE, line=BORDER)
rect(s, bx, CY, CW, 0.6, CORAL)
tb(s, "SOCIEDAD  ·  SpA / EIRL Pro Pyme", bx + 0.2, CY, CW - 0.2,
   0.6, 12.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)

rows_b = [
    ("Ingreso bruto anual (factura exenta)", "$40.000.000"),
    ("Gastos reales del giro (ej. 25%)", "($10.000.000)"),
    ("Base imponible 1ª Categoría", "$30.000.000"),
    ("IDPC 12,5% Pro Pyme", "$3.750.000"),
    ("Utilidad disponible (retirable)", "$26.250.000"),
    ("IGC sobre retiro neto al ejercicio*", "Solo si retira"),
]
yb = CY + 0.85
for label, val in rows_b:
    tb(s, label, bx + 0.3, yb, CW - 3.0, 0.4, 11, color=GRAY_DK,
       valign=MSO_ANCHOR.MIDDLE)
    tb(s, val, bx + CW - 2.8, yb, 2.6, 0.4, 11, color=GRAPH, bold=True,
       align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)
    rect(s, bx + 0.3, yb + 0.42, CW - 0.6, 0.01, BORDER)
    yb += 0.52

# Resultado B
rrect(s, bx + 0.3, CY + CH - 0.8, CW - 0.6, 0.6, CORAL)
tb(s, "IDPC = $3.750.000  ·  IGC diferible", bx + 0.3, CY + CH - 0.8,
   CW - 0.6, 0.6, 13, color=WHITE, bold=True,
   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# Nota
tb(s, "* IDPC pagado por la sociedad se imputa como crédito al IGC del socio al retirar. "
   "Si reinvierte, posterga la tributación personal.",
   0.5, 6.75, 12.333, 0.35, 9.5, color=GRAY, italic=True,
   align=PP_ALIGN.CENTER)

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — CUÁNDO SÍ CONVIENE
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 5, "¿Cuándo SÍ conviene crear sociedad?",
       "Casos en que la estructura genera valor real")

si_items = [
    "Ingresos anuales sobre $40.000.000 – $50.000.000, sostenidos en el tiempo",
    "Tienes gastos reales relevantes: arriendo, secretaria, insumos, equipos",
    "Atiendes a clínicas, ISAPRES o instituciones que prefieren facturas",
    "Quieres reinvertir utilidades en tu consulta (más box, equipamiento)",
    "Buscas separar tu patrimonio personal del riesgo profesional",
    "Quieres ordenar tu flujo con sueldo patronal y retiros planificados",
    "Proyectas crecer, incorporar profesionales o abrir nuevos puntos",
]
y = 1.55
for it in si_items:
    rrect(s, 0.5, y, 12.333, 0.65, GRAY_BG)
    oval(s, 0.7, y + 0.18, 0.3, 0.3, GREEN)
    tb(s, "✓", 0.7, y + 0.16, 0.3, 0.3, 13, color=WHITE, bold=True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tb(s, it, 1.2, y, 11.0, 0.65, 12.5, color=GRAPH,
       valign=MSO_ANCHOR.MIDDLE)
    y += 0.72

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — CUÁNDO NO CONVIENE
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 6, "¿Cuándo NO conviene?",
       "Casos donde la sociedad es innecesaria o riesgosa")

no_items = [
    "Ingresos bajos o variables sin gastos significativos asociados",
    "No tienes capacidad de asumir contabilidad completa ni costos fijos",
    "Pretendes incorporar socios sin sustancia económica real (riesgo NGA)",
    "Tu única motivación es bajar impuestos sin razón económica detrás",
    "Vas a retirar el 100% de las utilidades — no reinvertirás",
    "No quieres asumir las obligaciones formales mensuales y anuales",
]
y = 1.55
for it in no_items:
    rrect(s, 0.5, y, 12.333, 0.75, CORAL_LT)
    oval(s, 0.7, y + 0.22, 0.32, 0.32, CORAL)
    tb(s, "✗", 0.7, y + 0.2, 0.32, 0.32, 14, color=WHITE, bold=True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tb(s, it, 1.2, y, 11.0, 0.75, 12.5, color=GRAPH,
       valign=MSO_ANCHOR.MIDDLE)
    y += 0.82

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — COSTOS Y OBLIGACIONES
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 7, "Costos y obligaciones de la sociedad",
       "Antes de constituir, considera el costo total de operar")

cols = [
    ("Costos iniciales", [
        "Constitución (Empresa en un Día — gratis)",
        "Asesoría tributaria y contable de partida",
        "Inicio de actividades en el SII",
    ]),
    ("Costos recurrentes", [
        "Contabilidad completa mensual",
        "Declaraciones F29 (mensual) y F22 (anual)",
        "Libros, balance y registros de rentas",
        "Patente municipal",
    ]),
    ("Obligaciones formales", [
        "Emisión electrónica de DTE",
        "Mantener registros de rentas empresariales",
        "Cumplir plazos de declaración y pago",
        "Actualizar información societaria",
    ]),
]

CW = 4.0
gap = 0.17
CH = 5.0
CY = 1.65
for i, (t, lst) in enumerate(cols):
    cx = 0.5 + i*(CW + gap)
    rrect(s, cx, CY, CW, CH, WHITE, line=BORDER)
    rect(s, cx, CY, CW, 0.75, CORAL)
    tb(s, t, cx, CY, CW, 0.75, 14, color=WHITE, bold=True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    yy = CY + 1.0
    for item in lst:
        oval(s, cx + 0.3, yy + 0.18, 0.12, 0.12, CORAL)
        tb(s, item, cx + 0.55, yy, CW - 0.75, 0.6, 11, color=GRAY_DK,
           valign=MSO_ANCHOR.MIDDLE)
        yy += 0.7

footer(s)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — RECOMENDACIONES
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, 8, "Recomendaciones prácticas",
       "7 pasos antes de tomar la decisión")

recs = [
    ("Proyecta tus ingresos a 3 años",
     "Antes de constituir, simula el ahorro real esperado."),
    ("Cuantifica tus gastos reales",
     "Identifica qué deducirías efectivamente del giro."),
    ("Evalúa tu modo de trabajo",
     "Solo, con equipo, en clínicas o consulta propia."),
    ("Decide retiro o reinversión",
     "Esto define el régimen tributario óptimo."),
    ("Considera el costo administrativo",
     "Una sociedad implica disciplina contable mensual."),
    ("Escoge el régimen correcto",
     "14 D N°3 si reinvertirás · 14 D N°8 transparente si retirarás todo."),
    ("Asesórate antes de decidir",
     "Cada caso es único — no hay solución universal."),
]

y = 1.55
rh = 0.72
for i, (t, d) in enumerate(recs):
    rrect(s, 0.5, y, 12.333, rh, WHITE, line=BORDER)
    rect(s, 0.5, y, 0.12, rh, CORAL)
    oval(s, 0.85, y + 0.16, 0.4, 0.4, CORAL)
    tb(s, str(i+1), 0.85, y + 0.16, 0.4, 0.4, 13, color=WHITE, bold=True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tb(s, t, 1.5, y + 0.05, 6.5, 0.35, 13, color=GRAPH, bold=True,
       valign=MSO_ANCHOR.MIDDLE)
    tb(s, d, 1.5, y + 0.38, 10.5, 0.32, 10.5, color=GRAY_DK,
       valign=MSO_ANCHOR.TOP)
    y += rh + 0.04

footer(s)


prs.save(OUTPUT)
print(f"✔  PPT generada: {OUTPUT}  —  {len(prs.slides)} slides")
