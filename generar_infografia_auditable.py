#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Infografía Auditable — Boleta vs Sociedad para profesionales de la salud
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

OUTPUT = "/home/user/Prueba/Infografia_Auditable.pdf"

# Paleta Auditable
CORAL    = HexColor("#F25C4F")
CORAL_LT = HexColor("#FCE4E1")
GRAPH    = HexColor("#2C2C2C")
GRAY_DK  = HexColor("#555555")
GRAY     = HexColor("#888888")
GRAY_LT  = HexColor("#F4F4F4")
GRAY_BG  = HexColor("#FAFAFA")
GREEN    = HexColor("#2E8B57")
WHITE    = HexColor("#FFFFFF")
BORDER   = HexColor("#E0E0E0")

# Página vertical larga estilo infografía
PW = 21 * cm
PH = 80 * cm

c = canvas.Canvas(OUTPUT, pagesize=(PW, PH))

# ── helpers ───────────────────────────────────────────────────────────────────
def rect(x, y, w, h, fill, stroke=None, sw=0):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
        c.rect(x, y, w, h, stroke=1, fill=1)
    else:
        c.rect(x, y, w, h, stroke=0, fill=1)

def rrect(x, y, w, h, r, fill, stroke=None, sw=0):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
        c.roundRect(x, y, w, h, r, stroke=1, fill=1)
    else:
        c.roundRect(x, y, w, h, r, stroke=0, fill=1)

def txt(s, x, y, size, color=GRAPH, font="Helvetica", align="left"):
    c.setFillColor(color)
    c.setFont(font, size)
    if align == "center":
        c.drawCentredString(x, y, s)
    elif align == "right":
        c.drawRightString(x, y, s)
    else:
        c.drawString(x, y, s)

def para(s, x, y, w, h, size=10, color=GRAPH, font="Helvetica", align=TA_LEFT, leading=None):
    style = ParagraphStyle(
        "p", fontName=font, fontSize=size, leading=leading or size*1.35,
        textColor=color, alignment=align)
    p = Paragraph(s, style)
    p.wrapOn(c, w, h)
    p.drawOn(c, x, y)

def section_header(y, num, title):
    """Título de sección con número coral y barra."""
    rrect(2*cm, y, 0.8*cm, 0.8*cm, 0.15*cm, CORAL)
    txt(str(num), 2.4*cm, y + 0.22*cm, 14, WHITE, "Helvetica-Bold", "center")
    txt(title.upper(), 3.2*cm, y + 0.28*cm, 14, GRAPH, "Helvetica-Bold")
    rect(2*cm, y - 0.15*cm, 17*cm, 0.04*cm, CORAL)

def bullet(x, y, text, size=10, color=GRAPH, bullet_color=CORAL):
    c.setFillColor(bullet_color)
    c.circle(x + 0.1*cm, y + 0.13*cm, 0.07*cm, stroke=0, fill=1)
    txt(text, x + 0.4*cm, y, size, color)

def check_item(x, y, text, w=15*cm, size=10):
    c.setFillColor(GREEN)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, "✓")
    para(text, x + 0.5*cm, y - 0.1*cm, w, 1*cm, size=size, color=GRAPH)

def cross_item(x, y, text, w=15*cm, size=10):
    c.setFillColor(CORAL)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, "✗")
    para(text, x + 0.5*cm, y - 0.1*cm, w, 1*cm, size=size, color=GRAPH)

# ── HEADER / PORTADA ──────────────────────────────────────────────────────────
Y = PH

# Top coral band
rect(0, Y - 4.5*cm, PW, 4.5*cm, CORAL)

# Brand
txt("AUDITABLE", PW/2, Y - 1.5*cm, 20, WHITE, "Helvetica-Bold", "center")
txt("SOLUCIONES EMPRESARIALES", PW/2, Y - 2.05*cm, 9, WHITE, "Helvetica", "center")

# Decorative line
rect(PW/2 - 1.5*cm, Y - 2.4*cm, 3*cm, 0.04*cm, WHITE)

txt("GUÍA TRIBUTARIA · 2026", PW/2, Y - 3.1*cm, 8.5, CORAL_LT, "Helvetica-Bold", "center")
txt("Profesionales de la Salud", PW/2, Y - 3.85*cm, 13, WHITE, "Helvetica", "center")

Y -= 5.8*cm

# Main title
para('<b>Boleta de honorarios<br/>o sociedad:<br/><font color="#F25C4F">¿cuál te conviene?</font></b>',
     2*cm, Y - 3.8*cm, 17*cm, 4.5*cm, size=24, leading=29, align=TA_CENTER)

Y -= 4.5*cm

para("Guía práctica para profesionales de la salud que quieren ordenar su tributación, "
     "tomar mejores decisiones y pagar lo justo.",
     2.5*cm, Y - 1.5*cm, 16*cm, 2*cm, size=11, color=GRAY_DK, align=TA_CENTER, leading=15)

Y -= 2.5*cm

# Separator
rect(8*cm, Y, 5*cm, 0.08*cm, CORAL)
Y -= 1.2*cm

# ── BLOQUE 2 — LAS DOS RUTAS ──────────────────────────────────────────────────
section_header(Y - 0.5*cm, 1, "Las dos rutas")
Y -= 2*cm

# Two cards side by side
card_w = 8.2*cm
card_h = 9.5*cm
x1 = 2*cm
x2 = 2*cm + card_w + 0.6*cm

# Card A
rrect(x1, Y - card_h, card_w, card_h, 0.3*cm, WHITE, BORDER, 1)
rrect(x1, Y - 1.1*cm, card_w, 1.1*cm, 0.3*cm, GRAPH)
rect(x1, Y - 1.1*cm, card_w, 0.4*cm, GRAPH)
txt("RUTA A", x1 + 0.4*cm, Y - 0.5*cm, 9, CORAL_LT, "Helvetica-Bold")
txt("Persona Natural", x1 + 0.4*cm, Y - 0.95*cm, 13, WHITE, "Helvetica-Bold")

ya = Y - 1.6*cm
items_a = [
    "Emite <b>boletas de honorarios</b> electrónicas",
    "Tributa en <b>2ª Categoría</b> (Art. 42 N°2 LIR)",
    "<b>Retención 15,25%</b> año 2026, imputable al IGC anual",
    "Paga <b>IGC</b> sobre el total anual (hasta 40%)",
    "Deduce gastos efectivos o <b>presunción 30%</b> (tope 15 UTA)",
    "<b>Exenta de IVA</b> (Art. 12 letra E N°8 DL 825)",
]
for it in items_a:
    c.setFillColor(CORAL)
    c.circle(x1 + 0.5*cm, ya - 0.25*cm, 0.08*cm, stroke=0, fill=1)
    para(it, x1 + 0.8*cm, ya - 0.7*cm, card_w - 1*cm, 1.2*cm, size=9.5, leading=12)
    ya -= 1.25*cm

# Card B
rrect(x2, Y - card_h, card_w, card_h, 0.3*cm, WHITE, BORDER, 1)
rrect(x2, Y - 1.1*cm, card_w, 1.1*cm, 0.3*cm, CORAL)
rect(x2, Y - 1.1*cm, card_w, 0.4*cm, CORAL)
txt("RUTA B", x2 + 0.4*cm, Y - 0.5*cm, 9, WHITE, "Helvetica-Bold")
txt("Sociedad (SpA / EIRL)", x2 + 0.4*cm, Y - 0.95*cm, 13, WHITE, "Helvetica-Bold")

yb = Y - 1.6*cm
items_b = [
    "Emite <b>facturas exentas</b> por servicios de salud ambulatorios",
    "Tributa en <b>1ª Categoría</b> (Art. 20 LIR)",
    "<b>IDPC 12,5%</b> régimen Pro Pyme (ejemplo aplicable)",
    "Los socios pagan IGC <b>solo sobre lo que retiran</b>",
    "Deduce <b>gastos reales</b> del giro (sin tope)",
    "<b>Exenta de IVA</b> por prestaciones de salud ambulatorias",
]
for it in items_b:
    c.setFillColor(CORAL)
    c.circle(x2 + 0.5*cm, yb - 0.25*cm, 0.08*cm, stroke=0, fill=1)
    para(it, x2 + 0.8*cm, yb - 0.7*cm, card_w - 1*cm, 1.2*cm, size=9.5, leading=12)
    yb -= 1.25*cm

Y -= card_h + 1.5*cm

# ── BLOQUE 3 — TABLA COMPARATIVA ──────────────────────────────────────────────
section_header(Y, 2, "Comparativa rápida")
Y -= 1.8*cm

# Tabla
tbl_x = 2*cm
tbl_w = 17*cm
col1 = 6.5*cm
col2 = 5.25*cm
col3 = 5.25*cm
rh = 1.0*cm

# Header
rrect(tbl_x, Y - rh, tbl_w, rh, 0.15*cm, GRAPH)
txt("CONCEPTO", tbl_x + 0.3*cm, Y - 0.6*cm, 10, WHITE, "Helvetica-Bold")
txt("PERSONA NATURAL", tbl_x + col1 + col2/2, Y - 0.6*cm, 10, WHITE, "Helvetica-Bold", "center")
txt("SOCIEDAD", tbl_x + col1 + col2 + col3/2, Y - 0.6*cm, 10, WHITE, "Helvetica-Bold", "center")
Y -= rh

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
    rect(tbl_x, Y - rh, tbl_w, rh, bg)
    txt(a, tbl_x + 0.3*cm, Y - 0.6*cm, 9.5, GRAPH, "Helvetica-Bold")
    para(b, tbl_x + col1 + 0.1*cm, Y - 0.85*cm, col2 - 0.2*cm, rh, size=9, align=TA_CENTER, leading=11)
    para(cc, tbl_x + col1 + col2 + 0.1*cm, Y - 0.85*cm, col3 - 0.2*cm, rh, size=9, align=TA_CENTER, leading=11)
    Y -= rh

# table border
c.setStrokeColor(BORDER)
c.setLineWidth(0.5)
c.rect(tbl_x, Y, tbl_w, rh*(len(rows)+1), stroke=1, fill=0)

Y -= 1.5*cm

# ── BLOQUE 4 — 3 FORMAS DE REMUNERARSE ────────────────────────────────────────
section_header(Y, 3, "3 formas de remunerarse desde tu sociedad")
Y -= 1.8*cm

forms = [
    ("Retiro de utilidades",
     "Tributas solo cuando retiras. El IDPC pagado por la sociedad se imputa como "
     "crédito al IGC. Permite <b>diferir impuestos</b> reinvirtiendo en la empresa."),
    ("Sueldo patronal",
     "Te pagas un sueldo como administrador. <b>Es gasto para la sociedad</b> y tributa "
     "como renta del trabajo. Sujeto a cotizaciones. Útil para flujo mensual estable."),
    ("Honorarios a tu propia sociedad",
     "Emites boleta a tu sociedad. Tributa en 2ª categoría. <b>Atención:</b> el SII puede "
     "objetar la operación si no existe sustancia económica real."),
]
for i, (t, d) in enumerate(forms):
    card_h_f = 3.0*cm
    rrect(2*cm, Y - card_h_f, 17*cm, card_h_f, 0.2*cm, GRAY_BG, BORDER, 0.5)
    rrect(2*cm, Y - card_h_f, 0.15*cm, card_h_f, 0.05*cm, CORAL)
    # Numerito
    c.setFillColor(CORAL)
    c.circle(2.8*cm, Y - 0.7*cm, 0.4*cm, stroke=0, fill=1)
    txt(str(i+1), 2.8*cm, Y - 0.85*cm, 13, WHITE, "Helvetica-Bold", "center")
    txt(t, 3.6*cm, Y - 0.8*cm, 12, GRAPH, "Helvetica-Bold")
    para(d, 3.6*cm, Y - 2.6*cm, 15*cm, 2*cm, size=10, color=GRAY_DK, leading=13)
    Y -= card_h_f + 0.3*cm

Y -= 0.8*cm

# ── BLOQUE 5 — IMPACTO IGC ────────────────────────────────────────────────────
section_header(Y, 4, "Impacto en el IGC — ejemplo ilustrativo")
Y -= 1.8*cm

# Caja contexto
rrect(2*cm, Y - 1*cm, 17*cm, 1*cm, 0.15*cm, CORAL_LT)
para("<b>Profesional con ingresos anuales por $120.000.000</b> — ejemplo referencial.",
     2.3*cm, Y - 0.75*cm, 16.5*cm, 1*cm, size=10, color=GRAPH, align=TA_CENTER)
Y -= 1.3*cm

# Tabla ejemplo
ex_rh = 0.85*cm
ex_rows = [
    ("Ingreso bruto anual", "$120.000.000", "$120.000.000"),
    ("Gastos deducibles", "Presunción 30%", "Reales (ej. 25%)"),
    ("Base imponible", "≈ $84.000.000", "≈ $90.000.000"),
    ("IDPC 12,5%", "—", "≈ $11.250.000"),
    ("Retiro estimado", "—", "≈ $40.000.000"),
    ("IGC estimado", "Alto (tramo superior)", "Moderado (sobre retiro)"),
]
# Header
rrect(2*cm, Y - ex_rh, 17*cm, ex_rh, 0.1*cm, GRAPH)
txt("CONCEPTO", 2.3*cm, Y - 0.55*cm, 9.5, WHITE, "Helvetica-Bold")
txt("BOLETA HONORARIOS", 2*cm + col1 + col2/2, Y - 0.55*cm, 9.5, WHITE, "Helvetica-Bold", "center")
txt("SOCIEDAD PRO PYME", 2*cm + col1 + col2 + col3/2, Y - 0.55*cm, 9.5, WHITE, "Helvetica-Bold", "center")
Y -= ex_rh
for i, (a, b, cc) in enumerate(ex_rows):
    bg = WHITE if i % 2 == 0 else GRAY_BG
    rect(2*cm, Y - ex_rh, 17*cm, ex_rh, bg)
    txt(a, 2.3*cm, Y - 0.55*cm, 9, GRAPH, "Helvetica-Bold")
    txt(b, 2*cm + col1 + col2/2, Y - 0.55*cm, 9, GRAY_DK, "Helvetica", "center")
    txt(cc, 2*cm + col1 + col2 + col3/2, Y - 0.55*cm, 9, GRAY_DK, "Helvetica", "center")
    Y -= ex_rh

# Resultado destacado
rrect(2*cm, Y - 1.3*cm, 17*cm, 1.3*cm, 0.15*cm, CORAL)
para("<b>Carga total estimada:</b> mayor en boleta · menor en sociedad <b>con buena planificación</b>",
     2.3*cm, Y - 0.95*cm, 16.5*cm, 1.3*cm, size=10.5, color=WHITE, align=TA_CENTER)
Y -= 1.6*cm

txt("Cifras referenciales. El resultado real depende de gastos, retiros y régimen escogido.",
    PW/2, Y, 8, GRAY, "Helvetica-Oblique", "center")
Y -= 1.5*cm

# ── BLOQUE 6 — CUÁNDO SÍ ──────────────────────────────────────────────────────
section_header(Y, 5, "¿Cuándo SÍ conviene crear sociedad?")
Y -= 1.5*cm

si_items = [
    "<b>Ingresos anuales sobre $40.000.000 – $50.000.000</b> aproximados y sostenidos",
    "Tienes <b>gastos reales relevantes</b>: arriendo, secretaria, insumos, equipos",
    "Atiendes a <b>clínicas, ISAPRES o instituciones</b> que prefieren facturas",
    "Quieres <b>reinvertir utilidades</b> en tu consulta (más box, equipamiento)",
    "Buscas <b>separar tu patrimonio personal</b> del riesgo profesional",
    "Quieres <b>ordenar tu flujo</b> con sueldo patronal y retiros planificados",
    "Proyectas <b>crecer</b>, incorporar profesionales o abrir nuevos puntos",
]
for it in si_items:
    rrect(2*cm, Y - 0.95*cm, 17*cm, 0.85*cm, 0.1*cm, GRAY_BG)
    c.setFillColor(GREEN)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(2.3*cm, Y - 0.6*cm, "✓")
    para(it, 2.9*cm, Y - 0.78*cm, 15.8*cm, 0.85*cm, size=10, leading=12)
    Y -= 1.0*cm

Y -= 0.8*cm

# ── BLOQUE 7 — CUÁNDO NO ──────────────────────────────────────────────────────
section_header(Y, 6, "¿Cuándo NO conviene?")
Y -= 1.5*cm

no_items = [
    "<b>Ingresos bajos o variables</b> sin gastos significativos asociados",
    "No tienes capacidad de asumir <b>contabilidad completa</b> ni costos fijos",
    "Pretendes incorporar socios <b>sin sustancia económica real</b> (riesgo NGA)",
    "Tu única motivación es <b>bajar impuestos</b> sin razón económica detrás",
    "Vas a <b>retirar el 100%</b> de las utilidades — no reinvertirás",
    "No quieres asumir las <b>obligaciones formales</b> mensuales y anuales",
]
for it in no_items:
    rrect(2*cm, Y - 0.95*cm, 17*cm, 0.85*cm, 0.1*cm, CORAL_LT)
    c.setFillColor(CORAL)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(2.3*cm, Y - 0.6*cm, "✗")
    para(it, 2.9*cm, Y - 0.78*cm, 15.8*cm, 0.85*cm, size=10, leading=12)
    Y -= 1.0*cm

Y -= 0.8*cm

# ── BLOQUE 8 — COSTOS Y OBLIGACIONES ──────────────────────────────────────────
section_header(Y, 7, "Costos y obligaciones de la sociedad")
Y -= 1.8*cm

# 3 columnas
ccw = 5.4*cm
gap = 0.4*cm
cy = Y
col_h = 6.5*cm

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
for i, (t, lst) in enumerate(cols):
    cx = 2*cm + i*(ccw + gap)
    rrect(cx, cy - col_h, ccw, col_h, 0.2*cm, WHITE, BORDER, 0.7)
    rrect(cx, cy - 1*cm, ccw, 1*cm, 0.2*cm, CORAL)
    rect(cx, cy - 1*cm, ccw, 0.3*cm, CORAL)
    para(f"<b>{t}</b>", cx + 0.3*cm, cy - 0.85*cm, ccw - 0.5*cm, 1*cm,
         size=10.5, color=WHITE, align=TA_CENTER)
    yy = cy - 1.4*cm
    for item in lst:
        c.setFillColor(CORAL)
        c.circle(cx + 0.4*cm, yy - 0.1*cm, 0.06*cm, stroke=0, fill=1)
        para(item, cx + 0.65*cm, yy - 0.7*cm, ccw - 0.9*cm, 1.1*cm,
             size=8.5, color=GRAY_DK, leading=11)
        yy -= 1.05*cm

Y -= col_h + 1.5*cm

# ── BLOQUE 9 — RECOMENDACIONES ────────────────────────────────────────────────
section_header(Y, 8, "Recomendaciones prácticas")
Y -= 1.8*cm

recs = [
    ("Proyecta tus ingresos a 3 años", "Antes de constituir, simula el ahorro real."),
    ("Cuantifica tus gastos reales", "Identifica qué deducirías efectivamente."),
    ("Evalúa tu modo de trabajo", "Solo, con equipo, en clínicas o consulta propia."),
    ("Decide retiro o reinversión", "Esto define el régimen tributario óptimo."),
    ("Considera el costo administrativo", "Una sociedad implica disciplina contable."),
    ("Escoge el régimen correcto",
     "<b>14 D N°3</b> si reinvertirás · <b>14 D N°8</b> (transparente) si retirarás todo."),
    ("Asesórate antes de decidir", "Cada caso es único — no hay solución universal."),
]
for i, (t, d) in enumerate(recs):
    rrect(2*cm, Y - 1.4*cm, 17*cm, 1.3*cm, 0.15*cm, WHITE, BORDER, 0.5)
    rrect(2*cm, Y - 1.4*cm, 0.15*cm, 1.3*cm, 0.05*cm, CORAL)
    c.setFillColor(CORAL)
    c.circle(2.9*cm, Y - 0.7*cm, 0.35*cm, stroke=0, fill=1)
    txt(str(i+1), 2.9*cm, Y - 0.83*cm, 11, WHITE, "Helvetica-Bold", "center")
    txt(t, 3.6*cm, Y - 0.55*cm, 11, GRAPH, "Helvetica-Bold")
    para(d, 3.6*cm, Y - 1.25*cm, 15*cm, 1*cm, size=9.5, color=GRAY_DK, leading=11)
    Y -= 1.5*cm

Y -= 1.2*cm

# ── FOOTER ────────────────────────────────────────────────────────────────────
rect(0, 0, PW, 3.5*cm, GRAPH)
rect(0, 3.5*cm, PW, 0.1*cm, CORAL)

txt("AUDITABLE", PW/2, 2.5*cm, 16, WHITE, "Helvetica-Bold", "center")
txt("SOLUCIONES EMPRESARIALES", PW/2, 2.05*cm, 8, CORAL_LT, "Helvetica", "center")

para("Contenido informativo. La decisión tributaria depende del análisis particular de cada caso. "
     "Cifras y tasas referenciales conforme a normativa vigente al año 2026.",
     2*cm, 0.7*cm, 17*cm, 1.2*cm, size=7.5, color=GRAY, align=TA_CENTER, leading=10)

c.showPage()
c.save()
print(f"✔  Infografía generada: {OUTPUT}")
