#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Caso 86 — Catálogo de Esquemas Tributarios SII 2025
PPT Rediseño moderno — mismo contenido, diseño completamente nuevo
Grupo 4: Andrea Añasco, Gema Sepúlveda, Karen Rebolledo, Nicolás Muñoz
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

OUTPUT = "/home/user/Prueba/Caso86_Presentacion.pptx"

# ── Paleta moderna ────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1A, 0x3A, 0x5C)
NAVY_LT = RGBColor(0x25, 0x4E, 0x80)
ORANGE  = RGBColor(0xE8, 0x50, 0x0A)
BLUE    = RGBColor(0x2E, 0x7B, 0xC4)
TEAL    = RGBColor(0x00, 0x87, 0x97)
GREEN   = RGBColor(0x1A, 0x7A, 0x3C)
RED     = RGBColor(0xC0, 0x28, 0x28)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE= RGBColor(0xF5, 0xF7, 0xFA)
CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)
BORDER  = RGBColor(0xD0, 0xD8, 0xE8)
TEXT_D  = RGBColor(0x1C, 0x1C, 0x1C)
TEXT_M  = RGBColor(0x55, 0x55, 0x55)
TEXT_L  = RGBColor(0xAA, 0xBB, 0xCC)
GOLD    = RGBColor(0xF5, 0xC5, 0x18)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height
BLANK = prs.slide_layouts[6]

# ── Core helpers ─────────────────────────────────────────────────────────────

def rect(slide, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def rrect(slide, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def oval(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def tb(slide, text, x, y, w, h, size, color=TEXT_D, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, font="Calibri",
       italic=False, line_sp=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf  = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    if line_sp:
        from pptx.util import Pt as Pt2
        from pptx.oxml.ns import qn
        from pptx.oxml import OxmlElement
        pPr = p._p.get_or_add_pPr()
        lnSpc = OxmlElement('a:lnSpc')
        spcPts = OxmlElement('a:spcPts')
        spcPts.set(qn('val'), str(int(line_sp * 100)))
        lnSpc.append(spcPts)
        pPr.append(lnSpc)
    r = p.add_run()
    r.text = text
    r.font.size  = Pt(size)
    r.font.bold  = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name  = font
    return box

def box(slide, text, x, y, w, h, fill, fc=WHITE, size=11, bold=False,
        align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE,
        font="Calibri", line_color=None, multiline=False, lpad=0.12):
    """Filled rectangle with text."""
    sh = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line_color:
        sh.line.color.rgb = line_color
        sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap  = True
    tf.margin_left   = Inches(lpad)
    tf.margin_right  = Inches(lpad)
    tf.margin_top    = Inches(0.07)
    tf.margin_bottom = Inches(0.07)
    tf.vertical_anchor = valign
    lines = text.split('\n') if multiline else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size  = Pt(size)
        r.font.bold  = bold
        r.font.color.rgb = fc
        r.font.name  = font
    return sh

def rbox(slide, text, x, y, w, h, fill, fc=WHITE, size=11, bold=False,
         align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE, font="Calibri",
         line_color=None, multiline=False):
    """Rounded rectangle with text."""
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line_color:
        sh.line.color.rgb = line_color
        sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.15)
    tf.margin_top  = tf.margin_bottom = Inches(0.08)
    tf.vertical_anchor = valign
    lines = text.split('\n') if multiline else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = fc
        r.font.name = font
    return sh

def circle_label(slide, letter, cx, cy, r=0.28, fill=NAVY, fc=WHITE, size=16):
    d = r * 2
    sh = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(cx - r), Inches(cy - r), Inches(d), Inches(d))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r_run = p.add_run()
    r_run.text = letter
    r_run.font.size = Pt(size)
    r_run.font.bold = True
    r_run.font.color.rgb = fc
    r_run.font.name = "Calibri"
    return sh

def send_back(shape):
    el = shape._element
    sp = el.getparent()
    sp.remove(el)
    sp.insert(2, el)

def header(slide, title, sub=None):
    """Standard modern header — navy band + orange accent line."""
    bg = rect(slide, 0, 0, 13.333, 1.45, NAVY)
    send_back(bg)
    rect(slide, 0, 1.45, 13.333, 0.06, ORANGE)
    tb(slide, title, 0.55, 0.12, 10.5, 1.0, 28,
       color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    if sub:
        tb(slide, sub, 0.55, 0.95, 10.5, 0.45, 13,
           color=TEXT_L, valign=MSO_ANCHOR.TOP)

def footer(slide):
    """Bottom navy bar with group info."""
    rect(slide, 0, 7.33, 13.333, 0.17, NAVY)
    tb(slide, "Grupo 4: Andrea Añasco · Gema Sepúlveda · Karen Rebolledo · Nicolás Muñoz   |   "
       "Caso 86 — Catálogo de Esquemas Tributarios SII 2025",
       0.4, 7.34, 12.5, 0.14, 8, color=TEXT_L, align=PP_ALIGN.RIGHT)

def new_slide(bg=OFFWHITE, do_footer=True):
    s = prs.slides.add_slide(BLANK)
    bg_sh = rect(s, 0, 0, 13.333, 7.5, bg)
    send_back(bg_sh)
    if do_footer:
        footer(s)
    return s

def divider(slide, x, y, w, color=ORANGE, h=0.035):
    rect(slide, x, y, w, h, color)

def person(slide, letter, cx, cy, color=NAVY):
    """Geometric person icon: circle + body."""
    HR = 0.20
    BW, BH = 0.28, 0.35
    oval(slide, cx - HR, cy,      HR*2, HR*2, color)
    rrect(slide, cx - BW/2, cy + HR*2 + 0.03, BW, BH, color)
    tb(slide, letter, cx - 0.22, cy - 0.05, 0.44, 0.42,
       20, color=color, bold=True, align=PP_ALIGN.CENTER,
       valign=MSO_ANCHOR.MIDDLE)

def arrow_h(slide, x, y, w=0.5, color=TEXT_M):
    """Simple horizontal arrow label."""
    tb(slide, "→", x, y, w, 0.35, 20, color=color,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

def arrow_dn(slide, x, y, h=0.4, color=TEXT_M):
    tb(slide, "↓", x, y, 0.35, h, 20, color=color,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — PORTADA
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide(bg=NAVY, do_footer=False)

# Accent top line
rect(s, 0, 0, 13.333, 0.12, ORANGE)

# Decorative right stripe
rect(s, 12.0, 0, 1.333, 7.5, NAVY_LT)
rect(s, 12.0, 0, 0.06, 7.5, ORANGE)

# Institution label
tb(s, "MAGÍSTER EN DIRECCIÓN TRIBUTARIA  —  UNIVERSIDAD VIÑA DEL MAR",
   0.6, 0.25, 10.5, 0.4, 9.5, color=TEXT_L, bold=False,
   align=PP_ALIGN.LEFT)

# Main title
tb(s, "Análisis Catálogo de\nEsquemas Tributarios",
   0.6, 0.8, 9.0, 2.2, 40, color=WHITE, bold=True,
   align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

# Subtitle orange
rect(s, 0.6, 3.15, 0.1, 0.8, ORANGE)
tb(s, "Caso 86", 0.85, 3.15, 5.0, 0.85, 34, color=ORANGE, bold=True,
   align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)

divider(s, 0.6, 4.1, 8.5, color=RGBColor(0x40, 0x60, 0x90))

# Members
tb(s, "Integrantes", 0.6, 4.3, 3.0, 0.4, 11, color=TEXT_L)
miembros = ["Andrea Añasco", "Gema Sepúlveda", "Karen Rebolledo", "Nicolás Muñoz"]
x_m = 0.6
for m in miembros:
    rbox(s, m, x_m, 4.75, 2.9, 0.55, NAVY_LT, fc=WHITE, size=12, bold=True,
         align=PP_ALIGN.CENTER)
    x_m += 3.05

# Bottom label
tb(s, "Catálogo de Esquemas Tributarios SII 2025",
   0.6, 5.55, 8.0, 0.4, 11, color=TEXT_L, italic=True)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — CASO 86: ¿QUÉ ES?
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Caso 86")

# Left badge
rect(s, 0.5, 1.65, 0.08, 5.55, ORANGE)
box(s, "CASO\n86", 0.65, 2.1, 1.5, 1.5,
    NAVY, fc=WHITE, size=30, bold=True, align=PP_ALIGN.CENTER, multiline=True)

# Description card
rbox(s, "Prestación de servicios profesionales a través de una sociedad de "
     "profesionales en la que los servicios son prestados por uno de sus socios "
     "(A), en la que participan otros socios (B y C) que no prestan sus servicios, "
     "ni tienen algún otro rol estratégico dentro de la sociedad.",
     2.4, 1.75, 10.4, 1.7, CARD_BG, fc=TEXT_D, size=13.5,
     valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

divider(s, 2.4, 3.6, 10.4)

# Key elements in 3 cards
datos = [
    ("Socios", "A, B y C — personas naturales\n33,33% de participación c/u", NAVY),
    ("Sociedad D", "Prestación de servicios profesionales\nCapital no significativo vs. ingresos", BLUE),
    ("Problema", "Solo A presta los servicios\nB y C no tienen rol activo ni aportan valor", ORANGE),
]
x_d = 2.4
for titulo, txt, color in datos:
    rect(s, x_d, 3.75, 3.3, 0.42, color)
    tb(s, titulo, x_d + 0.15, 3.78, 3.0, 0.38,
       13, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    rbox(s, txt, x_d, 4.22, 3.3, 1.1, CARD_BG, fc=TEXT_D, size=11.5,
         valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    x_d += 3.45

# Result box
box(s, "RESULTADO: El esquema afecta indebidamente la carga tributaria del IGC de A "
    "por los ingresos obtenidos mediante la sociedad interpuesta.",
    2.4, 5.55, 10.4, 0.65, NAVY, fc=WHITE, size=12, bold=False,
    align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — ÍNDICE
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Índice")

items_idx = [
    ("1", "Normativa base",              "Art. 42 N°2 LIR · Circ. N°21/1991 y N°50/2020 SII", NAVY),
    ("2", "Descripción del esquema",     "Planificación tributaria — interposición sociedad de profesionales", BLUE),
    ("3", "Puntos de análisis del SII",  "Criterios de fiscalización y efectos del esquema", TEAL),
    ("4", "Calificación jurídica",       "¿Abuso (Art. 4 ter CT) o Simulación (Art. 4 quáter CT)?", ORANGE),
    ("5", "Alternativas de solución",    "5 estructuras legítimas compatibles con las NGA", GREEN),
    ("6", "Conclusiones",                "Síntesis del análisis — Posición del Grupo 4", RED),
]

y_i = 1.7
for num, title, sub, color in items_idx:
    circle_label(s, num, 0.97, y_i + 0.28, r=0.25, fill=color, fc=WHITE, size=14)
    rect(s, 1.4, y_i, 0.06, 0.56, color)
    tb(s, title, 1.6, y_i, 6.0, 0.30, 13.5, color=TEXT_D, bold=True,
       valign=MSO_ANCHOR.BOTTOM)
    tb(s, sub, 1.6, y_i + 0.28, 10.8, 0.28, 10.5, color=TEXT_M,
       valign=MSO_ANCHOR.TOP)
    divider(s, 1.4, y_i + 0.62, 11.4, color=BORDER)
    y_i += 0.76


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — NORMATIVA BASE + MALLA SOCIETARIA
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Normativa Base", "Art. 42 N°2 LIR  ·  Circular N°21 de 1991  ·  Circular N°50 de 2020")

# Left column — normativa (width 7.4)
LCX, LCW = 0.5, 7.3

rect(s, LCX, 1.65, LCW, 0.42, NAVY)
tb(s, "Artículo 42 N°2 — Ley sobre Impuesto a la Renta",
   LCX + 0.15, 1.66, LCW - 0.2, 0.40,
   11.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)

rbox(s, "Ingresos del ejercicio de profesiones liberales u otras ocupaciones lucrativas "
     "(2ª categoría). Las sociedades de profesionales que presten exclusivamente servicios "
     "o asesorías profesionales pueden optar por declarar en 1ª categoría.",
     LCX, 2.12, LCW, 0.85, CARD_BG, fc=TEXT_D, size=11,
     valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

rect(s, LCX, 3.08, LCW, 0.42, BLUE)
tb(s, "Circular N°21/1991 y Circular N°50/2020  —  Requisitos copulativos",
   LCX + 0.15, 3.09, LCW - 0.2, 0.40,
   11.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)

reqs = [
    "Debe tratarse de una sociedad de personas.",
    "Objeto exclusivo: prestación de servicios o asesorías profesionales.",
    "Servicios prestados por intermedio de sus socios (o dependientes que coadyuven).",
    "TODOS los socios deben ejercer su profesión para la sociedad.",
    "No es aceptable que un socio solo aporte capital sin ejercer labor profesional.",
    "Lo relevante es la realización efectiva de labores, no la forma de remuneración.",
]
y_r = 3.57
for req in reqs:
    circle_label(s, "·", LCX + 0.22, y_r + 0.13, r=0.08, fill=BLUE, fc=WHITE, size=9)
    tb(s, req, LCX + 0.42, y_r, LCW - 0.5, 0.30,
       10.5, color=TEXT_D, valign=MSO_ANCHOR.MIDDLE)
    y_r += 0.31

# Right column — Malla Societaria
RCX = 8.1
tb(s, "Malla Societaria", RCX, 1.65, 5.0, 0.42,
   15, color=NAVY, bold=True, valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# Persons B, A, C
positions_p4 = [("B", RCX + 0.3), ("A", RCX + 1.8), ("C", RCX + 3.3)]
for letter, cx in positions_p4:
    color_p = NAVY if letter == "A" else BLUE
    person(s, letter, cx, 2.15, color=color_p)

# Lines (connecting to society below)
for cx_arrow in [RCX + 0.6, RCX + 2.1, RCX + 3.6]:
    arrow_dn(s, cx_arrow - 0.18, 3.1, 0.5, TEXT_M)

# Percentages
pct_x = [RCX + 0.1, RCX + 1.62, RCX + 3.1]
for px in pct_x:
    tb(s, "33,33%", px, 2.85, 0.85, 0.28, 9.5,
       color=TEXT_M, align=PP_ALIGN.CENTER)

# Society D
box(s, "Sociedad de Profesionales D", RCX + 0.3, 3.7, 4.1, 0.65,
    NAVY, fc=WHITE, size=12.5, bold=True, align=PP_ALIGN.CENTER)

# Note
rbox(s, "Solo A presta los servicios profesionales.\nB y C no tienen rol activo.",
     RCX + 0.3, 4.55, 4.1, 0.75, RGBColor(0xFB, 0xE5, 0xE5),
     fc=RED, size=11, bold=False, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — DESCRIPCIÓN DEL ESQUEMA — PLANIFICACIÓN
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Descripción del Esquema — Planificación")

# === LEFT PANEL: A, B, C → PN ===
rect(s, 0.45, 1.65, 2.65, 5.55, RGBColor(0xE8, 0xEF, 0xF8))
rect(s, 0.45, 1.65, 0.07, 5.55, NAVY)
tb(s, "Socios", 0.6, 1.72, 2.35, 0.38,
   12, color=NAVY, bold=True, valign=MSO_ANCHOR.MIDDLE)

for letter, y_p, color_p in [("A", 2.2, ORANGE), ("B", 3.35, NAVY), ("C", 4.5, NAVY)]:
    circle_label(s, letter, 1.15, y_p + 0.25, r=0.30, fill=color_p, fc=WHITE, size=16)
    arrow_h(s, 1.6, y_p + 0.12, 0.5, TEXT_M)
    rbox(s, "Persona\nNatural", 2.15, y_p, 0.85, 0.60, CARD_BG,
         fc=TEXT_D, size=9.5, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# === CENTER: description + flow ===
rect(s, 3.25, 1.65, 5.7, 0.42, NAVY)
tb(s, "El esquema", 3.4, 1.66, 5.5, 0.40,
   12.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)

rbox(s, '"A", "B" y "C" constituyen una sociedad de profesionales "D", destinada a la '
     'prestación de servicios. La participación de los socios se divide en partes iguales, '
     'tanto en capital como en utilidades. El capital social no es significativo en '
     'comparación con los ingresos anuales. Es importante señalar que A es el socio que '
     'presta los servicios profesionales de la compañía.',
     3.25, 2.12, 5.7, 1.75, CARD_BG, fc=TEXT_D, size=11,
     valign=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)

# Flow diagram B A C → D
for i, (letter, cx2, color_f) in enumerate([("B", 3.9, NAVY), ("A", 4.85, ORANGE), ("C", 5.8, NAVY)]):
    circle_label(s, letter, cx2, 4.15, r=0.28, fill=color_f, fc=WHITE, size=15)
    tb(s, "33,33%", cx2 - 0.3, 3.85, 0.75, 0.27, 9, color=TEXT_M, align=PP_ALIGN.CENTER)

# Converging arrows
for cx3 in [3.9, 4.85, 5.8]:
    arrow_dn(s, cx3 - 0.18, 4.55, 0.5, TEXT_M)

box(s, "Sociedad D", 3.7, 5.15, 2.5, 0.55,
    NAVY, fc=WHITE, size=13, bold=True, align=PP_ALIGN.CENTER)

# === RIGHT PANEL: RESULTADO ===
rect(s, 9.15, 1.65, 4.0, 5.55, RGBColor(0xFF, 0xF0, 0xE5))
rect(s, 9.15, 1.65, 0.07, 5.55, ORANGE)

tb(s, "RESULTADO", 9.3, 1.72, 3.7, 0.40,
   14, color=ORANGE, bold=True, valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
divider(s, 9.3, 2.2, 3.65, color=ORANGE)

tb(s, "El esquema afecta indebidamente la carga tributaria del Impuesto Global "
   "Complementario de 'A' por los ingresos obtenidos mediante la interposición "
   "de una sociedad de profesionales, en la que participan otros socios que "
   "no prestan sus servicios, ni tienen algún otro rol estratégico dentro de la sociedad.",
   9.3, 2.35, 3.7, 4.65,
   11.5, color=TEXT_D, valign=MSO_ANCHOR.TOP)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — PUNTOS DE ANÁLISIS DEL SII
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Puntos de Análisis del SII")

# Top wide card
box(s, "¿Existen razones económicas o jurídicas relevantes distintas a las meramente tributarias "
    "para que los servicios que realiza A sean prestados a través de la sociedad D?",
    0.45, 1.62, 8.3, 0.72, NAVY, fc=WHITE, size=11.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)

# Grid of analysis points (2 columns × 5 rows)
puntos = [
    "Si existe vínculo entre los socios",
    "El origen y naturaleza de los ingresos de la sociedad",
    "Qué socios prestan los servicios y sus capacidades técnicas",
    "El comportamiento histórico de impuestos finales de los intervinientes",
    "Seguimiento de flujos de las utilidades distribuidas",
    "Rol de socios que no prestan servicios (estrategia, administración, imagen, flujos)",
    "Política de retiro de utilidades",
    "Si los socios reciben otras remuneraciones o ingresos desde la sociedad",
    "Si se han solicitado devoluciones de impuestos por parte de los socios",
]
y_g = 2.46
col_w = 4.0
for i, p_txt in enumerate(puntos):
    col = i % 2
    row = i // 2
    xp = 0.45 + col * (col_w + 0.2)
    yp = y_g + row * 0.78
    circle_label(s, str(i+1), xp + 0.27, yp + 0.22, r=0.2, fill=BLUE, fc=WHITE, size=11)
    rbox(s, p_txt, xp + 0.6, yp, col_w - 0.65, 0.50,
         CARD_BG, fc=TEXT_D, size=10.5, line_color=BORDER,
         valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

# Right panel — Effects + Faculties + Normativa
rect(s, 9.0, 1.62, 4.2, 5.6, RGBColor(0xF0, 0xF4, 0xFA))
rect(s, 9.0, 1.62, 0.07, 5.6, NAVY)

secciones = [
    ("Efectos del Esquema", NAVY,
     "Disminuir la base imponible del IGC que le correspondería al socio que "
     "presta mayoritariamente los servicios (A), diluyendo el ingreso en 3 RUT."),
    ("Facultades del SII", BLUE,
     "El SII podría evaluar la aplicación de las NGA si se establece que la "
     "operación tuvo el efecto de afectar la base imponible de IGC de A."),
    ("Normativa aplicable", TEAL,
     "Art. 42 N°2, 43 N°2 y 52 LIR  ·  Art. 14, 20 LIR\n"
     "Art. 2 N°2 LIVS  ·  Arts. 4 bis, 4 ter y 4 quáter CT"),
]
y_s = 1.72
for t_s, color_s, txt_s in secciones:
    rect(s, 9.15, y_s, 3.9, 0.38, color_s)
    tb(s, t_s, 9.3, y_s + 0.02, 3.7, 0.36, 11.5,
       color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    rbox(s, txt_s, 9.15, y_s + 0.42, 3.9, 1.0, CARD_BG,
         fc=TEXT_D, size=10.5, valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    y_s += 1.58


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — ALTERNATIVAS DE SOLUCIÓN
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Alternativas de Solución",
       "Estructuras legítimas compatibles con el ordenamiento tributario")

alts = [
    ("1", "Profesional independiente — Art. 42 N°2 LIR",
     "A ejerce como persona natural, emite boletas de honorarios. Puede deducir gastos "
     "efectivos o presuntos (30%, tope 15 UTA). Paga IGC sobre el total. "
     "Opción más simple, sin exposición a las NGA.",
     NAVY),
    ("2", "Sociedad de profesionales con socios que efectivamente trabajen",
     "A se asocia con profesionales que coadyuven a la prestación de servicios. "
     "D califica legítimamente bajo Art. 42 N°2 inc. 3° LIR y puede optar "
     "por tributar en 1ª o 2ª categoría.",
     BLUE),
    ("3", "Distribución de utilidades proporcional al trabajo real",
     "Estatutos con participación diferenciada (ej. A 90%, B 5%, C 5%) "
     "o sueldo patronal a A previo al reparto. La distribución debe reflejar "
     "el aporte económico real de cada socio.",
     TEAL),
    ("4", "Sociedad por Acciones (SpA) — régimen Pro Pyme (Art. 14 D N°3 LIR)",
     "SpA con IDPC al 25%. Permite planificación de retiros y reinversión. "
     "Sin las restricciones de la sociedad de profesionales, "
     "pero exige sustancia económica real.",
     GREEN),
    ("5", "Régimen Pro Pyme Transparente — Art. 14 D N°8 LIR",
     "La empresa no paga IDPC; socios tributan directamente con IGC. "
     "Válido solo si los socios tienen sustancia económica real — "
     "si no, el SII puede igualmente aplicar las NGA.",
     ORANGE),
]

y_a = 1.65
h_a = 1.04
for num, tit, desc, color in alts:
    circle_label(s, num, 0.78, y_a + h_a/2, r=0.28, fill=color, fc=WHITE, size=15)
    rect(s, 1.2, y_a, 0.07, h_a, color)
    rect(s, 1.32, y_a, 11.55, 0.38, color)
    tb(s, tit, 1.5, y_a + 0.01, 11.3, 0.38,
       12.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    rbox(s, desc, 1.32, y_a + 0.40, 11.55, h_a - 0.42,
         CARD_BG, fc=TEXT_D, size=11.5, line_color=BORDER,
         valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    y_a += h_a + 0.04

# Note
box(s, "Toda alternativa debe poder acreditar SUSTANCIA ECONÓMICA REAL. "
    "La sola forma jurídica no es suficiente frente a las NGA.",
    1.2, 6.95, 11.67, 0.30, NAVY, fc=WHITE, size=10.5,
    align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — CONCLUSIONES
# ═════════════════════════════════════════════════════════════════════════════
s = new_slide()
header(s, "Conclusiones", "Caso 86 — Síntesis del análisis — Posición del Grupo 4")

concls = [
    ("1", "ABUSO de formas jurídicas — Art. 4 ter CT",
     "El Caso 86 configura abuso, no simulación. Los actos son reales "
     "pero la forma jurídica es inapropiada: el único efecto relevante es tributario.",
     RED),
    ("2", "La sociedad D no califica como sociedad de profesionales",
     "B y C no ejercen su profesión para D, incumpliendo el requisito copulativo "
     "de las Circulares N°21/1991 y N°50/2020 del SII.",
     ORANGE),
    ("3", "El SII puede requerir declaración de abuso al TTA",
     "Art. 4 quinquies CT: el Director puede solicitar al Tribunal Tributario y Aduanero "
     "la declaración de abuso → recalificación + atribución íntegra a A + intereses + multas.",
     NAVY),
    ("4", "Existen alternativas legítimas de organización",
     "Profesional independiente, sociedad con socios reales, distribución proporcional "
     "al trabajo, SpA Pro Pyme (14 D N°3) o Régimen Transparente (14 D N°8).",
     GREEN),
    ("5", "La clave es la sustancia económica real",
     "Toda estructura debe demostrar que los socios aportan trabajo, capital, clientela "
     "o dirección efectiva. La forma jurídica por sí sola no es suficiente.",
     TEAL),
]

y_c = 1.65
h_c = 0.92
for num, tit, desc, color in concls:
    circle_label(s, num, 0.78, y_c + h_c/2, r=0.28, fill=color, fc=WHITE, size=15)
    rect(s, 1.2, y_c, 0.07, h_c, color)
    rect(s, 1.32, y_c, 11.55, 0.36, color)
    tb(s, tit, 1.5, y_c + 0.01, 11.3, 0.36,
       12.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    rbox(s, desc, 1.32, y_c + 0.37, 11.55, h_c - 0.38,
         CARD_BG, fc=TEXT_D, size=11.5, line_color=BORDER,
         valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    y_c += h_c + 0.04

# Closing strip
rect(s, 0, 6.6, 13.333, 0.65, NAVY)
rect(s, 0, 6.6, 13.333, 0.05, ORANGE)
tb(s, "Grupo 4   ·   Andrea Añasco  ·  Gema Sepúlveda  ·  Karen Rebolledo  ·  Nicolás Muñoz",
   0.5, 6.67, 12.3, 0.50, 13, color=WHITE, bold=True,
   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


prs.save(OUTPUT)
print(f"✔  PPT moderno generado: {OUTPUT}  —  {len(prs.slides)} slides")
