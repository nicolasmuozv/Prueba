#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Caso 86 — Catálogo de Esquemas Tributarios SII 2025
Presentación PPT v2 — Enfocada en el caso, con ejemplos, dinámica
Continúa slides 1-2 del grupo. Slides 3 en adelante.
Grupo 4: Andrea Añasco, Gema Sepúlveda, Karen Rebolledo, Nicolás Muñoz
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

OUTPUT = "/home/user/Prueba/Caso86_Presentacion.pptx"

# Paleta igual a la del template de la compañera
FONDO       = RGBColor(0xCB, 0xCB, 0xCB)
CAJA_AZUL   = RGBColor(0x8E, 0xB4, 0xE3)
CAJA_DARK   = RGBColor(0x2B, 0x2B, 0x2B)
CAJA_NARAN  = RGBColor(0xFD, 0xE9, 0xD9)
ACCENT_NAR  = RGBColor(0xC5, 0x50, 0x1A)
TEXTO_DARK  = RGBColor(0x1C, 0x1C, 0x1C)
TEXTO_CLARO = RGBColor(0xFF, 0xFF, 0xFF)
AZUL_TIT    = RGBColor(0x1A, 0x3A, 0x5C)
GRIS_TXT    = RGBColor(0x55, 0x55, 0x55)
AMARILLO    = RGBColor(0xE6, 0xA4, 0x3C)
ROJO        = RGBColor(0xB0, 0x30, 0x30)
VERDE       = RGBColor(0x1A, 0x6B, 0x3C)
BLANCO      = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

SW   = prs.slide_width
SH   = prs.slide_height
BLANK = prs.slide_layouts[6]


# ── Helpers ──────────────────────────────────────────────────────────────────

def set_bg(slide, color=FONDO):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    sp = bg._element.getparent()
    sp.remove(bg._element)
    sp.insert(2, bg._element)
    return bg

def add_title(slide, text, left=Inches(0.6), top=Inches(0.3),
              width=Inches(12), height=Inches(0.8),
              size=32, color=TEXTO_DARK, bold=False, italic=False,
              font="Cambria", align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size   = Pt(size)
    r.font.bold   = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font
    return tb

def add_box(slide, text, left, top, width, height,
            fill=CAJA_DARK, font_color=TEXTO_CLARO, size=11,
            bold=False, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, no_fill=False, font="Calibri"):
    box = slide.shapes.add_shape(shape, left, top, width, height)
    if no_fill:
        box.fill.background()
    else:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    box.line.fill.background()
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left   = Inches(0.15)
    tf.margin_right  = Inches(0.15)
    tf.margin_top    = Inches(0.1)
    tf.margin_bottom = Inches(0.1)
    tf.vertical_anchor = valign

    lineas = text.split('\n')
    for i, linea in enumerate(lineas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if linea.startswith("**") and linea.endswith("**"):
            r = p.add_run()
            r.text = linea.strip("*")
            r.font.bold = True
        else:
            r = p.add_run()
            r.text = linea
            r.font.bold = bold
        r.font.size = Pt(size)
        r.font.color.rgb = font_color
        r.font.name = font
    return box

def add_circle(slide, letra, cx, cy, r_in=0.45, fill=AZUL_TIT, font_color=TEXTO_CLARO):
    diam = Inches(r_in * 2)
    box  = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                   cx - diam//2, cy - diam//2, diam, diam)
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.fill.background()
    box.shadow.inherit = False
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p  = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r  = p.add_run()
    r.text = letra
    r.font.size  = Pt(16)
    r.font.bold  = True
    r.font.color.rgb = font_color
    r.font.name  = "Calibri"
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return box

def footer(slide, num):
    add_box(slide,
            f"Caso 86 · Catálogo SII 2025 · Grupo 4 — Magíster Dirección Tributaria UVM  |  {num}",
            Inches(0.3), Inches(7.1), Inches(12.7), Inches(0.3),
            no_fill=True, font_color=GRIS_TXT, size=8, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Presentación del grupo y del caso
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s, AZUL_TIT)

# Título grande
add_title(s, "CASO 86", left=Inches(1), top=Inches(0.9), width=Inches(11),
          height=Inches(1.4), size=54, color=BLANCO, bold=True, font="Cambria",
          align=PP_ALIGN.CENTER)

add_title(s, "Prestación de Servicios Profesionales mediante Sociedad Interpuesta",
          left=Inches(1), top=Inches(2.2), width=Inches(11), height=Inches(0.7),
          size=18, color=CAJA_AZUL, bold=False, font="Calibri", align=PP_ALIGN.CENTER)

add_title(s, "Catálogo de Esquemas Tributarios SII 2025",
          left=Inches(1), top=Inches(2.9), width=Inches(11), height=Inches(0.5),
          size=14, color=RGBColor(0xAA, 0xBB, 0xCC), font="Calibri", align=PP_ALIGN.CENTER)

# Línea divisoria
linea = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                            Inches(1.5), Inches(3.55), Inches(10.3), Inches(0.04))
linea.fill.solid()
linea.fill.fore_color.rgb = CAJA_AZUL
linea.line.fill.background()

# Integrantes
add_title(s, "Grupo 4",
          left=Inches(1), top=Inches(3.8), width=Inches(11), height=Inches(0.5),
          size=14, color=AMARILLO, bold=True, font="Calibri", align=PP_ALIGN.CENTER)

miembros = ["Andrea Añasco", "Gema Sepúlveda", "Karen Rebolledo", "Nicolás Muñoz"]
x0 = Inches(1.0)
w  = Inches(2.6)
for i, m in enumerate(miembros):
    add_box(s, m, x0 + i * (w + Inches(0.18)), Inches(4.4), w, Inches(0.65),
            fill=RGBColor(0x25, 0x4E, 0x80), font_color=BLANCO,
            size=13, bold=False, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_title(s, "Magíster en Dirección Tributaria — Universidad Viña del Mar",
          left=Inches(1), top=Inches(6.8), width=Inches(11), height=Inches(0.4),
          size=10, color=RGBColor(0x88, 0x99, 0xAA), font="Calibri", align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — El esquema: ¿qué hace Pedro?
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "El Esquema — ¿Qué hace Pedro?", size=30, color=TEXTO_DARK)
add_box(s, "Caso real: un médico y la sociedad que reduce su IGC",
        Inches(0.6), Inches(1.05), Inches(12), Inches(0.45),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Perfil del contribuyente
add_box(s,
        "PEDRO — Médico cirujano\n"
        "• Ingresos anuales: $300 millones en honorarios\n"
        "• Sin esquema: tributa en tramo IGC ≈ 40%\n"
        "• Motivación: reducir impuesto fraccionando el ingreso en varios RUT",
        Inches(0.6), Inches(1.7), Inches(5.8), Inches(1.8),
        fill=AZUL_TIT, font_color=BLANCO, size=11.5,
        align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

# El mecanismo
add_box(s,
        "EL MECANISMO\n"
        "Pedro constituye con su señora y su cuñado\n"
        "la 'Sociedad Médica Ltda.' — 33% c/u\n\n"
        "Los $300 MM entran a la sociedad\n"
        "y se reparten en 3 partes iguales:\n"
        "→ Pedro, la señora y el cuñado\n"
        "declaran $100 MM cada uno",
        Inches(6.55), Inches(1.7), Inches(6.15), Inches(1.8),
        fill=CAJA_DARK, font_color=BLANCO, size=11.5,
        align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

# Diagrama de flujo del esquema
y_d = Inches(3.75)
add_box(s, "Pedro\nmédico\n$300 MM", Inches(0.6), y_d, Inches(2.4), Inches(1.3),
        fill=CAJA_AZUL, font_color=AZUL_TIT, size=12, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

CELESTE2 = RGBColor(0x2E, 0x6D, 0xA4)

add_box(s, "→  honorarios  →", Inches(3.1), y_d + Inches(0.4), Inches(2.1), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=12, align=PP_ALIGN.CENTER)

add_box(s, "Sociedad\nMédica Ltda.\n(vehículo)", Inches(5.3), y_d, Inches(2.5), Inches(1.3),
        fill=AZUL_TIT, font_color=BLANCO, size=12, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_box(s, "→  divide 1/3  →", Inches(7.9), y_d + Inches(0.4), Inches(2.1), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=12, align=PP_ALIGN.CENTER)

add_box(s, "Pedro · Señora · Cuñado\n$100 MM c/u → IGC tramo bajo",
        Inches(10.1), y_d, Inches(2.9), Inches(1.3),
        fill=ROJO, font_color=BLANCO, size=11, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# Resultado
add_box(s,
        "RESULTADO: lo que debería tributar al ~40% queda dividido en tres tramos más bajos. "
        "El ahorro puede superar los $20–40 MM al año. "
        "La señora y el cuñado NUNCA trabajaron para la sociedad.",
        Inches(0.6), Inches(5.3), Inches(12.1), Inches(0.75),
        fill=ACCENT_NAR, font_color=BLANCO, size=12, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# Pregunta disparadora
add_box(s, "¿Es esto legal? ¿Abuso o planificación legítima?",
        Inches(0.6), Inches(6.2), Inches(12.1), Inches(0.5),
        fill=CAJA_DARK, font_color=AMARILLO, size=14, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

footer(s, 4)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — ¿La sociedad es legítima? Requisitos
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "¿La Sociedad D es Legítima?", size=30, color=TEXTO_DARK)
add_box(s, "Requisitos legales vs. realidad del Caso 86",
        Inches(0.6), Inches(1.05), Inches(12), Inches(0.45),
        no_fill=True, font_color=GRIS_TXT, size=14)

y_hdr = Inches(1.7)
for txt, left, w in [("Requisito (Circ. N°21/1991 y N°50/2020)", Inches(0.6), Inches(7.5)),
                     ("¿Cumple el Caso 86?", Inches(8.2), Inches(4.5))]:
    add_box(s, txt, left, y_hdr, w, Inches(0.5),
            fill=AZUL_TIT, font_color=BLANCO, size=11.5, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

reqs = [
    ("Sociedad de personas",                                         "✔  Cumple",       VERDE),
    ("Objeto exclusivo: servicios profesionales",                    "✔  Cumple",       VERDE),
    ("Servicios prestados por los socios",                           "✔  Solo A (Pedro) trabaja", AMARILLO),
    ("TODOS los socios ejercen su profesión para la sociedad",       "✖  NO — B y C no prestan servicios", ROJO),
    ("No se acepta que un socio solo aporte capital sin trabajar",   "✖  NO — B y C solo aportan capital", ROJO),
]
y = Inches(2.25)
h = Inches(0.68)
for req, estado, color in reqs:
    add_box(s, req,    Inches(0.6), y, Inches(7.5), h,
            fill=RGBColor(0xF4, 0xF6, 0xF9), font_color=TEXTO_DARK,
            size=10.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, estado, Inches(8.2), y, Inches(4.5), h,
            fill=color, font_color=BLANCO,
            size=11, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    y += h + Inches(0.05)

add_box(s,
        "La sociedad D NO califica como 'sociedad de profesionales'. "
        "Sin esa calificación, el esquema cae: no hay amparo legal para la distribución desigual del IGC.",
        Inches(0.6), Inches(6.35), Inches(12.1), Inches(0.65),
        fill=ROJO, font_color=BLANCO, size=12.5, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

footer(s, 5)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — ¿Abuso o Simulación? La calificación importa
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "¿Abuso o Simulación?", size=30, color=TEXTO_DARK)
add_box(s, "La calificación correcta define qué norma aplica y qué debe probar el SII",
        Inches(0.6), Inches(1.05), Inches(12), Inches(0.45),
        no_fill=True, font_color=GRIS_TXT, size=14)

lx = Inches(0.6)
rx = Inches(6.95)
cw = Inches(6.15)

add_box(s, "ABUSO  (Art. 4 ter CT)", lx, Inches(1.7), cw, Inches(0.55),
        fill=ROJO, font_color=BLANCO, size=15, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_box(s,
        "✔  Los actos son REALES: la sociedad existe, los socios son reales, "
        "el capital fue aportado.\n\n"
        "✔  El PROBLEMA es la FORMA: se usa la figura de 'sociedad de profesionales' "
        "para algo que no corresponde.\n\n"
        "✔  B y C no aportan valor económico real. Su único rol es absorber utilidades "
        "de Pedro para bajar su tramo de IGC.\n\n"
        "✔  Único efecto distinto del tributario: NINGUNO.",
        lx, Inches(2.3), cw, Inches(4.45),
        fill=RGBColor(0xFD, 0xEA, 0xEA), font_color=TEXTO_DARK,
        size=11.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

add_box(s, "SIMULACIÓN  (Art. 4 quáter CT)", rx, Inches(1.7), cw, Inches(0.55),
        fill=ACCENT_NAR, font_color=BLANCO, size=15, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_box(s,
        "✖  Requiere que los actos DISIMULEN algo distinto de lo declarado.\n\n"
        "✖  En el caso NO hay acto oculto: la sociedad y sus socios son lo que parecen.\n\n"
        "✖  No hay divergencia entre voluntad real y declarada.\n\n"
        "✖  Ejemplo de simulación: un contrato de 'arriendo' que en realidad es "
        "una donación. En el Caso 86 no hay nada oculto.",
        rx, Inches(2.3), cw, Inches(4.45),
        fill=RGBColor(0xFD, 0xF0, 0xE6), font_color=TEXTO_DARK,
        size=11.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

add_box(s,
        "POSICIÓN DEL GRUPO:  ABUSO de formas jurídicas — Art. 4 ter CT",
        Inches(0.6), Inches(6.87), Inches(12.1), Inches(0.5),
        fill=AZUL_TIT, font_color=BLANCO, size=14, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

footer(s, 6)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — El impacto en números
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "El Impacto en Números", size=30, color=TEXTO_DARK)
add_box(s, "Ejemplo ilustrativo — ¿Cuánto pierde el Fisco?",
        Inches(0.6), Inches(1.05), Inches(12), Inches(0.45),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Encabezado tabla
y_h = Inches(1.7)
cols_t = [
    ("Escenario",              Inches(0.6),  Inches(3.0)),
    ("¿Quién declara qué?",    Inches(3.65), Inches(4.4)),
    ("Tramo IGC aprox.",       Inches(8.1),  Inches(2.5)),
    ("Efecto",                 Inches(10.65),Inches(2.0)),
]
for txt, l, w in cols_t:
    add_box(s, txt, l, y_h, w, Inches(0.5),
            fill=AZUL_TIT, font_color=BLANCO, size=11, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

filas_t = [
    ("Sin esquema\n(Pedro trabaja solo)",
     "Pedro declara $300 MM",
     "Tramo marginal ≈ 40%",
     "Tributación\ncorrecta",
     RGBColor(0xE8, 0xF3, 0xE8)),
    ("Con esquema D\n(3 socios simbólicos)",
     "Pedro $100 MM · Señora $100 MM · Cuñado $100 MM",
     "Cada uno en tramo ≈ 23–35%",
     "Ahorro fiscal\nARTIFICIAL",
     RGBColor(0xFB, 0xE5, 0xE5)),
    ("Diferencial\n(perjuicio fiscal)",
     "Pedro genera el 100% de los ingresos. La señora y el cuñado no atendieron pacientes.",
     "Diferencia real de tramos = ahorro real de Pedro",
     "ABUSO\nArt. 4 ter CT",
     RGBColor(0xFD, 0xE9, 0xD9)),
]
y_t = Inches(2.25)
for esc, atrib, tramo, res, color in filas_t:
    h = Inches(1.05)
    add_box(s, esc,   Inches(0.6),  y_t, Inches(3.0),  h, fill=color, font_color=TEXTO_DARK, size=10.5, bold=True,  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, atrib, Inches(3.65), y_t, Inches(4.4),  h, fill=color, font_color=TEXTO_DARK, size=10.5,             align=PP_ALIGN.LEFT,   valign=MSO_ANCHOR.MIDDLE)
    add_box(s, tramo, Inches(8.1),  y_t, Inches(2.5),  h, fill=color, font_color=TEXTO_DARK, size=10.5,             align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, res,   Inches(10.65),y_t, Inches(2.0),  h, fill=color, font_color=TEXTO_DARK, size=10,   bold=True,  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    y_t += h + Inches(0.05)

add_box(s,
        "Si el TTA declara el abuso:  Pedro queda con $300 MM en su IGC + intereses + multas. "
        "La señora y el cuñado también deben reliquidar sus IGC. (Art. 4 quinquies CT)",
        Inches(0.6), Inches(5.6), Inches(12.1), Inches(0.6),
        fill=ROJO, font_color=BLANCO, size=12, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_box(s,
        "Nota: cifras ilustrativas. El ahorro real depende de los ingresos totales, "
        "rentas de los socios y otras deducciones.",
        Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.4),
        no_fill=True, font_color=GRIS_TXT, size=9)

footer(s, 7)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — ¿Qué puede hacer Pedro? Alternativas legítimas
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "¿Qué Puede Hacer Pedro?", size=30, color=TEXTO_DARK)
add_box(s, "5 alternativas legítimas para organizar su actividad profesional",
        Inches(0.6), Inches(1.05), Inches(12), Inches(0.45),
        no_fill=True, font_color=GRIS_TXT, size=14)

alts = [
    ("1", "Profesional independiente",
     "Pedro emite boletas de honorarios y declara directamente. Puede deducir gastos "
     "efectivos o presuntos (30%, tope 15 UTA). Sin riesgo NGA.",
     VERDE),
    ("2", "Sociedad con socios que trabajen de verdad",
     "Pedro se asocia con otros médicos que efectivamente atiendan pacientes. "
     "Cada socio aporta trabajo real → distribución de utilidades legítima.",
     AZUL_TIT),
    ("3", "Distribución proporcional al trabajo (misma sociedad)",
     "Si D se mantiene: Pedro recibe el 90% de utilidades y B/C el 5% c/u, "
     "acorde a su aporte real. Debe haber algún rol efectivo de B y C.",
     CAJA_AZUL),
    ("4", "SpA — régimen Pro Pyme 14D N°3",
     "Pedro constituye SpA. IDPC al 25%; Pedro paga IGC sobre retiros. "
     "Permite reinversión y planificación. Exige sustancia económica.",
     ACCENT_NAR),
    ("5", "Régimen Transparente 14D N°8",
     "La empresa no paga IDPC; socios pagan IGC directo. "
     "Solo válido si los socios tienen rol económico real; no resuelve el problema "
     "si B y C siguen sin trabajar.",
     AMARILLO),
]
y_a = Inches(1.7)
for num, tit, txt, color in alts:
    h = Inches(0.98)
    add_circle(s, num, Inches(0.9), y_a + h // 2, r_in=0.28, fill=color)
    add_box(s, tit, Inches(1.35), y_a, Inches(4.0), h,
            fill=color, font_color=BLANCO if color != CAJA_AZUL else AZUL_TIT,
            size=11.5, bold=True, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, txt, Inches(5.45), y_a, Inches(7.5), h,
            fill=RGBColor(0xF4, 0xF6, 0xF9), font_color=TEXTO_DARK,
            size=10.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    y_a += h + Inches(0.05)

footer(s, 8)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — Conclusiones
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Conclusiones", size=34, color=TEXTO_DARK)
add_box(s, "Caso 86 — ¿Qué aprendemos?",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.45),
        no_fill=True, font_color=GRIS_TXT, size=14)

concls = [
    ("1", "Abuso, no simulación",
     "Los actos son reales. El problema es la FORMA inapropiada. "
     "La sociedad y los socios existen, pero B y C no aportan valor económico real.",
     ROJO),
    ("2", "La sociedad D no califica",
     "Sin cumplir los requisitos de las Circulares N°21/1991 y N°50/2020, "
     "D no es sociedad de profesionales → el esquema no tiene amparo legal.",
     ACCENT_NAR),
    ("3", "El SII puede actuar — consecuencias concretas",
     "Art. 4 quinquies CT: declaración de abuso ante el TTA → "
     "Pedro recibe los $300 MM íntegros en su IGC + intereses + multas.",
     AZUL_TIT),
    ("4", "La clave: sustancia económica",
     "Cualquier estructura es aceptable si cada socio aporta trabajo real, "
     "capital relevante, clientela o dirección. La forma jurídica sola no basta.",
     VERDE),
    ("5", "Mensaje del Caso 86",
     "Planificación tributaria ≠ abuso. El límite está en la sustancia económica. "
     "Si el único motivo es el ahorro fiscal, el SII tiene herramientas para actuar.",
     CAJA_AZUL),
]
y_c = Inches(1.75)
for num, tit, txt, color in concls:
    h = Inches(0.88)
    add_circle(s, num, Inches(0.88), y_c + h // 2, r_in=0.28, fill=color,
               font_color=BLANCO if color != CAJA_AZUL else AZUL_TIT)
    add_box(s, tit, Inches(1.35), y_c, Inches(3.2), h,
            fill=color, font_color=BLANCO if color != CAJA_AZUL else AZUL_TIT,
            size=11.5, bold=True, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, txt, Inches(4.65), y_c, Inches(8.3), h,
            fill=RGBColor(0xF4, 0xF6, 0xF9), font_color=TEXTO_DARK,
            size=10.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    y_c += h + Inches(0.08)

add_box(s, "Grupo 4: Andrea Añasco · Gema Sepúlveda · Karen Rebolledo · Nicolás Muñoz",
        Inches(0.6), Inches(7.0), Inches(12.1), Inches(0.35),
        fill=AZUL_TIT, font_color=BLANCO, size=10, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

footer(s, 9)


prs.save(OUTPUT)
print(f"✔ Presentación PPT generada: {OUTPUT}")
