#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Caso 86 — Catálogo de Esquemas Tributarios SII 2025
Presentación PPT — Continuación del avance del grupo
Se respeta el formato y estilo de las slides 1 y 2 ya entregadas:
- Fondo gris claro
- Títulos negros, fuente serif
- Cajas azules para participantes
- Cajas oscuras (negro/grafito) para texto principal
- Caja lateral naranja claro con RESULTADO
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

OUTPUT = "/home/user/Prueba/Caso86_Presentacion.pptx"

# ── Colores del template del compañera ────────────────────────────────────────
FONDO       = RGBColor(0xCB, 0xCB, 0xCB)   # gris claro
CAJA_AZUL   = RGBColor(0x8E, 0xB4, 0xE3)   # azul claro participantes
CAJA_DARK   = RGBColor(0x2B, 0x2B, 0x2B)   # negro/grafito texto
CAJA_NARAN  = RGBColor(0xFD, 0xE9, 0xD9)   # naranja claro resultado
ACCENT_NAR  = RGBColor(0xC5, 0x50, 0x1A)   # naranja oscuro
TEXTO_DARK  = RGBColor(0x1C, 0x1C, 0x1C)
TEXTO_CLARO = RGBColor(0xFF, 0xFF, 0xFF)
AZUL_TIT    = RGBColor(0x1A, 0x3A, 0x5C)
GRIS_TXT    = RGBColor(0x55, 0x55, 0x55)
AMARILLO    = RGBColor(0xE6, 0xA4, 0x3C)
ROJO        = RGBColor(0xB0, 0x30, 0x30)
VERDE       = RGBColor(0x1A, 0x6B, 0x3C)

# ── Crear presentación 16:9 ───────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

SW = prs.slide_width
SH = prs.slide_height

BLANK = prs.slide_layouts[6]

def set_bg(slide, color=FONDO):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    # Mandar al fondo
    spTree = bg._element.getparent()
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)
    return bg

def add_title(slide, text, left=Inches(0.6), top=Inches(0.3),
              width=Inches(12), height=Inches(0.8),
              size=32, color=TEXTO_DARK, bold=False, italic=False,
              font="Cambria"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
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
    tf.margin_left  = Inches(0.15)
    tf.margin_right = Inches(0.15)
    tf.margin_top   = Inches(0.1)
    tf.margin_bottom= Inches(0.1)
    tf.vertical_anchor = valign

    # Soportar texto con múltiples párrafos separados por \n
    lineas = text.split('\n')
    for i, linea in enumerate(lineas):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        # Detectar si la línea empieza con marcas especiales
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

def add_circle_label(slide, letra, cx, cy, r_in=0.45, fill=AZUL_TIT, font_color=TEXTO_CLARO):
    """Letra en círculo, útil para indicar A, B, C, D."""
    diam = Inches(r_in*2)
    left = cx - diam/2
    top  = cy - diam/2
    box = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, diam, diam)
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.fill.background()
    box.shadow.inherit = False
    tf = box.text_frame
    tf.margin_left  = 0
    tf.margin_right = 0
    tf.margin_top   = 0
    tf.margin_bottom= 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = letra
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = font_color
    r.font.name = "Calibri"
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return box

def add_pill(slide, text, left, top, width, height,
             fill=AMARILLO, font_color=TEXTO_CLARO, size=12, bold=True):
    """Píldora redondeada para destacar."""
    return add_box(slide, text, left, top, width, height,
                   fill=fill, font_color=font_color, size=size, bold=bold,
                   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE)

def add_bullets(slide, items, left, top, width, height,
                size=11, font_color=TEXTO_DARK, bullet_char="•",
                font="Calibri"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(6)
        r1 = p.add_run()
        r1.text = f"{bullet_char}  "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = AZUL_TIT
        r1.font.name = font

        r2 = p.add_run()
        r2.text = item
        r2.font.size = Pt(size)
        r2.font.color.rgb = font_color
        r2.font.name = font
    return tb

def slide_footer(slide, num):
    """Numeración inferior."""
    add_box(slide, f"Caso 86 · Magíster Dirección Tributaria UVM · Slide {num}",
            Inches(0.3), Inches(7.1), Inches(13), Inches(0.3),
            no_fill=True, font_color=GRIS_TXT, size=8, align=PP_ALIGN.RIGHT)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Requisitos de Sociedad de Profesionales
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Requisitos Sociedad de Profesionales", size=32, color=TEXTO_DARK)
add_box(s, "Verificación de cumplimiento — Caso 86",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Tabla de requisitos
encabezado = [
    ("Requisito (Circular 21/1991 y 50/2020)", Inches(0.6),  Inches(7.5)),
    ("¿Se cumple en el Caso 86?",              Inches(8.2),  Inches(4.5)),
]
y_hdr = Inches(1.85)
for txt, left, w in encabezado:
    add_box(s, txt, left, y_hdr, w, Inches(0.5),
            fill=AZUL_TIT, font_color=TEXTO_CLARO, size=12, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

requisitos = [
    ("Debe tratarse de una sociedad de personas", "Cumple — D es sociedad de personas", VERDE),
    ("Objeto exclusivo: prestación de servicios o asesorías profesionales", "Cumple", VERDE),
    ("Servicios prestados por intermedio de sus socios", "Cumple — A presta los servicios", VERDE),
    ("TODOS los socios deben ejercer su profesión para la sociedad", "NO CUMPLE — B y C no prestan servicios", ROJO),
    ("No es aceptable que un socio solo aporte capital", "NO CUMPLE — B y C solo aportan capital", ROJO),
    ("La forma de remuneración no es relevante; sí la prestación efectiva", "Solo A presta labor efectiva", ROJO),
]

y = Inches(2.4)
alto = Inches(0.55)
for req, cumple, color in requisitos:
    add_box(s, req, Inches(0.6), y, Inches(7.5), alto,
            fill=RGBColor(0xF4,0xF6,0xF9), font_color=TEXTO_DARK,
            size=10, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, cumple, Inches(8.2), y, Inches(4.5), alto,
            fill=color, font_color=TEXTO_CLARO,
            size=10.5, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    y += alto + Inches(0.05)

# Conclusión
add_box(s, "→ La sociedad D NO califica como 'sociedad de profesionales' por incumplimiento "
        "de requisitos formales esenciales (Circulares N°21/1991 y N°50/2020 del SII).",
        Inches(0.6), Inches(6.4), Inches(12), Inches(0.55),
        fill=ACCENT_NAR, font_color=TEXTO_CLARO, size=12, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

slide_footer(s, 3)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Marco NGA: Arts. 4 bis, ter, quáter y quinquies CT
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Normas Generales Antielusivas (NGA)", size=32, color=TEXTO_DARK)
add_box(s, "Arts. 4 bis, 4 ter, 4 quáter y 4 quinquies del Código Tributario",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Cuatro cajas — una por norma
art_data = [
    ("Art. 4 bis CT",
     "Principio de Buena Fe",
     "Las obligaciones tributarias nacen y se exigen conforme a la NATURALEZA JURÍDICA de los "
     "hechos, cualquiera sea la forma o denominación que las partes les den.",
     AZUL_TIT),
    ("Art. 4 ter CT",
     "ABUSO de las formas jurídicas",
     "Existe abuso cuando se evita total o parcialmente el hecho gravado, o se disminuye la "
     "base imponible, mediante actos que no produzcan efectos jurídicos o económicos relevantes "
     "distintos de los meramente tributarios.",
     ROJO),
    ("Art. 4 quáter CT",
     "SIMULACIÓN",
     "Habrá simulación cuando los actos disimulen la configuración del hecho gravado o la "
     "naturaleza de los elementos de la obligación tributaria, o su monto o data de nacimiento.",
     ACCENT_NAR),
    ("Art. 4 quinquies CT",
     "Procedimiento",
     "La existencia de abuso o simulación será declarada por el Tribunal Tributario y "
     "Aduanero, a requerimiento del Director del SII. Circular N°65 de 2015 imparte instrucciones.",
     VERDE),
]
x0 = Inches(0.6)
y0 = Inches(1.85)
w_card = Inches(6.05)
h_card = Inches(2.5)
for i, (codigo, titulo, texto, color) in enumerate(art_data):
    col = i % 2
    fila = i // 2
    x = x0 + col*(w_card + Inches(0.15))
    y = y0 + fila*(h_card + Inches(0.2))
    # Header coloreado
    add_box(s, f"{codigo}  ·  {titulo}", x, y, w_card, Inches(0.5),
            fill=color, font_color=TEXTO_CLARO, size=12.5, bold=True,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    # Cuerpo
    add_box(s, texto, x, y + Inches(0.5), w_card, h_card - Inches(0.5),
            fill=RGBColor(0xF4,0xF6,0xF9), font_color=TEXTO_DARK,
            size=10.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

slide_footer(s, 4)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — ¿Abuso o Simulación? Análisis comparado
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "¿Abuso o Simulación?", size=32, color=TEXTO_DARK)
add_box(s, "Calificación jurídica del esquema",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Dos columnas comparativas
left_col_x = Inches(0.6)
right_col_x = Inches(6.95)
col_w = Inches(6.15)

# ABUSO
add_box(s, "ABUSO  (Art. 4 ter CT)", left_col_x, Inches(1.85), col_w, Inches(0.55),
        fill=ROJO, font_color=TEXTO_CLARO, size=14, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_box(s, "✔ Los actos son REALES, pero la FORMA es inapropiada.\n\n"
           "✔ La sociedad D existe formalmente, hay socios, capital aportado, "
           "utilidades repartidas.\n\n"
           "✔ El abuso radica en utilizar una figura societaria que carece de sustancia "
           "económica para B y C: solo aportan capital insignificante.\n\n"
           "✔ Único efecto relevante distinto del tributario: NO HAY uno.\n\n"
           "✔ El esquema reduce artificialmente la progresión del IGC de A.",
        left_col_x, Inches(2.45), col_w, Inches(4.4),
        fill=RGBColor(0xFD, 0xEA, 0xEA), font_color=TEXTO_DARK,
        size=11, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

# SIMULACIÓN
add_box(s, "SIMULACIÓN  (Art. 4 quáter CT)", right_col_x, Inches(1.85), col_w, Inches(0.55),
        fill=ACCENT_NAR, font_color=TEXTO_CLARO, size=14, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

add_box(s, "✖ Supondría que los actos jurídicos disimulan algo distinto de lo declarado.\n\n"
           "✖ En el caso, NO hay un acto disimulado: los participantes son realmente socios.\n\n"
           "✖ No hay divergencia consciente entre voluntad real y declarada.\n\n"
           "✖ El SII y la doctrina entienden simulación como existencia de un acto "
           "secreto que oculta otro.\n\n"
           "✖ Se descarta esta figura para el Caso 86.",
        right_col_x, Inches(2.45), col_w, Inches(4.4),
        fill=RGBColor(0xFD, 0xF0, 0xE6), font_color=TEXTO_DARK,
        size=11, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)

# Caja conclusión inferior
add_box(s, "POSICIÓN DEL GRUPO:  el Caso 86 configura ABUSO DE FORMAS JURÍDICAS (Art. 4 ter CT)",
        Inches(0.6), Inches(6.95), Inches(12.1), Inches(0.5),
        fill=AZUL_TIT, font_color=TEXTO_CLARO, size=13, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

slide_footer(s, 5)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — Fundamentos de la posición
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "¿Por qué es Abuso? — Fundamentos", size=32, color=TEXTO_DARK)
add_box(s, "Caso 86 configura los 3 elementos del Art. 4 ter CT",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Tres cajas verticales con los elementos
elementos = [
    ("1",
     "Reducción de la carga tributaria",
     "El esquema diluye los ingresos de A en tres tramos independientes del IGC, "
     "reduciendo la tasa marginal efectiva. La carga del IGC del contribuyente A se "
     "ve disminuida en relación a lo que correspondería si los ingresos los percibiera "
     "directamente, sea como independiente del Art. 42 N°2 o como único prestador "
     "del servicio.",
     ROJO),
    ("2",
     "Ausencia de efectos jurídicos o económicos relevantes",
     "La participación de B y C carece de motivo económico válido distinto del tributario:\n"
     "• Capital aportado insignificante frente a los ingresos\n"
     "• No prestan servicios profesionales\n"
     "• No aportan clientela ni cartera de clientes\n"
     "• No tienen rol estratégico ni decisional relevante\n"
     "• Su única función es absorber utilidades para diluir el IGC de A.",
     ACCENT_NAR),
    ("3",
     "Forma jurídica artificiosa para el sustrato económico",
     "La figura 'sociedad de profesionales' del Art. 42 N°2 LIR está concebida para "
     "profesionales que efectivamente trabajan en conjunto. Su uso para canalizar la "
     "labor de uno solo, con socios pasivos, constituye una utilización inapropiada "
     "y artificial de la institución jurídica.",
     AZUL_TIT),
]

y = Inches(1.85)
for num, titulo_e, texto_e, color in elementos:
    h = Inches(1.6)
    # Número grande circular
    add_circle_label(s, num, Inches(0.95), y + h/2, r_in=0.35, fill=color)
    # Título y texto
    add_box(s, titulo_e, Inches(1.5), y, Inches(11.3), Inches(0.4),
            no_fill=True, font_color=color, size=14, bold=True)
    add_box(s, texto_e, Inches(1.5), y + Inches(0.4), Inches(11.3), h - Inches(0.4),
            fill=RGBColor(0xF4,0xF6,0xF9), font_color=TEXTO_DARK,
            size=10.5, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)
    y += h + Inches(0.05)

slide_footer(s, 6)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Efecto tributario / perjuicio fiscal
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Efecto Tributario del Esquema", size=32, color=TEXTO_DARK)
add_box(s, "Ejemplo ilustrativo — Dilución de la progresión del IGC",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

# Encabezado tabla
y_h = Inches(1.85)
cols = [
    ("Escenario",                     Inches(0.6), Inches(3.2)),
    ("Atribución de la renta",        Inches(3.85), Inches(4.5)),
    ("Tramo marginal IGC aprox.",     Inches(8.4),  Inches(2.6)),
    ("Resultado",                     Inches(11.05),Inches(1.7)),
]
for txt, left, w in cols:
    add_box(s, txt, left, y_h, w, Inches(0.5),
            fill=AZUL_TIT, font_color=TEXTO_CLARO, size=11, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

# Filas
filas = [
    ("Sin esquema\n(A trabajaría solo, sin sociedad)",
     "A: $300 MM íntegro",
     "Tramo alto (≈ 40%)",
     "Tributación correcta",
     RGBColor(0xE8, 0xF3, 0xE8)),
    ("Con esquema D\n(3 socios, capital simbólico)",
     "A: $100 MM\nB: $100 MM\nC: $100 MM",
     "Cada uno en tramo medio",
     "Ahorro fiscal artificial",
     RGBColor(0xFB, 0xE5, 0xE5)),
    ("Diferencial",
     "Mismo flujo económico real:\nlos ingresos los generó A",
     "Disminución total IGC",
     "ABUSO Art. 4 ter CT",
     RGBColor(0xFD, 0xE9, 0xD9)),
]
y = Inches(2.4)
for esc, atrib, tramo, result, color in filas:
    h = Inches(1.0)
    add_box(s, esc, Inches(0.6), y, Inches(3.2), h,
            fill=color, font_color=TEXTO_DARK, size=10.5, bold=True,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, atrib, Inches(3.85), y, Inches(4.5), h,
            fill=color, font_color=TEXTO_DARK, size=10.5,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, tramo, Inches(8.4), y, Inches(2.6), h,
            fill=color, font_color=TEXTO_DARK, size=10.5,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, result, Inches(11.05), y, Inches(1.7), h,
            fill=color, font_color=TEXTO_DARK, size=10, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    y += h + Inches(0.05)

# Nota
add_box(s, "Nota: cifras ilustrativas. El perjuicio fiscal efectivo dependerá del nivel real "
        "de ingresos, tasas marginales aplicables y otras rentas de los socios.",
        Inches(0.6), Inches(6.0), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=10)

# Caja consecuencia
add_box(s, "Consecuencia jurídica:  el SII puede solicitar al TTA la declaración de ABUSO "
        "(Art. 4 quinquies CT), con efecto de recalificación y atribución íntegra de los "
        "ingresos a A para el cálculo del IGC + intereses + multas.",
        Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.55),
        fill=ROJO, font_color=TEXTO_CLARO, size=11.5, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

slide_footer(s, 7)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Alternativas de Solución (1 de 2)
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Alternativas de Solución (1/2)", size=32, color=TEXTO_DARK)
add_box(s, "Estructuras legítimas que A puede adoptar",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

alts_1 = [
    ("Alt. 1",
     "Tributación directa como profesional independiente (Art. 42 N°2 LIR)",
     "A ejerce como persona natural, emite boletas de honorarios, declara en segunda "
     "categoría. Opción gastos efectivos o presuntos (30%, tope 15 UTA). Paga IGC sobre "
     "la totalidad. Opción más simple y libre de cuestionamiento.",
     VERDE),
    ("Alt. 2",
     "Sociedad de profesionales con socios que efectivamente trabajen",
     "A se asocia solo con profesionales que ejerzan su profesión para la sociedad. "
     "Todos coadyuvan al servicio. D califica legítimamente y puede optar por tributar "
     "en 1ª o 2ª categoría conforme Art. 42 N°2 inciso 3°.",
     CAJA_AZUL),
    ("Alt. 3",
     "Distribución de utilidades proporcional al trabajo prestado",
     "Si D se mantiene con A, B y C, los estatutos pueden contemplar:\n"
     "(i) participación diferenciada (ej. A 90%, B 5%, C 5%)\n"
     "(ii) sueldo patronal a A previo a repartir utilidades.\n"
     "La distribución debe reflejar el aporte económico real.",
     AMARILLO),
]
y = Inches(1.85)
for tag, titulo_a, texto_a, color in alts_1:
    h = Inches(1.65)
    add_box(s, tag, Inches(0.6), y, Inches(1.4), h,
            fill=color, font_color=TEXTO_CLARO if color != CAJA_AZUL else AZUL_TIT,
            size=18, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, titulo_a, Inches(2.1), y, Inches(10.7), Inches(0.5),
            fill=color, font_color=TEXTO_CLARO if color != CAJA_AZUL else AZUL_TIT,
            size=12.5, bold=True,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, texto_a, Inches(2.1), y + Inches(0.5), Inches(10.7), h - Inches(0.5),
            fill=RGBColor(0xF4,0xF6,0xF9), font_color=TEXTO_DARK,
            size=11, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)
    y += h + Inches(0.1)

slide_footer(s, 8)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — Alternativas de Solución (2 de 2)
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Alternativas de Solución (2/2)", size=32, color=TEXTO_DARK)
add_box(s, "Vehículos societarios alternativos con sustancia económica",
        Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

alts_2 = [
    ("Alt. 4",
     "Sociedad por Acciones (SpA) — régimen 14D N°3 Pro Pyme General",
     "Constituir SpA acogida al régimen 14D N°3 con tasa IDPC del 25%. A puede ser único "
     "accionista o tener accionistas con rol económico real. Permite planificación de retiros, "
     "reinversión y deducción amplia de gastos. Sin restricción de la figura 'sociedad de "
     "profesionales' pero exige sustancia.",
     ACCENT_NAR),
    ("Alt. 5",
     "Régimen Pro Pyme Transparente — 14D N°8",
     "Si socios son personas naturales, considerar 14D N°8 (Transparencia Tributaria). "
     "La empresa no paga IDPC; los socios pagan IGC directo. ATENCIÓN: este régimen NO "
     "soluciona el problema de fondo si B y C siguen siendo ficticios — el SII puede "
     "igualmente aplicar NGA. Solo viable con socios reales.",
     ROJO),
]
y = Inches(1.85)
for tag, titulo_a, texto_a, color in alts_2:
    h = Inches(1.85)
    add_box(s, tag, Inches(0.6), y, Inches(1.4), h,
            fill=color, font_color=TEXTO_CLARO, size=18, bold=True,
            align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, titulo_a, Inches(2.1), y, Inches(10.7), Inches(0.5),
            fill=color, font_color=TEXTO_CLARO, size=12.5, bold=True,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    add_box(s, texto_a, Inches(2.1), y + Inches(0.5), Inches(10.7), h - Inches(0.5),
            fill=RGBColor(0xF4,0xF6,0xF9), font_color=TEXTO_DARK,
            size=11, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP)
    y += h + Inches(0.15)

# Recomendación final caja
add_box(s, "Recomendación práctica:  cualquier estructura debe poder acreditar SUSTANCIA "
        "ECONÓMICA REAL. La forma jurídica por sí sola no es suficiente; el SII y los tribunales "
        "exigirán demostración del aporte efectivo (trabajo, capital, clientela, dirección) de cada socio.",
        Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.95),
        fill=AZUL_TIT, font_color=TEXTO_CLARO, size=11.5, bold=True,
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

slide_footer(s, 9)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — Conclusiones
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
set_bg(s)
add_title(s, "Conclusiones", size=36, color=TEXTO_DARK)
add_box(s, "Caso 86 — Catálogo de Esquemas Tributarios SII 2025",
        Inches(0.6), Inches(1.15), Inches(12), Inches(0.5),
        no_fill=True, font_color=GRIS_TXT, size=14)

conclusiones = [
    ("1",
     "El esquema configura ABUSO de las formas jurídicas (Art. 4 ter CT). "
     "NO es simulación: los actos son reales pero la forma es inapropiada.",
     ROJO),
    ("2",
     "La sociedad D NO califica como sociedad de profesionales conforme a las "
     "Circulares N°21/1991 y N°50/2020 del SII (B y C no ejercen profesión para D).",
     ACCENT_NAR),
    ("3",
     "El SII puede solicitar al TTA la declaración de abuso (Art. 4 quinquies CT), "
     "con recalificación: atribución íntegra de ingresos a A para el IGC + intereses + multas.",
     AZUL_TIT),
    ("4",
     "Existen alternativas legítimas: profesional independiente, sociedad de profesionales "
     "con socios reales, SpA Pro Pyme, distribución proporcional al trabajo real.",
     VERDE),
    ("5",
     "Toda estructura debe acreditar SUSTANCIA ECONÓMICA. La forma jurídica no basta; "
     "los socios deben aportar trabajo, capital, clientela o rol estratégico verificable.",
     CAJA_AZUL),
]
y = Inches(1.85)
for num, texto_c, color in conclusiones:
    h = Inches(0.85)
    add_circle_label(s, num, Inches(0.9), y + h/2, r_in=0.32, fill=color,
                     font_color=TEXTO_CLARO if color != CAJA_AZUL else AZUL_TIT)
    add_box(s, texto_c, Inches(1.4), y, Inches(11.4), h,
            fill=RGBColor(0xF4,0xF6,0xF9), font_color=TEXTO_DARK,
            size=12, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE)
    y += h + Inches(0.1)

# Cierre
add_box(s, "Magíster en Dirección Tributaria — Universidad Viña del Mar",
        Inches(0.6), Inches(6.85), Inches(12.1), Inches(0.35),
        no_fill=True, font_color=GRIS_TXT, size=10,
        align=PP_ALIGN.CENTER)

slide_footer(s, 10)

prs.save(OUTPUT)
print(f"✔ Presentación PPT generada: {OUTPUT}")
