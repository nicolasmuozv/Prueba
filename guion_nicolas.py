#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Guion de presentación — Nicolás Muñoz
Caso 86 — Catálogo de Esquemas Tributarios SII 2025
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT = "/home/user/Prueba/Guion_Nicolas.docx"

NAVY   = RGBColor(0x1A, 0x3A, 0x5C)
ORANGE = RGBColor(0xE8, 0x50, 0x0A)
GRAY   = RGBColor(0x55, 0x55, 0x55)
BLACK  = RGBColor(0x1C, 0x1C, 0x1C)

doc = Document()

# Márgenes
for sec in doc.sections:
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin   = Cm(3.0)
    sec.right_margin  = Cm(2.5)

def set_font(run, size, bold=False, italic=False, color=BLACK):
    run.font.name  = "Calibri"
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = color

def shade_paragraph(paragraph, hex_color="1A3A5C"):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)

def heading(text, color=NAVY):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run(text)
    set_font(r, 13, bold=True, color=color)
    return p

def slide_label(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after  = Pt(2)
    shade_paragraph(p, "1A3A5C")
    r = p.add_run(f"  {text}")
    set_font(r, 11, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    return p

def speech(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Cm(0.5)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    set_font(r, 12, italic=True, color=BLACK)
    return p

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Cm(0.5)
    r = p.add_run(f"[{text}]")
    set_font(r, 9.5, italic=True, color=GRAY)
    return p

def divider():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run("─" * 80)
    set_font(r, 7, color=RGBColor(0xCC, 0xCC, 0xCC))

# ── PORTADA ───────────────────────────────────────────────────────────────────
titulo = doc.add_paragraph()
titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
titulo.paragraph_format.space_before = Pt(0)
titulo.paragraph_format.space_after  = Pt(4)
r = titulo.add_run("GUION DE PRESENTACIÓN")
set_font(r, 16, bold=True, color=NAVY)

subtitulo = doc.add_paragraph()
subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitulo.paragraph_format.space_after = Pt(2)
r = subtitulo.add_run("Caso 86 — Catálogo de Esquemas Tributarios SII 2025")
set_font(r, 12, color=GRAY)

expositor = doc.add_paragraph()
expositor.alignment = WD_ALIGN_PARAGRAPH.CENTER
expositor.paragraph_format.space_after = Pt(2)
r = expositor.add_run("Expositor: Nicolás Muñoz  ·  Grupo 4  ·  Magíster en Dirección Tributaria UVM")
set_font(r, 10.5, color=GRAY)

nota_slides = doc.add_paragraph()
nota_slides.alignment = WD_ALIGN_PARAGRAPH.CENTER
nota_slides.paragraph_format.space_after = Pt(16)
r = nota_slides.add_run("Slides a cargo: 1 — Portada  ·  2 — Caso 86  ·  7 — Alternativas de Solución")
set_font(r, 10, italic=True, color=ORANGE)

divider()

# ── SLIDE 1 ───────────────────────────────────────────────────────────────────
slide_label("SLIDE 1  —  PORTADA")
heading("Introducción")
speech(
    "Buenas tardes. Somos el Grupo 4 del Magíster en Dirección Tributaria de la "
    "Universidad Viña del Mar. Hoy presentamos el análisis del Caso 86 del Catálogo "
    "de Esquemas Tributarios del SII 2025, que aborda la interposición de una sociedad "
    "de profesionales con socios que no prestan servicios."
)

divider()

# ── SLIDE 2 ───────────────────────────────────────────────────────────────────
slide_label("SLIDE 2  —  CASO 86: ¿QUÉ ES?")
heading("Descripción del esquema")
speech(
    "El Caso 86 describe la situación de tres socios —A, B y C— que constituyen una "
    "sociedad de profesionales D, con participación igualitaria del 33,33% cada uno. "
    "El problema central es que solo A presta los servicios profesionales. B y C no "
    "tienen ningún rol activo ni aportan valor real a la sociedad. El efecto es que el "
    "ingreso que debería tributar íntegramente como renta de segunda categoría de A, "
    "se diluye en tres RUT distintos, reduciendo artificialmente su IGC."
)

divider()

# ── SLIDE 7 ───────────────────────────────────────────────────────────────────
slide_label("SLIDE 7  —  ALTERNATIVAS DE SOLUCIÓN")

heading("Introducción al bloque")
speech(
    "Una vez identificado el problema, la pregunta natural es: ¿cómo se organiza esto "
    "de manera legítima? El Caso 86 no prohíbe planificar ni usar estructuras societarias "
    "— lo que prohíbe es que la forma jurídica tenga como único propósito reducir la carga "
    "tributaria sin sustancia económica real detrás. Por eso presentamos cinco alternativas "
    "que logran eficiencia tributaria dentro del marco legal, analizando en cada una el "
    "impacto tanto en renta como en IVA, porque como verán, la elección del régimen no "
    "solo afecta cuánto paga el cliente en impuesto a la renta — también define si sus "
    "servicios quedan o no afectos al 19% de IVA."
)

heading("Alternativa 1 — Profesional independiente · Art. 42 N°2 LIR", color=NAVY)
speech(
    "A ejerce directamente como persona natural, emite boletas de honorarios y paga IGC "
    "sobre sus ingresos, pudiendo deducir gastos efectivos o la presunción legal del 30%, "
    "con tope de 15 UTA al año. La empresa que lo contrata retiene el 12,25% mensual, "
    "que se imputa al IGC anual. La ventaja adicional es que los servicios profesionales "
    "de personas naturales están expresamente exentos de IVA conforme al Art. 20 del "
    "DL 825, por lo que el cliente no paga el 19% adicional."
)

heading("Alternativa 2 — Sociedad de profesionales legítima · Art. 42 N°2 inc. 3°", color=NAVY)
speech(
    "Si B y C efectivamente trabajan en la sociedad, D califica bajo el Art. 42 N°2 y "
    "mantiene la exención de IVA mientras su giro sea exclusivamente profesional. En renta, "
    "puede optar por segunda categoría —donde los socios pagan IGC directo— o primera "
    "categoría con IDPC al 27% imputable como crédito al pagar el IGC."
)

heading("Alternativa 3 — Distribución proporcional al trabajo real", color=NAVY)
speech(
    "Se mantiene la misma sociedad, pero los estatutos reflejan el aporte real de cada "
    "socio —por ejemplo 90% para A— o se asigna un sueldo patronal a A previo al reparto. "
    "El tratamiento de IVA y renta es idéntico a la alternativa anterior, con la diferencia "
    "de que la carga tributaria queda correctamente radicada en quien genera el ingreso."
)

heading("Alternativa 4 — SpA · Régimen Pro Pyme · Art. 14 D N°3 LIR", color=NAVY)
speech(
    "La SpA paga IDPC al 25% y los socios tributan con IGC solo sobre los retiros efectivos, "
    "lo que permite diferir la tributación reinvirtiendo utilidades en la empresa. El punto "
    "crítico es que al salir del Art. 42 N°2 se pierde la exención de IVA: los servicios "
    "quedan afectos al 19%, lo que puede encarecer la oferta al cliente. Este punto hay que "
    "analizarlo siempre antes de recomendar esta estructura."
)

heading("Alternativa 5 — Régimen Pro Pyme Transparente · Art. 14 D N°8 LIR", color=NAVY)
speech(
    "La empresa no paga IDPC y los socios tributan directamente con IGC sobre su "
    "participación en los resultados del ejercicio, sin doble tributación. Sin embargo, "
    "también queda afecta al 19% de IVA igual que la SpA anterior. Esta alternativa "
    "conviene cuando los clientes son empresas que pueden usar ese IVA como crédito fiscal; "
    "si atienden a personas naturales, el IVA es un costo real que hay que evaluar."
)

heading("Cierre del bloque — Cuadro comparativo")
speech(
    "En resumen: las alternativas 1, 2 y 3 mantienen la exención de IVA y son las más "
    "eficientes cuando se atiende a personas naturales. Las alternativas 4 y 5 ofrecen "
    "mayor flexibilidad en renta, pero pierden esa exención. La recomendación del Grupo 4 "
    "es que no existe una alternativa universalmente mejor — la elección depende del perfil "
    "del cliente y de quién soporta el IVA en la cadena."
)

note("Tiempo estimado de este bloque: 8-10 minutos")

divider()

doc.save(OUTPUT)
print(f"✔  Guion generado: {OUTPUT}")
