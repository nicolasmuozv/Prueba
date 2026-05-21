#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Toma el PPT base de la compañera (8 slides) y completa las slides 7 y 8
con contenido de alternativas y conclusiones, respetando el estilo exacto.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from lxml import etree
import copy

INPUT  = "/home/user/Prueba/base_companera.pptx"
OUTPUT = "/home/user/Prueba/Caso86_Presentacion.pptx"

prs = Presentation(INPUT)

NSMAP = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def emu(inches):
    return int(inches * 914400)

def sp_accent1(sid, x, y, w, h, lines, sz=1100, anchor="ctr", bold_first=False):
    """Rectángulo con relleno accent1 (azul del tema) y texto blanco — igual a slide 6."""
    lines_xml = ""
    for i, line in enumerate(lines):
        b_open  = '<a:b>1</a:b>' if (i == 0 and bold_first) else ''
        lines_xml += f'''
    <a:p>
      <a:r>
        <a:rPr lang="es-CL" sz="{sz}" dirty="0">{b_open}</a:rPr>
        <a:t>{_esc(line)}</a:t>
      </a:r>
    </a:p>'''
    return etree.fromstring(f'''<p:sp
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="Rect {sid}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{emu(x)}" y="{emu(y)}"/>
      <a:ext cx="{emu(w)}" cy="{emu(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="{anchor}"/>
    <a:lstStyle/>
    {lines_xml}
  </p:txBody>
</p:sp>''')

def sp_dark_panel(sid, x, y, w, h):
    """Panel oscuro con degradado — igual a Rectángulo 11 de slide 6."""
    return etree.fromstring(f'''<p:sp
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="PanelOscuro {sid}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{emu(x)}" y="{emu(y)}"/>
      <a:ext cx="{emu(w)}" cy="{emu(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="73000">
          <a:schemeClr val="dk1"><a:tint val="60000"/><a:satMod val="105000"/><a:lumMod val="105000"/></a:schemeClr>
        </a:gs>
        <a:gs pos="100000">
          <a:schemeClr val="dk1"><a:tint val="65000"/><a:satMod val="100000"/><a:lumMod val="100000"/></a:schemeClr>
        </a:gs>
        <a:gs pos="100000">
          <a:schemeClr val="dk1"><a:tint val="70000"/><a:satMod val="100000"/><a:lumMod val="100000"/></a:schemeClr>
        </a:gs>
      </a:gsLst>
    </a:gradFill>
  </p:spPr>
  <p:style>
    <a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef>
    <a:fillRef idx="2"><a:schemeClr val="dk1"/></a:fillRef>
    <a:effectRef idx="1"><a:schemeClr val="dk1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:endParaRPr lang="es-CL" dirty="0"/></a:p>
  </p:txBody>
</p:sp>''')

def tb_text(sid, x, y, w, h, lines, sz=1100, color_hex=None, bold=False, align="l"):
    """Cuadro de texto simple con color de fuente explícito."""
    runs_xml = ""
    for i, line in enumerate(lines):
        clr = f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>' if color_hex else \
              '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>'
        b_tag = '<a:b>1</a:b>' if bold else ''
        runs_xml += f'''
    <a:p>
      <a:pPr algn="{align}"/>
      <a:r>
        <a:rPr lang="es-CL" sz="{sz}" dirty="0">{clr}{b_tag}</a:rPr>
        <a:t>{_esc(line)}</a:t>
      </a:r>
    </a:p>'''
    return etree.fromstring(f'''<p:sp
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="TB {sid}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{emu(x)}" y="{emu(y)}"/>
      <a:ext cx="{emu(w)}" cy="{emu(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"/>
    <a:lstStyle/>
    {runs_xml}
  </p:txBody>
</p:sp>''')

def _esc(s):
    return (s.replace('&','&amp;')
             .replace('<','&lt;')
             .replace('>','&gt;')
             .replace('"','&quot;'))

def add_to_slide(slide, *elements):
    spTree = slide._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld'
                                  '/{http://schemas.openxmlformats.org/presentationml/2006/main}spTree'
                                  if False else
                                  './/{http://schemas.openxmlformats.org/drawingml/2006/main}grpSpPr').getparent()
    for el in elements:
        spTree.append(el)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Alternativas de solución
# ═══════════════════════════════════════════════════════════════════════════════
s7 = prs.slides[6]
spTree7 = s7.shapes._spTree

# Dimensiones: área izquierda 0.65–7.68in, panel derecho 7.8–12.68in
LX, LW = 0.65, 7.03    # área de contenido izquierda
PX, PW = 7.80, 4.88    # panel oscuro derecho
PY, PH = 1.49, 5.52    # panel vertical

# 1. Header (accent1 oscuro, color tx1 semitransparente como "Normativa Base")
header7 = etree.fromstring(f'''<p:sp
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="100" name="Header7"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{emu(LX)}" y="{emu(1.49)}"/><a:ext cx="{emu(LW)}" cy="{emu(0.55)}"/></a:xfrm>
    <a:solidFill><a:schemeClr val="tx1"><a:alpha val="74000"/></a:schemeClr></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="es-CL" dirty="0">
          <a:solidFill><a:srgbClr val="FFC000"/></a:solidFill>
        </a:rPr>
        <a:t>Alternativas de solución</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>''')
spTree7.append(header7)

# 2. Panel oscuro derecho
spTree7.append(sp_dark_panel(101, PX, PY, PW, PH))

# 3. Panel derecho: contenidos (title + texto)
spTree7.append(tb_text(102, PX+0.12, PY+0.10, PW-0.24, 0.45,
    ["Consideraciones importantes"],
    sz=1300, color_hex="FFC000", bold=True, align="ctr"))

spTree7.append(tb_text(103, PX+0.12, PY+0.60, PW-0.24, 2.30,
    [
        "Para que cualquier alternativa sea",
        "válida frente a las NGA, debe poder",
        "acreditar SUSTANCIA ECONÓMICA REAL:",
        "",
        "▪  Todos los socios deben ejercer",
        "    su profesión para la sociedad.",
        "",
        "▪  La distribución de utilidades debe",
        "    reflejar el aporte efectivo.",
        "",
        "▪  El capital debe ser relevante en",
        "    relación a los ingresos.",
    ],
    sz=1000, color_hex="FFFFFF", align="l"))

spTree7.append(tb_text(104, PX+0.12, PY+3.05, PW-0.24, 0.45,
    ["Normativa de referencia"],
    sz=1200, color_hex="FFC000", bold=True, align="ctr"))

spTree7.append(tb_text(105, PX+0.12, PY+3.55, PW-0.24, 1.85,
    [
        "▪  Art. 42 N°2 LIR — rentas 2ª cat.",
        "▪  Art. 14 D N°3 y N°8 LIR — regímenes",
        "▪  Arts. 4 bis y 4 ter CT — NGA",
        "▪  Circ. N°21/1991 y N°50/2020 SII",
        "▪  Circ. N°65/2015 SII",
    ],
    sz=1000, color_hex="FFFFFF", align="l"))

# 4. Las 5 alternativas (accent1, apiladas verticalmente)
alternativas = [
    ("Alt. 1 — Profesional independiente (Art. 42 N°2 LIR)",
     ["A ejerce como persona natural. Emite boletas de honorarios. Gastos efectivos",
      "o presuntos (30%, tope 15 UTA). Paga IGC sobre el total. Sin riesgo NGA."]),
    ("Alt. 2 — Sociedad de profesionales con socios que trabajen",
     ["A se asocia con profesionales que efectivamente presten servicios",
      "para la sociedad. D califica legítimamente bajo Art. 42 N°2 inc. 3° LIR."]),
    ("Alt. 3 — Distribución de utilidades proporcional al trabajo real",
     ["Estatutos contemplan participación diferenciada acorde al aporte efectivo",
      "o sueldo patronal previo al reparto. Debe reflejar la realidad económica."]),
    ("Alt. 4 — SpA — régimen Pro Pyme General (Art. 14 D N°3 LIR)",
     ["SpA con IDPC al 25%. Permite planificación de retiros y reinversión.",
      "Sin restricciones de la sociedad de profesionales. Exige sustancia real."]),
    ("Alt. 5 — Régimen Pro Pyme Transparente (Art. 14 D N°8 LIR)",
     ["La empresa no paga IDPC; socios tributan con IGC. SOLO válido si B y C",
      "tienen sustancia económica real — de lo contrario el SII aplica NGA."]),
]

box_h   = 0.82
gap     = 0.06
y_start = 2.10
for i, (titulo, desc) in enumerate(alternativas):
    y = y_start + i * (box_h + gap)
    # Título: rectángulo accent1 (altura 0.35)
    spTree7.append(sp_accent1(110+i*2, LX, y, LW, 0.35, [titulo], sz=1100, bold_first=True))
    # Descripción: rectángulo accent1 más claro (misma familia, pero lo hacemos con tb sin fondo)
    spTree7.append(etree.fromstring(f'''<p:sp
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{111+i*2}" name="AltDesc{i}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{emu(LX)}" y="{emu(y+0.35)}"/><a:ext cx="{emu(LW)}" cy="{emu(box_h-0.35)}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>
    <a:fillRef idx="2"><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    {''.join(f'<a:p><a:r><a:rPr lang="es-CL" sz="950" dirty="0"/><a:t>{_esc(d)}</a:t></a:r></a:p>' for d in desc)}
  </p:txBody>
</p:sp>'''))


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Conclusiones
# ═══════════════════════════════════════════════════════════════════════════════
s8 = prs.slides[7]
spTree8 = s8.shapes._spTree

# 1. Header estilo "Normativa Base"
header8 = etree.fromstring(f'''<p:sp
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="200" name="Header8"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{emu(LX)}" y="{emu(1.49)}"/><a:ext cx="{emu(LW)}" cy="{emu(0.55)}"/></a:xfrm>
    <a:solidFill><a:schemeClr val="tx1"><a:alpha val="74000"/></a:schemeClr></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="es-CL" dirty="0">
          <a:solidFill><a:srgbClr val="FFC000"/></a:solidFill>
        </a:rPr>
        <a:t>Síntesis del análisis</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>''')
spTree8.append(header8)

# 2. Panel oscuro derecho
spTree8.append(sp_dark_panel(201, PX, PY, PW, PH))

# 3. Panel derecho: posición del grupo + normativa
spTree8.append(tb_text(202, PX+0.12, PY+0.10, PW-0.24, 0.45,
    ["Posición del Grupo 4"],
    sz=1300, color_hex="FFC000", bold=True, align="ctr"))

spTree8.append(tb_text(203, PX+0.12, PY+0.60, PW-0.24, 2.5,
    [
        "El Caso 86 configura ABUSO de",
        "las formas jurídicas (Art. 4 ter CT).",
        "",
        "NO es simulación: los actos son",
        "reales, pero la forma jurídica es",
        "inapropiada para el sustrato",
        "económico del caso.",
        "",
        "Integrantes:",
        "▪  Andrea Añasco",
        "▪  Gema Sepúlveda",
        "▪  Karen Rebolledo",
        "▪  Nicolás Muñoz",
    ],
    sz=1000, color_hex="FFFFFF", align="l"))

spTree8.append(tb_text(204, PX+0.12, PY+3.25, PW-0.24, 0.45,
    ["Norma aplicable"],
    sz=1200, color_hex="FFC000", bold=True, align="ctr"))

spTree8.append(tb_text(205, PX+0.12, PY+3.75, PW-0.24, 1.65,
    [
        "▪  Art. 4 ter CT — ABUSO",
        "▪  Art. 4 quinquies CT — Procedimiento",
        "▪  Art. 42 N°2 LIR",
        "▪  Circ. N°21/1991 y N°50/2020 SII",
        "▪  Ley N°20.780 y N°21.210",
    ],
    sz=1000, color_hex="FFFFFF", align="l"))

# 4. Las 5 conclusiones (accent1)
conclusiones = [
    ("ABUSO de formas jurídicas — Art. 4 ter CT",
     ["El esquema utiliza una forma jurídica inapropiada para el sustrato económico.",
      "Los actos son reales pero el único efecto relevante es tributario."]),
    ("La sociedad D NO califica como sociedad de profesionales",
     ["B y C no ejercen su profesión para la sociedad, incumpliendo requisito",
      "copulativo de las Circulares N°21/1991 y N°50/2020 del SII."]),
    ("El SII puede requerir declaración de abuso al TTA",
     ["Art. 4 quinquies CT: el Director puede solicitar al Tribunal Tributario y",
      "Aduanero la declaración de abuso → recalificación + intereses + multas."]),
    ("Existen alternativas legítimas de organización",
     ["Profesional independiente, sociedad de profesionales con socios reales,",
      "SpA Pro Pyme (14 D N°3) o distribución proporcional al trabajo real."]),
    ("La clave es la sustancia económica",
     ["Toda estructura debe poder demostrar que los socios aportan trabajo, capital,",
      "clientela o dirección real. La forma jurídica por sí sola no es suficiente."]),
]

y_start = 2.10
for i, (titulo, desc) in enumerate(conclusiones):
    y = y_start + i * (box_h + gap)
    spTree8.append(sp_accent1(210+i*2, LX, y, LW, 0.35, [titulo], sz=1100, bold_first=True))
    spTree8.append(etree.fromstring(f'''<p:sp
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{211+i*2}" name="ConcDesc{i}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{emu(LX)}" y="{emu(y+0.35)}"/><a:ext cx="{emu(LW)}" cy="{emu(box_h-0.35)}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>
    <a:fillRef idx="2"><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    {''.join(f'<a:p><a:r><a:rPr lang="es-CL" sz="950" dirty="0"/><a:t>{_esc(d)}</a:t></a:r></a:p>' for d in desc)}
  </p:txBody>
</p:sp>'''))


prs.save(OUTPUT)
print(f"✔ PPT generado: {OUTPUT}")
print(f"  Slides: {len(prs.slides)}")
