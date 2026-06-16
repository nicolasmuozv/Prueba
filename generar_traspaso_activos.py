#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Planificación tributaria — Traspaso de activos Enciende Producciones SpA
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT = "/home/user/Prueba/Traspaso_Activos_Enciende.docx"

BLACK = RGBColor(0x1C, 0x1C, 0x1C)
NAVY  = RGBColor(0x1A, 0x3A, 0x5C)
GRAY  = RGBColor(0x55, 0x55, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xE8, 0x50, 0x0A)

doc = Document()

for sec in doc.sections:
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin   = Cm(3.0)
    sec.right_margin  = Cm(2.5)

def sf(run, size, bold=False, italic=False, color=BLACK, name="Times New Roman"):
    run.font.name  = name
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = color

def shade_para(p, hex_color):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)

def titulo_principal(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(6)
    r = p.add_run(text)
    sf(r, 16, bold=True, color=NAVY)

def subtitulo(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    sf(r, 12, color=GRAY)

def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after  = Pt(4)
    shade_para(p, "1A3A5C")
    r = p.add_run(f"  {text}")
    sf(r, 12, bold=True, color=WHITE, name="Calibri")

def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(3)
    r = p.add_run(text)
    sf(r, 12, bold=True, color=NAVY)

def h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(text)
    sf(r, 11, bold=True, color=ORANGE)

def body(text, justify=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(6)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    sf(r, 11)

def bullet(text, nivel=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent   = Cm(0.5 + nivel * 0.5)
    p.paragraph_format.space_before  = Pt(1)
    p.paragraph_format.space_after   = Pt(2)
    r = p.add_run(text)
    sf(r, 11)

def nota_pie(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Cm(0.5)
    r = p.add_run(text)
    sf(r, 9, italic=True, color=GRAY)

def divider():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run("─" * 85)
    sf(r, 7, color=RGBColor(0xCC, 0xCC, 0xCC))

def add_table(headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        run = hdr_cells[i].paragraphs[0].runs[0]
        sf(run, 10, bold=True, color=WHITE, name="Calibri")
        tc = hdr_cells[i]._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1A3A5C')
        tcPr.append(shd)
    for r_idx, row in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        for c_idx, cell_text in enumerate(row):
            row_cells[c_idx].text = cell_text
            run = row_cells[c_idx].paragraphs[0].runs[0]
            sf(run, 10, name="Calibri")
    if col_widths:
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                cell.width = Cm(col_widths[i])
    doc.add_paragraph()

# ─── PORTADA ─────────────────────────────────────────────────────────────────
titulo_principal("PLANIFICACIÓN TRIBUTARIA")
titulo_principal("TRASPASO DE ACTIVOS ENTRE SOCIEDADES")
subtitulo("Enciende Producciones SpA → Nueva Sociedad")
subtitulo("Mismo grupo de accionistas · Sin tributación")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(20)
r = p.add_run("Elaborado por: Nicolás Muñoz Valdebenito  ·  Magíster en Dirección Tributaria — UVM")
sf(r, 10, italic=True, color=GRAY)

divider()

# ─── I. OBJETIVO ─────────────────────────────────────────────────────────────
h1("I.  OBJETIVO Y PREMISA")
body(
    "El presente informe analiza las implicancias tributarias y la vía óptima para traspasar activos "
    "desde Enciende Producciones SpA (RUT 77.019.731-7) a una nueva sociedad que tendrá idénticos "
    "accionistas: Diego Andrés Niño Carrasco (50%) y Rodrigo Alejandro Niño Carrasco (50%). "
    "El objetivo es ejecutar la reorganización sin generar renta imponible ni IVA, aprovechando "
    "las normas de reorganización empresarial del artículo 64 del Código Tributario."
)

# ─── II. MARCO NORMATIVO ─────────────────────────────────────────────────────
h1("II.  MARCO NORMATIVO APLICABLE")

h2("A)  Artículo 64 Código Tributario — La regla central")
body(
    "El artículo 64 inciso 4° y 5° CT establece la excepción a la tasación del SII cuando se "
    "cumplan copulativamente los siguientes requisitos:"
)
add_table(
    ["Requisito", "Descripción"],
    [
        ["1. Reorganización empresarial", "División, fusión, transformación o aporte dentro de un proceso de reorganización con legítima razón de negocios."],
        ["2. Mismo grupo empresarial", "Los activos deben quedar dentro del mismo grupo (mismos dueños en misma proporción)."],
        ["3. Sin flujo de caja", "No puede haber pago de precio en dinero — el traspaso es a cambio de participación en la nueva sociedad."],
        ["4. Valor tributario", "Los activos deben registrarse en la nueva entidad al mismo valor tributario que tenían en la sociedad original."],
        ["5. Legítima razón de negocios", "Debe existir una justificación económica real, no solo tributaria (Circular N°45/2001, N°17/2015 SII)."],
    ],
    col_widths=[4.5, 11.0]
)
body(
    "Si se cumplen estas condiciones, el SII no puede tasar la operación y el traspaso no genera "
    "Impuesto de Primera Categoría (IDPC) ni Impuesto Global Complementario (IGC)."
)

# ─── III. LAS TRES VÍAS ───────────────────────────────────────────────────────
h1("III.  VÍAS DISPONIBLES PARA ENCIENDE")

h2("Vía 1 — División de Sociedad  (Recomendada)")
body(
    "Base legal: Arts. 94 a 100 Ley N°18.046 (aplicable a SpA por Art. 424 Código de Comercio), "
    "Art. 64 CT."
)
body(
    "Enciende SpA se divide en dos entidades: (1) Enciende SpA, que mantiene las operaciones y el "
    "giro habitual, y (2) Nueva SpA, que recibe los activos seleccionados. Los accionistas conservan "
    "exactamente la misma proporción (50%/50%) en ambas sociedades. Los activos se traspasan al valor "
    "tributario neto registrado en Enciende, sin flujo de dinero."
)
body("Ventajas tributarias de esta vía:")
bullet("No hay enajenación → no hay renta → cero IDPC y cero IGC.")
bullet("No hay venta → cero IVA.")
bullet("El SII no puede tasar (Art. 64 CT) al cumplirse todos los requisitos.")
bullet("Las pérdidas tributarias de Enciende se distribuyen proporcionalmente entre ambas entidades.")
body(
    "Limitación relevante: la nueva sociedad hereda el valor tributario de los activos. Los equipos "
    "de sonido, iluminación y estructuras que fueron cargados a gasto al adquirirse no existen "
    "contablemente como activos — no pueden traspasarse por esta vía sin riesgo tributario."
)

h2("Vía 2 — Aporte de Activos a Sociedad Nueva  (Art. 64 CT)")
body(
    "Enciende crea una nueva SpA y aporta activos específicos como capital, recibiendo acciones de "
    "la nueva empresa. El aporte se valoriza al valor tributario de los activos (no al comercial). "
    "A diferencia de la División, en este caso Enciende queda como accionista (matriz) de la nueva "
    "sociedad. Conviene cuando el objetivo es mantener una estructura holding o traspasar solo "
    "algunos activos específicos, no dividir la totalidad del patrimonio."
)
body(
    "Riesgo adicional para Enciende: dado el Control Precautorio SII por IVA ($19,05M) y los flujos "
    "sin respaldo ($280,5M), un aporte puede ser revisado con mayor intensidad que una división "
    "formal, que tiene mayor protección procesal."
)

h2("Vía 3 — Fusión por Absorción  (No recomendada)")
body(
    "No aplica al objetivo del caso, ya que la fusión busca consolidar entidades, no separarlas. "
    "Se descarta."
)

# ─── IV. ACTIVOS A TRASPASAR ──────────────────────────────────────────────────
h1("IV.  ACTIVOS QUE PUEDEN TRASPASARSE")

h2("4.1  Activos registrados en balance — aptos para el traspaso")
add_table(
    ["Activo", "Valor Tributario", "Observación"],
    [
        ["Vehículos", "$15.200.000", "Verificar si el crédito asociado es deuda personal del socio o de la empresa."],
        ["Instrumentos musicales", "$1.800.000", "Apto para traspasar."],
        ["Muebles y enseres", "$170.000", "Apto para traspasar."],
        ["Equipos computacionales (2026)", "$7.000.000", "Apto para traspasar."],
        ["Muebles 2026", "$700.000", "Apto para traspasar."],
        ["TOTAL REGISTRADO", "~$24.870.000", ""],
    ],
    col_widths=[5.5, 4.0, 6.0]
)

h2("4.2  Activos no registrados — requieren tratamiento previo")
body(
    "Los equipos de sonido, iluminación y estructuras que fueron cargados a gasto al adquirirse "
    "no figuran en el activo fijo de Enciende. Para incluirlos en el traspaso existiría que revertir "
    "el gasto (alto riesgo tributario — SII puede impugnar) o hacer un inventario a valor comercial. "
    "Este segundo camino genera IVA y posiblemente IDPC, perdiendo el beneficio del Art. 64 CT."
)
body(
    "Recomendación: levantar un cardex completo y valorizado antes de ejecutar la reorganización. "
    "Sin este inventario no hay base cierta para fijar el valor de aporte ni para distribuir "
    "correctamente los pasivos."
)

# ─── V. CONTINGENCIAS ─────────────────────────────────────────────────────────
h1("V.  CONTINGENCIAS QUE DEBEN RESOLVERSE ANTES DEL TRASPASO")

body(
    "Las siguientes situaciones pueden invalidar el beneficio del Art. 64 CT, exponer a impugnación "
    "por parte del SII, o simplemente bloquear los trámites de la reorganización:"
)

h3("⚠  1. Flujos a relacionados sin contrato ($280.500.000)")
body(
    "Pagos a Ana María Lira ($215M), Espíritu Tecnológico ($21M), GH Setups ($13M), "
    "Inversiones A y D ($9M) y Los Acacios ($22M) no tienen contrato ni respaldo suficiente. "
    "El SII puede recaracterizarlos como retiros encubiertos (Art. 21 LIR), generando un impuesto "
    "único del 40% más recargos. Una reorganización ejecutada con esta contingencia activa puede "
    "ser vista como un mecanismo para evadir. Acción requerida: documentar contratos, "
    "prestaciones y pagos antes de reorganizar."
)

h3("⚠  2. Control Precautorio SII por IVA ($19.050.000)")
body(
    "El Control Precautorio puede bloquear el timbraje de documentos tributarios. La nueva sociedad "
    "requiere RUT, inicio de actividades y timbraje propio — si Enciende tiene bloqueos activos, "
    "esto puede complicarse o demorar el proceso. Acción requerida: resolver el Control Precautorio "
    "antes de iniciar el proceso de división."
)

h3("⚠  3. Préstamo $27.000.000 de Consuelo Quinteros")
body(
    "Esta deuda no está reflejada en el balance. Si en la división se traspasan activos sin reconocer "
    "el pasivo correspondiente, el SII podría objetar que se está transfiriendo valor neto sin "
    "distribuir correctamente las obligaciones. Acción requerida: formalizar con contrato de "
    "mutuo y registrar contablemente antes de la división."
)

h3("⚠  4. IVA postergado ($14.470.000 — vencido 20-03-2026)")
body(
    "Una deuda tributaria vencida puede generar notificaciones de incumplimiento que dificultan "
    "trámites ante el SII (inicio de actividades, modificaciones de sociedad). "
    "Acción requerida: regularizar el pago o negociar un convenio de pago con el SII."
)

# ─── VI. TRATAMIENTO IVA ──────────────────────────────────────────────────────
h1("VI.  TRATAMIENTO IVA DEL TRASPASO")

add_table(
    ["Vía de traspaso", "¿Genera IVA?", "Fundamento"],
    [
        ["División de sociedad", "No", "No hay venta ni transferencia onerosa — es reorganización empresarial."],
        ["Aporte al valor tributario (Art. 64 CT)", "No, si cumple requisitos", "Circular N°44/2006 SII; Art. 8 DL 825: el aporte en reorganización no es hecho gravado si no hay flujo de caja y mismos dueños."],
        ["Venta a precio de mercado", "Sí, 19%", "Art. 8 letra d) DL 825."],
        ["Aporte a valor comercial distinto del tributario", "Sí, por la diferencia", "El SII tasa conforme Art. 64 CT."],
    ],
    col_widths=[5.0, 3.5, 7.0]
)

# ─── VII. RUTA RECOMENDADA ────────────────────────────────────────────────────
h1("VII.  RUTA RECOMENDADA: DIVISIÓN DE SOCIEDAD")

h2("Paso a paso")

h3("Etapa 1 — Preparación (estimado 2-3 meses)")
bullet("Regularizar IVA postergado ($14,47M) o negociar convenio de pago con SII.")
bullet("Formalizar préstamo de Consuelo Quinteros ($27M) con contrato de mutuo y registro contable.")
bullet("Documentar contratos con todas las partes relacionadas ($280,5M).")
bullet("Levantar cardex completo de activos fijos (incluir equipos no registrados).")
bullet("Resolver Control Precautorio SII antes de iniciar trámites de reorganización.")

h3("Etapa 2 — Diseño de la División")
bullet("Definir qué activos van a la nueva sociedad y cuáles quedan en Enciende.")
bullet("Determinar el patrimonio neto a transferir (activos menos pasivos proporcionales).")
bullet("Confirmar que ambas sociedades quedan con participación 50%/50% (misma proporción).")
bullet("Establecer el giro de la nueva sociedad y sus actividades.")

h3("Etapa 3 — Acto Jurídico")
bullet("Junta de accionistas de Enciende SpA aprueba la división (requiere mayoría calificada).")
bullet("Escritura pública de división ante notario.")
bullet("Balance de división auditado o certificado por contador.")
bullet("Inscripción en Registro de Comercio y publicación en Diario Oficial.")

h3("Etapa 4 — Trámites SII")
bullet("Obtención de RUT para la nueva sociedad.")
bullet("Inicio de actividades nueva SpA ante SII.")
bullet("Timbraje de documentos tributarios.")
bullet("Aviso de modificación al registro de Enciende SpA.")

h3("Etapa 5 — Post-División")
bullet("Registrar activos en nueva sociedad al mismo valor tributario que tenían en Enciende.")
bullet("Distribuir las pérdidas tributarias proporcionalmente entre ambas entidades.")
bullet("Actualizar contratos de trabajo, proveedores y clientes según corresponda.")

# ─── VIII. CONCLUSIÓN ─────────────────────────────────────────────────────────
h1("VIII.  CONCLUSIÓN")
body(
    "La División de Sociedad es la única vía que garantiza tributación cero en el traspaso de "
    "activos, dado que cumple todos los requisitos del artículo 64 del Código Tributario: mismos "
    "dueños, misma proporción, valor tributario, sin flujo de caja y reorganización con legítima "
    "razón de negocios. Adicionalmente, no genera hecho gravado con IVA y tiene mayor solidez "
    "jurídica ante una eventual fiscalización del SII."
)
body(
    "Sin embargo, Enciende Producciones SpA presenta cuatro contingencias activas que deben "
    "resolverse antes de ejecutar la División: los flujos a relacionados sin contrato, el Control "
    "Precautorio por IVA, el préstamo no registrado de Consuelo Quinteros, y el IVA postergado "
    "vencido. Ejecutar la reorganización sin resolver estas contingencias expone a que el SII "
    "impugne la operación o bloquee los trámites de la nueva sociedad."
)
body(
    "La recomendación es avanzar en paralelo en la resolución de contingencias y el diseño de "
    "la División, con el objetivo de ejecutar el acto jurídico una vez que el escenario esté "
    "saneado — estimado en 2 a 3 meses."
)

# ─── NORMATIVA ────────────────────────────────────────────────────────────────
divider()
nota_pie(
    "Normativa consultada: Art. 64 Código Tributario; Art. 14 D N°3 DL 824 (LIR); Art. 21 LIR; "
    "Art. 8 DL 825; Circular SII N°45/2001; Circular SII N°17/2015; Circular SII N°44/2006; "
    "Ley N°18.046 Arts. 94-100; Art. 424 Código de Comercio."
)

doc.save(OUTPUT)
print(f"✔  Documento generado: {OUTPUT}")
