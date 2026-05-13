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
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus.flowables import HRFlowable
import datetime

# ── Configuración general ──────────────────────────────────────────────────────
OUTPUT = "/home/user/Prueba/Asesoria_Tributaria_Inmobiliaria.pdf"
AZUL   = colors.HexColor("#1A3A5C")
CELESTE= colors.HexColor("#2E6DA4")
GRIS   = colors.HexColor("#F4F6F9")
NARANJA= colors.HexColor("#D4581F")
NEGRO  = colors.HexColor("#1C1C1C")
VERDE  = colors.HexColor("#1A6B3C")
ROJO   = colors.HexColor("#C0392B")

doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=letter,
    rightMargin=2.2*cm, leftMargin=2.2*cm,
    topMargin=2.5*cm,   bottomMargin=2.2*cm,
    title="Asesoría Tributaria – Inversión Inmobiliaria"
)

# ── Estilos ────────────────────────────────────────────────────────────────────
st = getSampleStyleSheet()

def make_style(name, parent="Normal", **kw):
    s = ParagraphStyle(name, parent=st[parent], **kw)
    return s

titulo_doc   = make_style("TituloDoc",   fontSize=22, textColor=colors.white,
                           alignment=TA_CENTER, leading=28, fontName="Helvetica-Bold")
subtitulo_doc= make_style("SubtituloDoc",fontSize=13, textColor=colors.white,
                           alignment=TA_CENTER, leading=18, fontName="Helvetica")
fecha_doc    = make_style("FechaDoc",    fontSize=10, textColor=colors.HexColor("#BDD7EE"),
                           alignment=TA_CENTER, fontName="Helvetica-Oblique")

titulo_sec   = make_style("TituloSec",   fontSize=14, textColor=AZUL,
                           fontName="Helvetica-Bold", spaceBefore=14, spaceAfter=4, leading=18)
titulo_sub   = make_style("TituloSub",   fontSize=11, textColor=CELESTE,
                           fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=3, leading=15)
cuerpo       = make_style("Cuerpo",      fontSize=9.5, textColor=NEGRO,
                           leading=14, spaceAfter=5, alignment=TA_JUSTIFY)
viñeta       = make_style("Viñeta",      fontSize=9.5, textColor=NEGRO,
                           leading=13, leftIndent=16, spaceAfter=3)
nota_legal   = make_style("NotaLegal",   fontSize=8.5, textColor=colors.HexColor("#555555"),
                           leading=12, leftIndent=12, fontName="Helvetica-Oblique")
alerta       = make_style("Alerta",      fontSize=9.5, textColor=ROJO,
                           fontName="Helvetica-Bold", leading=13, leftIndent=12, spaceAfter=4)
positivo     = make_style("Positivo",    fontSize=9.5, textColor=VERDE,
                           fontName="Helvetica-Bold", leading=13, leftIndent=12, spaceAfter=4)
tabla_titulo = make_style("TablaTitulo", fontSize=9, textColor=colors.white,
                           fontName="Helvetica-Bold", alignment=TA_CENTER, leading=12)
tabla_cuerpo = make_style("TablaCuerpo", fontSize=8.5, textColor=NEGRO,
                           leading=12, alignment=TA_LEFT)
disclaimer   = make_style("Disclaimer",  fontSize=8, textColor=colors.HexColor("#777777"),
                           leading=11, alignment=TA_JUSTIFY, fontName="Helvetica-Oblique")

# ── Utilidades ─────────────────────────────────────────────────────────────────
def hr(color=CELESTE, width=1):
    return HRFlowable(width="100%", thickness=width, color=color, spaceAfter=6, spaceBefore=6)

def seccion_titulo(texto):
    return [
        Spacer(1, 0.3*cm),
        Table([[Paragraph(texto, titulo_sec)]],
              colWidths=["100%"],
              style=TableStyle([
                  ("BACKGROUND", (0,0), (-1,-1), GRIS),
                  ("LEFTPADDING", (0,0), (-1,-1), 10),
                  ("RIGHTPADDING",(0,0), (-1,-1), 10),
                  ("TOPPADDING",  (0,0), (-1,-1), 8),
                  ("BOTTOMPADDING",(0,0),(-1,-1), 8),
                  ("LINEBELOW", (0,0), (-1,-1), 2, CELESTE),
              ])),
        Spacer(1, 0.2*cm),
    ]

def cuadro_resumen(titulo, filas):
    """filas: lista de (clave, valor)"""
    data = [[Paragraph(titulo, tabla_titulo), ""]]
    for k, v in filas:
        data.append([Paragraph(k, tabla_cuerpo), Paragraph(v, tabla_cuerpo)])
    t = Table(data, colWidths=[6*cm, 11.8*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), AZUL),
        ("SPAN",         (0,0), (-1,0)),
        ("BACKGROUND",   (0,1), (-1,-1), colors.white),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, GRIS]),
        ("FONTNAME",     (0,1), (0,-1), "Helvetica-Bold"),
        ("FONTSIZE",     (0,1), (-1,-1), 8.5),
        ("GRID",         (0,0), (-1,-1), 0.4, colors.HexColor("#CCCCCC")),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING",   (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("VALIGN",       (0,0), (-1,-1), "TOP"),
    ]))
    return [t, Spacer(1, 0.3*cm)]

def cuadro_comparativo(titulo, headers, filas):
    data = [[Paragraph(h, tabla_titulo) for h in headers]]
    for fila in filas:
        data.append([Paragraph(c, tabla_cuerpo) for c in fila])
    n = len(headers)
    ancho = 17.8*cm / n
    t = Table(data, colWidths=[ancho]*n)
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), AZUL),
        ("BACKGROUND",   (0,1), (-1,-1), colors.white),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, GRIS]),
        ("FONTSIZE",     (0,0), (-1,-1), 8.5),
        ("GRID",         (0,0), (-1,-1), 0.4, colors.HexColor("#CCCCCC")),
        ("LEFTPADDING",  (0,0), (-1,-1), 7),
        ("RIGHTPADDING", (0,0), (-1,-1), 7),
        ("TOPPADDING",   (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("VALIGN",       (0,0), (-1,-1), "TOP"),
    ]))
    return [Paragraph(f"<b>{titulo}</b>", titulo_sub), t, Spacer(1, 0.3*cm)]

# ── Construcción del documento ──────────────────────────────────────────────────
story = []

# ─────────────────────── PORTADA ────────────────────────────────────────────────
portada = Table([
    [Paragraph("INFORME DE ASESORÍA TRIBUTARIA", titulo_doc)],
    [Spacer(1, 0.3*cm)],
    [Paragraph("Inversión en Bienes Inmuebles para Arrendamiento", subtitulo_doc)],
    [Spacer(1, 0.5*cm)],
    [Paragraph("Persona Natural · Art. 27 Bis · Empresa Inmobiliaria · IVA", subtitulo_doc)],
    [Spacer(1, 0.8*cm)],
    [Paragraph(f"Fecha: {datetime.date.today().strftime('%d de %B de %Y')}", fecha_doc)],
    [Paragraph("Elaborado en base al material del Magíster en Dirección Tributaria – UVM", fecha_doc)],
], colWidths=["100%"])
portada.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,-1), AZUL),
    ("TOPPADDING",    (0,0), (-1,-1), 18),
    ("BOTTOMPADDING", (0,0), (-1,-1), 18),
    ("LEFTPADDING",   (0,0), (-1,-1), 20),
    ("RIGHTPADDING",  (0,0), (-1,-1), 20),
    ("ROUNDEDCORNERS",(0,0), (-1,-1), [8,8,8,8]),
]))
story += [portada, Spacer(1, 0.6*cm)]

# ─────────────────────── RESUMEN EJECUTIVO ────────────────────────────────────
story += seccion_titulo("RESUMEN EJECUTIVO")
story.append(Paragraph(
    "El presente informe analiza la situación tributaria de un cliente persona natural que proyecta "
    "adquirir uno o más bienes raíces para su posterior arrendamiento. Se abordan los principales "
    "impactos en IVA, Impuesto a la Renta, Ingresos No Constitutivos de Renta (INR), la franquicia "
    "del artículo 27 bis del DL 825 (devolución e imputación de IVA activo fijo), y los efectos "
    "que tendría constituir una sociedad para llevar adelante este negocio.", cuerpo))
story.append(Spacer(1, 0.2*cm))

story += cuadro_resumen("TEMAS CUBIERTOS EN ESTE INFORME", [
    ("1. Compra del inmueble",        "IVA en la adquisición, habitualidad del vendedor, crédito fiscal."),
    ("2. Arrendamiento como PN",      "Tributación en Renta, INR (DFL2), IVA si inmueble amoblado."),
    ("3. Ingresos No Renta (INR)",    "Art. 17 N°8 LIR: mayor valor exento hasta 8.000 UF. DFL2: hasta 2 propiedades."),
    ("4. Art. 27 bis – Devolución",   "Recuperación de remanente CF IVA por activo fijo (desde 2 períodos)."),
    ("5. Art. 27 bis – Imputación",   "Alternativa a la devolución: imputación a impuestos fiscales."),
    ("6. Crear empresa (sociedad)",   "Regímenes 14A / 14D N°3 / 14D N°8, efectos en renta e IVA."),
    ("7. Efectos en IVA",             "Cuándo se aplica IVA, cuándo no, y cómo optimizar."),
])

# ─────────────────────── 1. COMPRA DEL INMUEBLE ────────────────────────────────
story += [PageBreak()]
story += seccion_titulo("1.  COMPRA DEL BIEN INMUEBLE — EFECTOS EN IVA")

story.append(Paragraph(
    "Desde el <b>1° de marzo de 2020</b> (Ley N° 21.210 – Modernización Tributaria), la compraventa de "
    "bienes corporales inmuebles <i>construidos</i> queda afecta a IVA siempre que el vendedor sea "
    "<b>habitual</b>. Los terrenos (sin construcción) no están gravados.", cuerpo))

story += [Paragraph("1.1 Concepto de habitualidad", titulo_sub)]
story.append(Paragraph(
    "El SII califica la habitualidad a su <b>juicio exclusivo</b>, evaluando las circunstancias "
    "particulares de cada caso (frecuencia de operaciones, volumen, giro registrado, etc.). "
    "Se eliminó la presunción automática de habitualidad cuando la venta ocurría antes de un año "
    "desde la adquisición o término de construcción.", cuerpo))

story += cuadro_comparativo("¿Cuándo se aplica IVA en la compra?",
    ["Situación", "¿IVA?", "Referencia legal"],
    [
        ["Compra a empresa constructora o inmobiliaria (habitual)", "✔ Sí – 19%", "Art. 8 + Art. 2 N°3 DL 825"],
        ["Compra a persona natural que no es habitual (ej.: vende su casa)", "✖ No", "Art. 2 N°3 DL 825"],
        ["Compra a empresa que vende inmueble que no formó parte de su activo fijo habitual", "Depende del SII", "Circ. 37/2020"],
        ["Compra de terreno sin construcción (sitio eriazo)", "✖ No", "Art. 2 N°1 DL 825"],
        ["Compra de sitio urbanizado (con obras al menos en parte)", "✔ Sí", "Circ. 37/2020"],
    ]
)

story += [Paragraph("1.2 Base imponible del IVA en la compra", titulo_sub)]
story.append(Paragraph(
    "En la venta de inmuebles nuevos se <b>descuenta obligatoriamente</b> el valor de adquisición "
    "del terreno (reajustado por IPC, sin tope desde el 01.03.2020). En inmuebles donde el vendedor "
    "no soportó IVA al adquirirlos, la base imponible es la <b>diferencia entre precio de venta y "
    "precio de compra</b>, descontando en ambos el valor comercial del terreno.", cuerpo))

story += [Paragraph("1.3 Crédito fiscal del comprador", titulo_sub)]
story.append(Paragraph(
    "Si el comprador es contribuyente de IVA (por ejemplo, una empresa o persona natural que arrienda "
    "inmuebles amoblados), el IVA soportado en la compra constituye <b>crédito fiscal</b> utilizable "
    "contra el débito fiscal de sus operaciones.", cuerpo))
story.append(Paragraph(
    "Si el comprador es <b>persona natural sin contabilidad</b> y arrienda inmuebles sin amoblar, "
    "<b>NO</b> puede recuperar el IVA pagado en la compra, ya que no es contribuyente de IVA.", cuerpo))

story += [Paragraph("1.4 CEEC — Crédito Especial Empresa Constructora (referencia)", titulo_sub)]
story.append(Paragraph(
    "El CEEC (DL 910/1975) permite a las constructoras deducir el 65% del IVA débito generado en ventas "
    "de viviendas ≤ 2.000 UF (tope 225 UF por vivienda). Este crédito está siendo eliminado "
    "gradualmente: íntegro hasta obras con permiso anterior al 30.04.2023; parcial (50%) hasta "
    "01.01.2025; parcial (25%) hasta 01.01.2027; <b>eliminación total desde 01.01.2027</b>.", cuerpo))

# ─────────────────────── 2. ARRENDAMIENTO ────────────────────────────────────
story += [PageBreak()]
story += seccion_titulo("2.  ARRENDAMIENTO DEL INMUEBLE — TRIBUTACIÓN EN RENTA E IVA")

story.append(Paragraph(
    "Los ingresos por arrendamiento de bienes inmuebles tienen tratamientos distintos según si la "
    "propiedad está <b>amoblada o no</b>, si el propietario es <b>persona natural o jurídica</b>, "
    "y si el inmueble califica como <b>DFL2</b>.", cuerpo))

story += cuadro_comparativo("Tributación del arrendamiento según tipo de propiedad y propietario",
    ["Situación", "Impuesto a la Renta", "IVA"],
    [
        ["PN – Inmueble DFL2 (máx. 2 prop.) – No amoblado",
         "Ingreso No Renta (INR) – exento", "No afecto"],
        ["PN – Inmueble no DFL2 – No amoblado",
         "Renta 1ª Categoría Art. 20 N°1 LIR (IDPC 25–27%)", "No afecto"],
        ["PN – Inmueble amoblado",
         "Renta 1ª Categoría Art. 20 N°1 (IDPC)", "Afecto: 19% s/ base reducida"],
        ["Empresa – Inmueble no amoblado",
         "Renta 1ª Categoría (IDPC según régimen)", "No afecto"],
        ["Empresa – Inmueble amoblado",
         "Renta 1ª Categoría (IDPC)", "Afecto: 19% s/ base reducida"],
    ]
)

story += [Paragraph("2.1 ¿Cuándo un inmueble se considera 'amoblado'?", titulo_sub)]
story.append(Paragraph(
    "El SII exige que los muebles sean <b>suficientes para el uso habitacional o comercial</b> "
    "al que está destinado el inmueble. Ejemplos:", cuerpo))
for item in [
    "Departamento 2D: cama (o similar), comedor y vajilla que permitan pernoctar.",
    "Oficina: escritorio, sillas y repisas suficientes para actividades administrativas.",
    "Local comercial: mesón de atención al público y estantes para mostrar productos.",
]:
    story.append(Paragraph(f"• {item}", viñeta))

story += [Paragraph("2.2 Base imponible del IVA en arriendo de inmueble amoblado", titulo_sub)]
story.append(Paragraph(
    "Desde la Ley N° 21.210 es obligatorio deducir de la renta mensual una cantidad equivalente al "
    "<b>11% del avalúo fiscal anual del inmueble</b> (proporcionalmente si el arriendo es parcial "
    "o por período distinto a un año). Solo el remanente queda afecto a IVA 19%.", cuerpo))
story.append(Paragraph(
    "Ejemplo: Arriendo mensual $800.000 · Avalúo fiscal $40.000.000 · "
    "Deducción mensual: 11% × $40M / 12 = $366.667 · "
    "<b>Base IVA: $433.333 · IVA: $82.333</b>", nota_legal))

story += [Paragraph("2.3 Renta de Primera Categoría — Art. 20 N°1 LIR", titulo_sub)]
story.append(Paragraph(
    "Los arrendamientos de inmuebles que no califican como INR (fuera de DFL2 o persona jurídica) "
    "tributan como rentas de primera categoría bajo el Art. 20 N°1. Esto significa que el "
    "propietario queda obligado a llevar contabilidad (al menos simplificada), declarar renta y pagar "
    "IDPC según el régimen al que pertenezca (25% o 27%) y finalmente IGC con los créditos "
    "correspondientes.", cuerpo))

# ─────────────────────── 3. INGRESOS NO RENTA ────────────────────────────────
story += [PageBreak()]
story += seccion_titulo("3.  INGRESOS NO CONSTITUTIVOS DE RENTA (INR)")

story.append(Paragraph(
    "Los INR son ingresos que <b>por expresa disposición legal</b> no forman parte del hecho gravado "
    "de la LIR. No se declaran ni se consideran en la progresión de los impuestos finales (IUSC/IGC).", cuerpo))

story += [Paragraph("3.1 INR por arrendamiento — DFL 2 (D.F.L. N°2 de 1959)", titulo_sub)]
story.append(Paragraph(
    "Las rentas de arrendamiento de propiedades DFL2 constituyen <b>INR para personas naturales</b>, "
    "siempre que no estén amobladas:", cuerpo))
for item in [
    "Desde la <b>Ley N° 21.420 de 2022</b>: el beneficio DFL2 aplica a un máximo de <b>2 propiedades</b> "
    "por persona natural, independientemente de cuándo fueron adquiridas.",
    "Personas jurídicas (sociedades): el beneficio DFL2 está <b>eliminado</b>. Sus rentas de "
    "arriendo siempre tributan como renta de primera categoría.",
    "Si el inmueble está amoblado, se pierde el INR y aplica IVA.",
]:
    story.append(Paragraph(f"• {item}", viñeta))

story += [Paragraph("3.2 INR por mayor valor en venta de bienes raíces — Art. 17 N°8 LIR", titulo_sub)]
story.append(Paragraph(
    "Las personas naturales que no llevan contabilidad completa pueden obtener un beneficio de "
    "<b>INR hasta 8.000 UF</b> al vender bienes raíces. La estructura es:", cuerpo))

story += cuadro_comparativo("Mayor valor en enajenación de bienes raíces (PN sin cont. completa)",
    ["Condición", "Tributación", "Observación"],
    [
        ["Bien adquirido desde 01.01.2004 · Venta a no relacionado · Transcurrido plazo (1 año no agríc. / 4 años agríc.)",
         "INR hasta 8.000 UF. Exceso: IUS 10% o IGC/IA a elección",
         "8.000 UF acumulables a lo largo de la vida del contribuyente"],
        ["Bien vendido ANTES del plazo de 1 o 4 años",
         "Impuesto Único Sustitutivo 10%",
         "No aplica INR de 8.000 UF"],
        ["Venta a relacionado",
         "IGC o IA (percibido o devengado)",
         "Sin INR ni IUS. Sin reliquidación."],
        ["Contribuyente que declara IDPC s/contabilidad completa",
         "Renta Art. 20 N°5 (régimen general)",
         "No aplica Art. 17 N°8 b)"],
        ["Inmueble adquirido hasta 31.12.2003",
         "Régimen vigente al 31.12.2014 (INR si cumple requisitos)",
         "No aplica recuadro N°2 del F22"],
    ]
)
story.append(Paragraph(
    "⚠ Importante: Si el inmueble está en el activo de una empresa que lleva contabilidad completa, "
    "el mayor valor tributa como renta del Art. 20 N°5 (IDPC + IGC), sin beneficio de INR.", alerta))

# ─────────────────────── 4. ART. 27 BIS — DEVOLUCIÓN ────────────────────────
story += [PageBreak()]
story += seccion_titulo("4.  ARTÍCULO 27 BIS — DEVOLUCIÓN DEL REMANENTE DE CRÉDITO FISCAL IVA")

story.append(Paragraph(
    "El <b>artículo 27 bis del DL 825</b> permite a los contribuyentes de IVA recuperar el remanente "
    "de crédito fiscal acumulado cuando éste proviene de adquisiciones destinadas al <b>activo fijo</b>. "
    "La Ley N° 21.210 simplificó y aceleró este mecanismo.", cuerpo))

story += [Paragraph("4.1 Requisitos para solicitar la devolución", titulo_sub)]
for item in [
    "<b>Ser contribuyente de IVA</b> (empresa, persona natural habitual, exportador).",
    "Tener remanente de crédito fiscal acumulado durante <b>2 o más períodos tributarios consecutivos</b> "
    "(antes de la Ley 21.210 eran 6 períodos).",
    "El remanente debe originarse en la adquisición de bienes o servicios destinados al <b>activo fijo</b> "
    "(bienes muebles o inmuebles).",
    "Para inmuebles: se consideran activo fijo <b>desde que la obra o etapa es recibida conforme</b> "
    "por quien la encargó. Si se obtienen devoluciones durante la construcción, deberá acreditarse "
    "la recepción definitiva y la incorporación efectiva al activo inmovilizado.",
    "No es procedente solicitar en el mismo período devolución por 27 bis <b>y</b> IVA exportador "
    "(Art. 36) o cambio de sujeto.",
]:
    story.append(Paragraph(f"• {item}", viñeta))

story += [Paragraph("4.2 Procedimiento de solicitud", titulo_sub)]
story += cuadro_resumen("Pasos para solicitar la devolución Art. 27 bis", [
    ("Paso 1 – F3280",
     "Presentar solicitud electrónica 'Solicitud de Devolución artículo 27 bis D.L. 825/74' (Formulario N°3280) "
     "en www.sii.cl, adjuntando expediente electrónico con documentación fundante. "
     "El F29 del último período debe estar previamente declarado y pagado."),
    ("Paso 2 – Plazo SII",
     "El SII tiene <b>20 días hábiles</b> para resolver (reducido desde 60 días por Ley 21.210). "
     "El plazo se cuenta desde que el contribuyente pone a disposición todos los antecedentes."),
    ("Paso 3 – Silencio positivo",
     "Si el SII no se pronuncia dentro del plazo, opera el silencio administrativo positivo: "
     "la solicitud se entiende aprobada y Tesorería General de la República debe efectuar la devolución."),
    ("Paso 4 – Devolución",
     "Tesorería paga mediante cheque o transferencia bancaria. El SII puede hacer revisiones "
     "posteriores conforme al Código Tributario."),
    ("Fiscalización Especial Previa (FEP)",
     "Para activo fijo de mayor valor o complejidad (ej.: construcción de inmuebles), el SII puede "
     "decretar FEP con plazo de 25 días adicionales para revisar los antecedentes."),
])

story += [Paragraph("4.3 Restitución del crédito devuelto", titulo_sub)]
story.append(Paragraph(
    "El contribuyente debe restituir el crédito recibido de la siguiente manera:", cuerpo))
for item in [
    "<b>Restitución normal</b>: mediante los pagos efectivos de IVA que realice en los meses siguientes "
    "por sus operaciones gravadas.",
    "<b>Restitución adicional</b>: si el contribuyente realiza operaciones exentas o no gravadas, "
    "deberá agregar al débito fiscal una cantidad proporcional calculada en base a la proporción "
    "de dichas operaciones sobre el total.",
    "<b>Término de giro / Fusión</b>: el remanente no restituido debe ser devuelto al Fisco, "
    "adicionándolo al débito fiscal del mes del término (Oficio N°2619 de 2021).",
]:
    story.append(Paragraph(f"• {item}", viñeta))

# ─────────────────────── 5. ART. 27 BIS — IMPUTACIÓN ────────────────────────
story += seccion_titulo("5.  ARTÍCULO 27 BIS — IMPUTACIÓN DEL REMANENTE")

story.append(Paragraph(
    "Alternativamente a la devolución en efectivo, el contribuyente puede optar por <b>imputar</b> "
    "el remanente de crédito fiscal directamente a otros impuestos. Esta es una herramienta de "
    "planificación tributaria muy relevante.", cuerpo))

story += [Paragraph("5.1 ¿A qué impuestos se puede imputar?", titulo_sub)]
story.append(Paragraph(
    "El inciso primero del Art. 27 bis dispone que el remanente puede imputarse a "
    "<b>cualquier clase de impuestos fiscales</b>, incluyendo:", cuerpo))
for item in [
    "Pagos Provisionales Mensuales (PPM) del Impuesto a la Renta.",
    "Impuesto de Primera Categoría anual.",
    "Impuesto Global Complementario o Impuesto Adicional.",
    "Retenciones y otros impuestos mensuales de declaración en el F29.",
    "Cualquier otro tributo fiscal vigente.",
]:
    story.append(Paragraph(f"• {item}", viñeta))

story.append(Paragraph(
    "✔ <b>Ventaja estratégica</b>: Si el contribuyente tiene importantes obligaciones en otros "
    "impuestos (ej.: IGC elevado), la imputación permite reducirlos directamente sin esperar "
    "la devolución en efectivo, optimizando el flujo de caja.", positivo))

story += [Paragraph("5.2 Elección entre devolución e imputación", titulo_sub)]
story.append(Paragraph(
    "El contribuyente elige en cada período cuál mecanismo utiliza (devolución o imputación). "
    "No es posible operar simultáneamente con IVA exportador (Art. 36) y Art. 27 bis en el mismo "
    "período, ya que la base para cada devolución es una sola. Puede solicitar uno en un período "
    "y el otro en períodos siguientes.", cuerpo))
story.append(Paragraph(
    "El porcentaje de devolución/imputación se calcula aplicando la proporción que representa el "
    "crédito fiscal de activo fijo sobre el total del crédito fiscal acumulado.", nota_legal))

# ─────────────────────── 6. CREAR EMPRESA ────────────────────────────────────
story += [PageBreak()]
story += seccion_titulo("6.  ¿QUÉ PASA SI SE CREA UNA EMPRESA?")

story.append(Paragraph(
    "Constituir una sociedad para llevar el negocio inmobiliario tiene importantes consecuencias "
    "tributarias, tanto positivas como negativas. A continuación se analizan los principales efectos.", cuerpo))

story += [Paragraph("6.1 Regímenes tributarios disponibles", titulo_sub)]
story += cuadro_comparativo("Comparativa de regímenes (Ley N°21.210, vigente desde 01.01.2020)",
    ["Régimen", "IDPC", "Crédito IDPC en IGC", "Contabilidad", "Aplica si"],
    [
        ["14A – Semi Integrado (General)", "27%",
         "Parcial (65% del IDPC como crédito)",
         "Completa obligatoria",
         "Cualquier empresa sin límite de ingresos"],
        ["14D N°3 – Pro Pyme General", "25%",
         "Parcial (mismo mecanismo 14A)",
         "Simplificada o completa",
         "Ventas ≤ 75.000 UF, capital ≤ 85.000 UF"],
        ["14D N°8 – Pro Pyme Transparente", "0% (empresa)",
         "N/A (socios pagan IGC/IA directamente)",
         "Libro ingresos/egresos + Caja",
         "Pyme, socios PN o pequeñas empresas"],
    ]
)

story += [Paragraph("6.2 Efectos positivos de tener empresa", titulo_sub)]
for item in [
    "<b>Crédito fiscal IVA recuperable</b>: La empresa como contribuyente de IVA puede usar el "
    "crédito fiscal de la compra del inmueble (si el vendedor era habitual) contra el débito "
    "fiscal de sus arrendamientos amoblados u otras operaciones.",
    "<b>Art. 27 bis disponible</b>: Si el inmueble es activo fijo, puede solicitar la devolución "
    "o imputación del remanente de CF IVA desde solo 2 períodos de acumulación.",
    "<b>Deducciones de gastos</b>: La empresa puede deducir como gastos necesarios: "
    "intereses hipotecarios, contribuciones de bienes raíces, mantención, administración, "
    "depreciación del inmueble construido, seguros, etc.",
    "<b>Mayor liquidez por deducción contribuciones</b>: Las contribuciones pagadas son crédito "
    "contra el IDPC (Art. 20 N°1 LIR).",
    "<b>Planificación del retiro</b>: En el régimen 14D N°8 (Transparente), los socios personas "
    "naturales tributan directamente con IGC según sus tasas marginales, sin capa de IDPC.",
]:
    story.append(Paragraph(f"✔ {item}", positivo))

story += [Paragraph("6.3 Efectos negativos / riesgos de tener empresa", titulo_sub)]
for item in [
    "<b>Pérdida del beneficio DFL2</b>: Las personas jurídicas no tienen derecho al INR por "
    "arrendamiento DFL2. Sus ingresos por arriendo siempre tributan como renta de primera categoría.",
    "<b>Pérdida del INR de 8.000 UF en venta de inmuebles</b>: Si el inmueble está en el activo "
    "de la empresa con contabilidad completa, el mayor valor al venderlo NO goza del INR de "
    "8.000 UF. Tributa como renta del Art. 20 N°5 (IDPC + IGC).",
    "<b>IVA en arrendamiento amoblado</b>: La empresa debe cobrar IVA en arriendos de inmuebles "
    "amoblados (19% sobre base reducida), lo que puede encarecer el arriendo para inquilinos "
    "no contribuyentes de IVA.",
    "<b>Doble tributación en régimen 14A</b>: IDPC 27% a nivel empresa + IGC del socio (parcialmente "
    "integrado). En 14D N°8 (Transparente) se evita la doble capa.",
    "<b>Costos de administración</b>: Contabilidad, declaraciones mensuales, registros empresariales "
    "RAI/REX/SAC requieren un contador.",
]:
    story.append(Paragraph(f"⚠ {item}", alerta))

story += [Paragraph("6.4 ¿Qué tipo de sociedad conviene?", titulo_sub)]
story.append(Paragraph(
    "Para un negocio inmobiliario de mediana escala, las opciones más habituales son:", cuerpo))
for item in [
    "<b>SpA (Sociedad por Acciones)</b>: Flexible, un solo socio posible, fácil de constituir. "
    "Puede acogerse a 14D N°3 o 14A según nivel de ingresos.",
    "<b>Sociedad de Personas (SP) o EIRL</b>: Puede acogerse a 14D N°3 o 14D N°8 (Transparente) "
    "si los socios son personas naturales, evitando la capa del IDPC.",
    "<b>S.A. cerrada</b>: Mayor formalismo. Se acoge automáticamente a 14A (27%). "
    "No puede ser 14D N°8.",
]:
    story.append(Paragraph(f"• {item}", viñeta))

# ─────────────────────── 7. EFECTOS IVA ─────────────────────────────────────
story += [PageBreak()]
story += seccion_titulo("7.  EFECTOS GLOBALES EN IVA — MAPA DE SITUACIONES")

story.append(Paragraph(
    "A continuación se resume el mapa completo de situaciones de IVA según la decisión del "
    "cliente de actuar como persona natural o mediante una empresa:", cuerpo))

story += cuadro_comparativo("IVA según situación del propietario y tipo de arrendamiento",
    ["Situación", "Compra inmueble", "Arriendo no amoblado", "Arriendo amoblado"],
    [
        ["PN – sin empresa",
         "Paga IVA si compra a habitual. No recupera CF.",
         "No IVA. INR si DFL2 (≤2 prop.), o 1ª Cat. si no DFL2.",
         "IVA 19% (base: renta − 11% av. fiscal). 1ª Cat. IDPC."],
        ["Empresa – contribuyente IVA",
         "Paga IVA si compra a habitual. CF recuperable o 27 bis.",
         "No IVA. Renta 1ª Categoría IDPC (25% o 27%).",
         "IVA 19% (base: renta − 11% av. fiscal). IDPC."],
        ["Empresa – vende inmueble (habitualmente)",
         "CF por compra si vendedor fue habitual.",
         "N/A",
         "IVA en venta del inmueble si es habitual."],
    ]
)

story += [Paragraph("7.1 ¿Cuándo conviene ser contribuyente de IVA?", titulo_sub)]
story.append(Paragraph(
    "Ser contribuyente de IVA (generalmente mediante empresa) conviene cuando:", cuerpo))
for item in [
    "Se adquieren <b>inmuebles nuevos a constructoras</b>: el IVA pagado en la compra "
    "(crédito fiscal alto) puede recuperarse vía Art. 27 bis o imputarse.",
    "Se proyectan <b>arrendamientos de inmuebles amoblados</b>: se genera débito fiscal con el "
    "cual compensar el crédito acumulado.",
    "El cliente tiene <b>otras actividades gravadas con IVA</b> que generan débito fiscal "
    "suficiente para absorber el crédito.",
    "Se planifica la <b>venta futura del inmueble</b> como parte del giro habitual "
    "(aunque en ese caso se pierde el INR de 8.000 UF en la venta).",
]:
    story.append(Paragraph(f"• {item}", viñeta))

story += [Paragraph("7.2 ¿Cuándo NO conviene ser contribuyente de IVA?", titulo_sub)]
story.append(Paragraph(
    "Mantenerse como persona natural sin empresa puede ser mejor cuando:", cuerpo))
for item in [
    "El inmueble califica como <b>DFL2 y se arrienda sin amoblar</b>: los ingresos son INR "
    "(exentos de renta) y no hay IVA.",
    "Se proyecta <b>vender el inmueble en el futuro</b> a largo plazo y se quiere mantener "
    "el beneficio de INR hasta 8.000 UF en la ganancia de capital.",
    "Los inquilinos son <b>personas naturales que no recuperan IVA</b>, por lo que el IVA "
    "en el arriendo amoblado los encarece.",
    "El inmueble se compra a <b>no habitual</b> (sin IVA), por lo que no hay crédito fiscal "
    "que recuperar.",
]:
    story.append(Paragraph(f"• {item}", viñeta))

# ─────────────────────── 8. CUADRO RESUMEN FINAL ────────────────────────────
story += [PageBreak()]
story += seccion_titulo("8.  CUADRO RESUMEN: ¿PERSONA NATURAL vs. EMPRESA?")

story += cuadro_comparativo("",
    ["Aspecto", "Persona Natural (sin empresa)", "Con Empresa Inmobiliaria"],
    [
        ["IVA en compra", "Paga IVA si vendedor habitual. No recupera CF.",
         "Paga IVA si vendedor habitual. CF recuperable."],
        ["Art. 27 bis", "No aplica (no es contribuyente IVA si no hay empresa).",
         "Sí aplica: devolución o imputación desde 2 períodos."],
        ["Arriendo no amoblado – DFL2 (≤2 prop.)", "INR — sin impuesto.", "No aplica INR. Renta 1ª Cat. IDPC."],
        ["Arriendo no amoblado – no DFL2", "Renta 1ª Cat. IDPC (sin IVA).", "Renta 1ª Cat. IDPC (sin IVA)."],
        ["Arriendo amoblado", "IVA + Renta 1ª Cat. IDPC.", "IVA + Renta 1ª Cat. IDPC."],
        ["Mayor valor venta inmueble", "INR hasta 8.000 UF (si cumple plazos y requisitos).",
         "Tributa régimen general IDPC + IGC. Sin INR."],
        ["Tasa IDPC", "No aplica (IGC directo si no lleva contabilidad).",
         "25% (14D N°3) o 27% (14A). 0% en 14D N°8 (transparente)."],
        ["Gastos deducibles", "Limitados.", "Amplia deducción: intereses, mantención, depreciación, contribuciones."],
        ["Complejidad administrativa", "Baja.", "Media–Alta (contabilidad, DJA, registros RAI/REX)."],
    ]
)

# ─────────────────────── CONCLUSIONES ────────────────────────────────────────
story += seccion_titulo("CONCLUSIONES Y RECOMENDACIONES")

story.append(Paragraph(
    "Sobre la base del análisis realizado, se formulan las siguientes recomendaciones preliminares "
    "para el cliente:", cuerpo))

recomendaciones = [
    ("1", "Evaluar si el vendedor del inmueble es habitual",
     "Si compra a una constructora o inmobiliaria, habrá IVA. Si el comprador será una empresa, "
     "ese IVA se convierte en crédito fiscal recuperable. Si es PN, ese IVA es un costo."),
    ("2", "Considerar el DFL2 si se quiere simplicidad",
     "Si el objetivo es arrendar sin amoblar y obtener INR, mantener hasta 2 propiedades DFL2 "
     "como persona natural puede ser la opción más eficiente. El ingreso está exento de impuesto."),
    ("3", "Usar Art. 27 bis si se constituye empresa",
     "La empresa puede recuperar el remanente de CF IVA por activo fijo en solo 2 meses, "
     "lo cual mejora significativamente el flujo de caja en proyectos de inversión inmobiliaria."),
    ("4", "Cuidar el INR de 8.000 UF en la venta",
     "Si se proyecta vender el inmueble en el futuro con ganancia, es fundamental no incorporarlo "
     "al activo de una empresa con contabilidad completa. Como persona natural (sin IDPC s/CC) "
     "el mayor valor hasta 8.000 UF es INR."),
    ("5", "Regime 14D N°8 si se constituye empresa con socios PN",
     "Si se opta por empresa y los socios son personas naturales, el régimen de Transparencia "
     "Tributaria (14D N°8) elimina la capa del IDPC, gravando directamente con IGC. Esto puede "
     "ser ventajoso para socios con tasas marginales bajas."),
]

for num, titulo_r, texto_r in recomendaciones:
    data = [
        [Paragraph(num, ParagraphStyle("N", fontSize=12, textColor=colors.white,
                                        fontName="Helvetica-Bold", alignment=TA_CENTER)),
         Paragraph(titulo_r, ParagraphStyle("T", fontSize=10, textColor=AZUL,
                                              fontName="Helvetica-Bold", leading=14)),
         Paragraph(texto_r, ParagraphStyle("B", fontSize=9, textColor=NEGRO, leading=13))]
    ]
    t = Table(data, colWidths=[0.8*cm, 5.5*cm, 11.5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (0,-1), CELESTE),
        ("BACKGROUND",    (1,0), (-1,-1), GRIS),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("TOPPADDING",    (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ("RIGHTPADDING",  (0,0), (-1,-1), 8),
        ("LINEBELOW",     (0,0), (-1,-1), 0.5, colors.HexColor("#DDDDDD")),
    ]))
    story += [t, Spacer(1, 0.15*cm)]

story.append(Spacer(1, 0.5*cm))

# ─────────────────────── DISCLAIMER ─────────────────────────────────────────
story.append(hr(color=colors.HexColor("#AAAAAA"), width=0.5))
story.append(Paragraph(
    "NOTA LEGAL: El presente informe ha sido elaborado con fines informativos y de apoyo a la asesoría "
    "tributaria, basado en el material académico del Magíster en Dirección Tributaria de la Universidad "
    "Viña del Mar (UVM) y la legislación tributaria vigente referenciada. No constituye una opinión "
    "legal formal ni reemplaza el análisis particularizado del caso de cada contribuyente. "
    "Se recomienda verificar la normativa vigente al momento de su aplicación práctica. "
    "Referencias: DL 825 (IVA), DL 824 (LIR), Ley N°21.210 de 2020, Ley N°21.420 de 2022, "
    "Circular N°37 de 2020, Circular N°94 de 2001, Circular N°73 de 2020.", disclaimer))

# ── Generar PDF ────────────────────────────────────────────────────────────────
doc.build(story)
print(f"PDF generado: {OUTPUT}")
