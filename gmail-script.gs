// ══════════════════════════════════════
//  WalletWolf – Gmail Import Script v2
//  script.google.com → Nuevo proyecto → pega este código
// ══════════════════════════════════════
//
//  SETUP v2 (hacer una sola vez):
//  1. Abre script.google.com, crea un proyecto nuevo.
//  2. Crea una Google Sheet en drive.google.com y copia su ID
//     (la parte larga entre /d/ y /edit en la URL).
//  3. Pega ese ID en SHEET_ID abajo.
//  4. Genera un token aleatorio (ej. abre DevTools → console →
//     crypto.randomUUID()) y pégalo en TOKEN abajo.
//  5. Pega todo este código en el editor.
//  6. Ejecuta setupTrigger() (menú Ejecutar → setupTrigger)
//     → acepta los permisos de Gmail, Sheets y ScriptApp.
//  7. Menú Implementar → Nueva implementación.
//     Tipo: Aplicación web / Ejecutar como: Yo /
//     Quién puede acceder: Cualquiera → Implementar.
//  8. Copia la URL que aparece y agrégale el token al final:
//        https://script.google.com/macros/s/XXXX/exec?token=TU_TOKEN
//  9. En WalletWolf → Ajustes → Banco → pega esa URL completa.
//
//  HOJAS que crea el script automáticamente:
//    Pendientes  – transacciones pendientes de confirmar por la app
//    Fallidos    – emails que no pudieron parsearse (para depurar)
//    Archivo     – filas procesadas con más de 60 días
// ══════════════════════════════════════

const TOKEN      = 'GENERA_UN_TOKEN_ALEATORIO_AQUI';
const SHEET_ID   = 'PEGA_AQUI_EL_ID_DE_TU_GOOGLE_SHEET';
const SHEET_NAME = 'Pendientes';

// ── Parsers por banco ────────────────────────────────────────────────
// Estructura: { banco, fromDomains[], queries[], parse(subject, body, emailFecha, emailDate, id) }
// parse() retorna un objeto tx o null.
// Si retorna null y el email matcheó el from → se registra en Fallidos.
// Los regex de Banco de Chile están validados en producción.
// Los de BCI y BancoEstado son provisionales — marcar con TODO al iterar.

const PARSERS = [
  {
    banco: 'Banco de Chile',
    fromDomains: ['enviodigital.bancochile.cl', 'bancochile.cl'],
    queries: [
      'subject:"Cargo en Cuenta"',
      'subject:"Compra con Tarjeta de Crédito"',
      'subject:"Compra con Tarjeta de Débito"',
      'subject:"Comprobante de Transferencia a terceros"',
      'subject:"Transferencia realizada"',
      'subject:"Has recibido una transferencia de fondos"',
      'subject:"Transferencia recibida"',
      'subject:"Abono en Cuenta"',
    ],
    parse: function(subject, body, emailFecha, emailDate, id) {
      if (/cargo en cuenta/i.test(subject)) {
        const m = body.match(/por\s*\$\s*([\d.,]+)/i);
        const d = body.match(/en\s+([A-ZÁÉÍÓÚÜÑ][A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,&'-]+?)\s+el\s+\d{2}\/\d{2}/i);
        if (m) return { id: id, fecha: fechaFromBodySafe(body, emailFecha, emailDate), desc: d ? clean(d[1]) : 'Cargo en cuenta', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/compra con tarjeta/i.test(subject)) {
        const m = body.match(/por\s*\$\s*([\d.,]+)/i);
        const d = body.match(/en\s+([A-ZÁÉÍÓÚÜÑ][A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,&'-]+?)\s+el\s+\d{2}\/\d{2}/i);
        if (m) return { id: id, fecha: fechaFromBodySafe(body, emailFecha, emailDate), desc: d ? clean(d[1]) : 'Compra con tarjeta', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/comprobante de transferencia|transferencia a terceros|transferencia realizada/i.test(subject)) {
        const m = body.match(/Monto[\s:$]*([0-9.,]+)/i);
        const d = body.match(/(?:Nombre|Destinatario|Beneficiario)[:\s]+([A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,]+)/i);
        if (m) return { id: id, fecha: emailFecha, desc: d ? 'TRF → ' + clean(d[1]) : 'Transferencia saliente', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/recibido una transferencia|transferencia recibida/i.test(subject)) {
        const m = body.match(/Monto transferido[:\s$]*([0-9.,]+)/i)
               || body.match(/Monto[:\s$]*([0-9.,]+)/i);
        const d = body.match(/transferencia de fondos de\s+([A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,]+?)[\r\n]/i)
               || body.match(/de\s+([A-Z][A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,]+?)\s+(?:Monto|RUT|por)/i);
        if (m) return { id: id, fecha: emailFecha, desc: d ? 'TRF ← ' + clean(d[1]) : 'Transferencia entrante', monto: parseMonto(m[1]), tipo: 'ingreso' };
      }
      if (/abono en cuenta/i.test(subject)) {
        const m = body.match(/\$\s*([\d.,]+)/i);
        if (m) return { id: id, fecha: emailFecha, desc: 'Abono en cuenta', monto: parseMonto(m[1]), tipo: 'ingreso' };
      }
      return null;
    }
  },

  {
    banco: 'BCI',
    fromDomains: ['bci.cl'],
    queries: [
      // TODO: validar asuntos exactos con emails reales de BCI
      'subject:"Notificación de compra"',
      'subject:"Cargo realizado"',
      'subject:"Transferencia enviada"',
      'subject:"Transferencia recibida"',
    ],
    parse: function(subject, body, emailFecha, emailDate, id) {
      // TODO: validar con email real — regex provisionales
      if (/notificaci.n de compra|cargo realizado/i.test(subject)) {
        const m = body.match(/\$\s*([\d.,]+)/i) || body.match(/monto[:\s]*\$?\s*([\d.,]+)/i);
        const d = body.match(/(?:comercio|establecimiento|lugar)[:\s]+([^\r\n]{2,50})/i);
        if (m) return { id: id, fecha: fechaFromBodySafe(body, emailFecha, emailDate), desc: d ? clean(d[1]) : 'Compra BCI', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/transferencia enviada/i.test(subject)) {
        const m = body.match(/monto[:\s]*\$?\s*([\d.,]+)/i);
        const d = body.match(/(?:destinatario|nombre)[:\s]+([^\r\n]{2,50})/i);
        if (m) return { id: id, fecha: emailFecha, desc: d ? 'TRF → ' + clean(d[1]) : 'Transferencia BCI', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/transferencia recibida/i.test(subject)) {
        const m = body.match(/monto[:\s]*\$?\s*([\d.,]+)/i);
        const d = body.match(/(?:origen|remitente|nombre)[:\s]+([^\r\n]{2,50})/i);
        if (m) return { id: id, fecha: emailFecha, desc: d ? 'TRF ← ' + clean(d[1]) : 'Transferencia recibida BCI', monto: parseMonto(m[1]), tipo: 'ingreso' };
      }
      return null;
    }
  },

  {
    banco: 'BancoEstado',
    fromDomains: ['bancoestado.cl'],
    queries: [
      // TODO: validar asuntos exactos con emails reales de BancoEstado
      'subject:"Cargo en tu cuenta"',
      'subject:"Compra realizada"',
      'subject:"Transferencia efectuada"',
      'subject:"Transferencia recibida"',
    ],
    parse: function(subject, body, emailFecha, emailDate, id) {
      // TODO: validar con email real — regex provisionales
      if (/cargo en tu cuenta|compra realizada/i.test(subject)) {
        const m = body.match(/\$\s*([\d.,]+)/i) || body.match(/monto[:\s]*\$?\s*([\d.,]+)/i);
        const d = body.match(/(?:comercio|establecimiento)[:\s]+([^\r\n]{2,50})/i);
        if (m) return { id: id, fecha: fechaFromBodySafe(body, emailFecha, emailDate), desc: d ? clean(d[1]) : 'Compra BancoEstado', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/transferencia efectuada/i.test(subject)) {
        const m = body.match(/monto[:\s]*\$?\s*([\d.,]+)/i);
        const d = body.match(/(?:beneficiario|destinatario)[:\s]+([^\r\n]{2,50})/i);
        if (m) return { id: id, fecha: emailFecha, desc: d ? 'TRF → ' + clean(d[1]) : 'Transferencia BancoEstado', monto: parseMonto(m[1]), tipo: 'gasto' };
      }
      if (/transferencia recibida/i.test(subject)) {
        const m = body.match(/monto[:\s]*\$?\s*([\d.,]+)/i);
        const d = body.match(/(?:origen|remitente|de)[:\s]+([^\r\n]{2,50})/i);
        if (m) return { id: id, fecha: emailFecha, desc: d ? 'TRF ← ' + clean(d[1]) : 'Transferencia recibida BancoEstado', monto: parseMonto(m[1]), tipo: 'ingreso' };
      }
      return null;
    }
  }
];

// ── Web App: GET ──────────────────────────────────────────────────────
function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN) {
    return jsonResp({ error: 'unauthorized' });
  }

  const action = e.parameter.action || null;

  if (action === 'confirm') {
    const sheet = getSheet();
    const ids = (e.parameter.ids || e.parameter.id || '').split(',').filter(Boolean);
    for (const id of ids) markProcessed(sheet, id.trim());
    return jsonResp({ ok: true });
  }

  if (action === 'status') {
    const sheet = getSheet();
    const fallSheet = getFallidosSheet();
    const rows = sheet.getDataRange().getValues().slice(1);
    const cutoff30 = Date.now() - 30 * 24 * 3600 * 1000;
    const pendientes = rows.filter(function(r) { return !r[5]; }).length;
    const procesados30d = rows.filter(function(r) {
      if (!r[5]) return false;
      const d = r[1] instanceof Date ? r[1].getTime() : new Date(String(r[1])).getTime();
      return d >= cutoff30;
    }).length;
    const fallidos = Math.max(0, fallSheet.getLastRow() - 1);
    const lastRun = parseInt(PropertiesService.getScriptProperties().getProperty('lastRun') || '0', 10);
    return jsonResp({ lastRun: lastRun, pendientes: pendientes, fallidos: fallidos, procesados30d: procesados30d });
  }

  // Default: return pending items
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const pending = [];
  for (var i = 1; i < rows.length; i++) {
    var id = rows[i][0], fecha = rows[i][1], desc = rows[i][2],
        monto = rows[i][3], tipo = rows[i][4], procesado = rows[i][5],
        posibleDup = rows[i][6];
    if (!procesado) {
      var fechaStr;
      if (fecha instanceof Date) {
        fechaStr = Utilities.formatDate(fecha, 'America/Santiago', 'yyyy-MM-dd');
      } else {
        fechaStr = String(fecha);
      }
      pending.push({
        id:        String(id),
        fecha:     fechaStr,
        desc:      String(desc),
        monto:     Number(monto),
        tipo:      String(tipo),
        posibleDup: !!posibleDup
      });
    }
  }
  return jsonResp(pending);
}

function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function markProcessed(sheet, id) {
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      sheet.getRange(i + 1, 6).setValue(true);
      return;
    }
  }
}

// ── Procesar emails nuevos ────────────────────────────────────────────
function processNewEmails() {
  var sheet = getSheet();
  var existingRows = sheet.getDataRange().getValues().slice(1);
  var existingIds = {};
  for (var r = 0; r < existingRows.length; r++) existingIds[String(existingRows[r][0])] = true;

  var fallidosSheet = getFallidosSheet();
  var fallidosRows = fallidosSheet.getDataRange().getValues().slice(1);
  var fallidosIds = {};
  for (var f = 0; f < fallidosRows.length; f++) fallidosIds[String(fallidosRows[f][0])] = true;

  // Filas recientes (últimas 48h) para detección de posibleDup
  var cutoff48h = Date.now() - 48 * 3600 * 1000;
  var recentRows = existingRows.filter(function(r) {
    var d = r[1] instanceof Date ? r[1].getTime() : new Date(String(r[1])).getTime();
    return d >= cutoff48h;
  });

  for (var pi = 0; pi < PARSERS.length; pi++) {
    var parser = PARSERS[pi];
    var fromClause = ' (' + parser.fromDomains.map(function(d) { return 'from:' + d; }).join(' OR ') + ')';

    for (var qi = 0; qi < parser.queries.length; qi++) {
      var q = parser.queries[qi];
      var threads = GmailApp.search(q + fromClause + ' newer_than:30d', 0, 50);

      for (var ti = 0; ti < threads.length; ti++) {
        var msgs = threads[ti].getMessages();
        for (var mi = 0; mi < msgs.length; mi++) {
          var msg = msgs[mi];
          var id = msg.getId();
          if (existingIds[id] || fallidosIds[id]) continue;

          var subject = msg.getSubject();
          var from    = msg.getFrom();
          var body    = msg.getPlainBody();
          var date    = msg.getDate();
          var emailFecha = Utilities.formatDate(date, 'America/Santiago', 'yyyy-MM-dd');

          // Excluir pagos de tarjeta de crédito antes de parsear
          if (/pago\s+(?:de\s+)?tarjeta\s+de\s+cr.dito|pago\s+tarjeta/i.test(subject)) {
            logFallido(fallidosSheet, id, emailFecha, from, subject, body, 'pago TC excluido');
            fallidosIds[id] = true;
            continue;
          }

          var tx = parser.parse(subject, body, emailFecha, date, id);

          if (!tx) {
            logFallido(fallidosSheet, id, emailFecha, from, subject, body, 'no parseado');
            fallidosIds[id] = true;
            continue;
          }

          // Detectar posible duplicado: mismo monto y fecha en últimas 48h con distinto id
          var posibleDup = recentRows.some(function(rr) {
            if (String(rr[0]) === id) return false;
            var rowFecha = rr[1] instanceof Date
              ? Utilities.formatDate(rr[1], 'America/Santiago', 'yyyy-MM-dd')
              : String(rr[1]);
            return Number(rr[3]) === tx.monto && rowFecha === tx.fecha;
          });

          sheet.appendRow([tx.id, tx.fecha, tx.desc, tx.monto, tx.tipo, false, posibleDup]);
          existingIds[id] = true;
          recentRows.push([tx.id, tx.fecha, tx.desc, tx.monto, tx.tipo, false, posibleDup]);
        }
      }
    }
  }

  // Archivar filas procesadas con más de 60 días
  archivarProcesadas(sheet);

  // Registrar hora de última ejecución
  PropertiesService.getScriptProperties().setProperty('lastRun', String(Date.now()));
}

function logFallido(sheet, id, fecha, from, subject, body, motivo) {
  sheet.appendRow([id, fecha, from, subject, body.substring(0, 500), motivo]);
}

function archivarProcesadas(sheet) {
  var archivo = getArchivoSheet();
  var rows = sheet.getDataRange().getValues();
  var cutoff60 = Date.now() - 60 * 24 * 3600 * 1000;
  var toDelete = [];
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!rows[i][5]) continue; // no procesado
    var d = rows[i][1] instanceof Date ? rows[i][1].getTime() : new Date(String(rows[i][1])).getTime();
    if (d < cutoff60) {
      archivo.appendRow(rows[i]);
      toDelete.push(i + 1); // 1-based
    }
  }
  // toDelete ya está en orden descendente (iteramos rows de abajo a arriba)
  for (var j = 0; j < toDelete.length; j++) {
    sheet.deleteRow(toDelete[j]);
  }
}

// ── Helpers de parseo ─────────────────────────────────────────────────
function fechaFromBodySafe(body, emailFecha, emailDate) {
  // Preferir fecha después de "el " para estar cerca del monto
  var nearEl = body.match(/el\s+(\d{2})\/(\d{2})\/(\d{4})/i);
  if (nearEl) {
    var candidate = nearEl[3] + '-' + nearEl[2] + '-' + nearEl[1];
    if (new Date(candidate) <= emailDate) return candidate;
  }
  // Fallback: primer dd/mm/yyyy en el cuerpo
  var m = body.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) {
    var cand2 = m[3] + '-' + m[2] + '-' + m[1];
    if (new Date(cand2) <= emailDate) return cand2;
  }
  return emailFecha;
}

function parseMonto(str) {
  return parseInt(str.replace(/\./g, '').replace(/,/g, ''), 10) || 0;
}

function clean(str) {
  return str.trim().replace(/\s+/g, ' ').substring(0, 60);
}

// ── Sheet helpers ─────────────────────────────────────────────────────
function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'fecha', 'desc', 'monto', 'tipo', 'procesado', 'posibleDup']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getFallidosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Fallidos');
  if (!sheet) {
    sheet = ss.insertSheet('Fallidos');
    sheet.appendRow(['id', 'fecha', 'remitente', 'asunto', 'cuerpo_500', 'motivo']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getArchivoSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Archivo');
  if (!sheet) {
    sheet = ss.insertSheet('Archivo');
    sheet.appendRow(['id', 'fecha', 'desc', 'monto', 'tipo', 'procesado', 'posibleDup']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Tests de parsers ──────────────────────────────────────────────────
// Ejecutar desde el menú Ejecutar → testParsers para verificar regex
function testParsers() {
  var tests = [
    {
      label: 'Banco de Chile – Cargo en Cuenta',
      subject: 'Cargo en Cuenta',
      from: 'notificaciones@enviodigital.bancochile.cl',
      body: 'Se ha realizado un cargo de $4.600 en FARMACIAS CRUZ VERDE el 01/06/2026.',
      date: new Date('2026-06-01')
    },
    {
      label: 'Banco de Chile – Compra Tarjeta Débito',
      subject: 'Compra con Tarjeta de Débito',
      from: 'notificaciones@enviodigital.bancochile.cl',
      body: 'Estimado cliente, usted realizó una compra por $22.000 en SUPERMERCADO TOTTUS el 10/06/2026.',
      date: new Date('2026-06-10')
    },
    {
      label: 'Banco de Chile – Transferencia entrante',
      subject: 'Has recibido una transferencia de fondos',
      from: 'notificaciones@bancochile.cl',
      body: 'Estimado cliente, ha recibido una transferencia de fondos de PEDRO GONZALEZ ROJAS\nMonto transferido: $150.000\nFecha: 05/06/2026',
      date: new Date('2026-06-05')
    },
    {
      label: 'Banco de Chile – Pago tarjeta (debe excluirse)',
      subject: 'Pago de Tarjeta de Crédito',
      from: 'notificaciones@enviodigital.bancochile.cl',
      body: 'Se realizó el pago de su Tarjeta de Crédito por $200.000.',
      date: new Date('2026-06-01')
    }
  ];

  for (var i = 0; i < tests.length; i++) {
    var t = tests[i];
    Logger.log('--- ' + t.label + ' ---');

    if (/pago\s+(?:de\s+)?tarjeta\s+de\s+cr.dito|pago\s+tarjeta/i.test(t.subject)) {
      Logger.log('EXCLUIDO: pago TC');
      continue;
    }

    var emailFecha = Utilities.formatDate(t.date, 'America/Santiago', 'yyyy-MM-dd');
    var parser = null;
    for (var p = 0; p < PARSERS.length; p++) {
      if (PARSERS[p].fromDomains.some(function(d) { return t.from.indexOf(d) !== -1; })) {
        parser = PARSERS[p];
        break;
      }
    }
    if (!parser) { Logger.log('Sin parser para from: ' + t.from); continue; }

    var result = parser.parse(t.subject, t.body, emailFecha, t.date, 'TEST_' + i);
    Logger.log(result ? JSON.stringify(result) : 'null → iría a Fallidos');
  }
}

// ── Configurar trigger automático ─────────────────────────────────────
// Ejecutar esta función UNA sola vez desde el menú Ejecutar
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'processNewEmails'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('Trigger creado: processNewEmails cada 15 minutos');
  Logger.log('Ejecutando primera revision de emails...');
  processNewEmails();
  Logger.log('Listo');
}
