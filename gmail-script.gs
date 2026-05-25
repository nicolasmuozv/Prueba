// ══════════════════════════════════════
//  WalletWolf – Gmail Import Script
//  script.google.com → Nuevo proyecto → pega este código
// ══════════════════════════════════════
//
//  SETUP (hacer una sola vez):
//  1. Abre script.google.com, crea un proyecto nuevo
//  2. Crea también una Google Sheet en drive.google.com
//     y copia su ID de la URL (la parte larga entre /d/ y /edit)
//  3. Pega ese ID en SHEET_ID abajo
//  4. Pega todo este código en el editor
//  5. Ejecuta setupTrigger() (menú Ejecutar → setupTrigger)
//     → acepta los permisos de Gmail, Sheets y ScriptApp
//  6. Menú Implementar → Nueva implementación
//     Tipo: Aplicación web
//     Ejecutar como: Yo
//     Quién puede acceder: Cualquiera
//     → Implementar → copia la URL que aparece
//  7. En WalletWolf → "Importar del banco" → pega esa URL
// ══════════════════════════════════════

const SHEET_ID   = 'PEGA_AQUI_EL_ID_DE_TU_GOOGLE_SHEET';
const SHEET_NAME = 'Pendientes';

// ── Web App: GET ──────────────────────
function doGet(e) {
  const sheet = getSheet();
  const action = e && e.parameter ? e.parameter.action : null;

  if (action === 'confirm') {
    const id = e.parameter.id || '';
    if (id) markProcessed(sheet, id);
    return jsonResp({ ok: true });
  }

  const rows = sheet.getDataRange().getValues();
  const pending = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, fecha, desc, monto, tipo, procesado] = rows[i];
    if (!procesado) {
      pending.push({
        id:    String(id),
        fecha: String(fecha),
        desc:  String(desc),
        monto: Number(monto),
        tipo:  String(tipo)
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
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      sheet.getRange(i + 1, 6).setValue(true);
      return;
    }
  }
}

// ── Procesar emails nuevos ────────────
function processNewEmails() {
  const sheet = getSheet();
  const existingIds = sheet.getDataRange().getValues().slice(1).map(r => String(r[0]));

  const queries = [
    // Débitos / compras Banco de Chile
    'subject:"Cargo en Cuenta"',
    'subject:"Compra con Tarjeta de Crédito"',
    'subject:"Compra con Tarjeta de Débito"',
    // Transferencias salientes
    'subject:"Comprobante de Transferencia a terceros"',
    'subject:"Transferencia realizada"',
    // Transferencias entrantes (cualquier banco)
    'subject:"Has recibido una transferencia de fondos"',
    'subject:"Transferencia recibida"',
    // Abonos
    'subject:"Abono en Cuenta"',
  ];

  for (const q of queries) {
    const threads = GmailApp.search(q + ' newer_than:30d', 0, 50);
    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        const id = msg.getId();
        if (existingIds.includes(id)) continue;

        const tx = parseTx(msg.getSubject(), msg.getPlainBody(), msg.getDate(), id);
        if (tx) {
          sheet.appendRow([tx.id, tx.fecha, tx.desc, tx.monto, tx.tipo, false]);
          existingIds.push(id);
        }
      }
    }
  }
}

// ── Parser de emails ─────────────────
function parseTx(subject, body, date, id) {
  const fecha = Utilities.formatDate(date, 'America/Santiago', 'yyyy-MM-dd');

  // Cargo en cuenta (débito)
  if (/cargo en cuenta/i.test(subject)) {
    const m = body.match(/por\s*\$\s*([\d.,]+)/i);
    const d = body.match(/en\s+([A-ZÁÉÍÓÚÜÑ][A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,&'-]+?)\s+el\s+\d{2}\/\d{2}/i);
    if (m) return { id, fecha, desc: d ? clean(d[1]) : 'Cargo en cuenta', monto: parseMonto(m[1]), tipo: 'gasto' };
  }

  // Compra tarjeta crédito o débito
  if (/compra con tarjeta/i.test(subject)) {
    const m = body.match(/por\s*\$\s*([\d.,]+)/i);
    const d = body.match(/en\s+([A-ZÁÉÍÓÚÜÑ][A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,&'-]+?)\s+el\s+\d{2}\/\d{2}/i);
    if (m) return { id, fecha, desc: d ? clean(d[1]) : 'Compra con tarjeta', monto: parseMonto(m[1]), tipo: 'gasto' };
  }

  // Transferencia saliente
  if (/comprobante de transferencia|transferencia a terceros|transferencia realizada/i.test(subject)) {
    const m = body.match(/Monto[\s:$]*([0-9.,]+)/i);
    const d = body.match(/(?:Nombre|Destinatario|Beneficiario)[:\s]+([A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,]+)/i);
    if (m) return { id, fecha, desc: d ? 'TRF → ' + clean(d[1]) : 'Transferencia saliente', monto: parseMonto(m[1]), tipo: 'gasto' };
  }

  // Transferencia entrante
  if (/recibido una transferencia|transferencia recibida/i.test(subject)) {
    const m = body.match(/Monto transferido[:\s$]*([0-9.,]+)/i)
           || body.match(/Monto[:\s$]*([0-9.,]+)/i);
    const d = body.match(/transferencia de fondos de\s+([A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,]+?)[\r\n]/i)
           || body.match(/de\s+([A-Z][A-Za-z0-9 ÁÉÍÓÚÜÑáéíóúüñ.,]+?)\s+(?:Monto|RUT|por)/i);
    if (m) return { id, fecha, desc: d ? 'TRF ← ' + clean(d[1]) : 'Transferencia entrante', monto: parseMonto(m[1]), tipo: 'ingreso' };
  }

  // Abono en cuenta
  if (/abono en cuenta/i.test(subject)) {
    const m = body.match(/\$\s*([\d.,]+)/i);
    if (m) return { id, fecha, desc: 'Abono en cuenta', monto: parseMonto(m[1]), tipo: 'ingreso' };
  }

  return null;
}

function parseMonto(str) {
  // Chilean: 4.600 → 4600   Mixed: 22,000 → 22000
  return parseInt(str.replace(/\./g, '').replace(/,/g, ''), 10) || 0;
}

function clean(str) {
  return str.trim().replace(/\s+/g, ' ').substring(0, 60);
}

// ── Helpers ───────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'fecha', 'desc', 'monto', 'tipo', 'procesado']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Configurar trigger automático ─────
// Ejecuta esta función UNA sola vez desde el menú Ejecutar
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processNewEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('✓ Trigger creado: processNewEmails cada 15 minutos');
  Logger.log('Ejecutando primera revisión de emails...');
  processNewEmails();
  Logger.log('✓ Listo');
}
