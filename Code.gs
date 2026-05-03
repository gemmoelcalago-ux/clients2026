// ════════════════════════════════════════════════════════
//  ST. PETER LIFE PLAN — GOOGLE APPS SCRIPT BACKEND
//  Code.gs  — Deploy as Web App (Execute as: Me, Access: Anyone)
// ════════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // auto-uses bound sheet
const ADMIN_PASSWORD = 'admin123'; // ← CHANGE THIS

// ── Sheet names ──
const SH_CLIENTS  = 'Clients';
const SH_PAYMENTS = 'Payments';
const SH_NOTES    = 'Notes';
const SH_PM       = 'PaymentMethods';

// ════════════════════════════════════════════════════════
//  WEB APP ENTRY POINTS
// ════════════════════════════════════════════════════════

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('St. Peter — Payment Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { action } = payload;

    switch (action) {
      case 'getClients':        return ok(getClients());
      case 'adminLogin':        return ok(adminLogin(payload));
      case 'clientLogin':       return ok(clientLogin(payload));
      case 'getPayments':       return ok(getPayments(payload));
      case 'setPayment':        return ok(setPayment(payload));
      case 'clearPayment':      return ok(clearPayment(payload));
      case 'getNotes':          return ok(getNotes(payload));
      case 'setNotes':          return ok(setNotes(payload));
      case 'getPaymentMethod':  return ok(getPaymentMethod(payload));
      case 'setPaymentMethod':  return ok(setPaymentMethod(payload));
      case 'getAllData':         return ok(getAllData(payload));
      default:                  return err('Unknown action: ' + action);
    }
  } catch (ex) {
    return err(ex.message);
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════
//  SHEET HELPERS
// ════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// Returns all rows as array of objects using header row
function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ════════════════════════════════════════════════════════
//  SETUP — run once to create sheets + seed data
// ════════════════════════════════════════════════════════

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Clients sheet ──
  let clients = ss.getSheetByName(SH_CLIENTS);
  if (!clients) {
    clients = ss.insertSheet(SH_CLIENTS);
    const headers = ['id','name','policy','plan','planType','batch','dob'];
    clients.appendRow(headers);
    clients.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#c9a84c')
      .setFontColor('#1a1610');

    const seed = [
      [1,'ILAGAN, KIMBERLY JANE','L26967488I','Monthly','ST Gregory',1,'06/28/1994'],
      [2,'BABATID, MCJOHN VELASCO','L26967489I','Monthly','ST Gregory',1,'11/17/1991'],
      [3,'VERDIDA, JUNJUN MABANAG','L26967491I','Monthly','ST Gregory',1,'09/14/1981'],
      [4,'VERDIDA, JHAYNE MARIE','L26967492I','Monthly','ST Gregory',1,'08/11/2007'],
      [5,'BABATID, JAYANN OLIVARES','L26967493I','Monthly','ST Gregory',1,'10/29/1986'],
      [6,'KING, ROMIE LLOYD CABRERA','L26967494I','Monthly','ST George',1,'07/29/1991'],
      [7,'LABUCA, GERARDO','L26967499I','Monthly','ST Gregory',1,'07/03/1970'],
      [8,'LABUCA, SALLY','L26967500I','Monthly','ST Gregory',1,'04/06/1970'],
      [9,'LABUCA, SAGE','L26967497I','Monthly','ST Gregory',1,'09/14/1999'],
      [10,'CALAGO, EULIE NEIL','L26967496I','Monthly','ST George',1,'04/30/1996'],
      [11,'CALAGO, SHIELA MAE','L26967495I','Monthly','ST George',1,'10/23/1988'],
      [12,'CALAGO, GEMMA','L26967490I','Annual','ST Gregory',1,'04/27/1970'],
      [13,'RIVERA, ERNA PATRIMONIO','L26967559I','Monthly','ST Gregory',2,'10/29/1965'],
      [14,'AMAD, REGINE MILLOR','L26967553I','Monthly','ST Gregory',2,'05/09/1996'],
      [15,'RIVERA, JOSELITO PELAYO','L26967558I','Monthly','ST Gregory',2,'08/27/1961'],
      [16,'ENRIQUEZ, FRAULEIN ANOVA','L26967557I','Quarterly','St Claire',2,'07/07/1995'],
      [17,'CALAGO, JACKIE NEIL MONTANER','L26967551I','Monthly','ST Gregory',2,'03/29/1999'],
      [18,'BABATID, ANITA OLIVARES','L26967555I','Monthly','ST Gregory',2,'04/17/1967'],
      [19,'TAN, RACHEAL SHEENA SUHURI','L26967552I','Monthly','ST Gregory',2,'10/03/1994'],
      [20,'RIVERA, KITH JHUNIEL PATRIMONIO','L26967556I','Monthly','ST Gregory',2,'12/16/1990'],
      [21,'AMAD, RESURRRECTION CATIRA','L26967560I','Monthly','ST Gregory',2,'03/28/1970'],
    ];
    seed.forEach(r => clients.appendRow(r));
    clients.autoResizeColumns(1, headers.length);
  }

  // ── Payments sheet ──
  let payments = ss.getSheetByName(SH_PAYMENTS);
  if (!payments) {
    payments = ss.insertSheet(SH_PAYMENTS);
    const headers = ['clientId','monthKey','paymentDate','recordedAt'];
    payments.appendRow(headers);
    payments.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#c9a84c')
      .setFontColor('#1a1610');
    payments.autoResizeColumns(1, headers.length);
  }

  // ── Notes sheet ──
  let notes = ss.getSheetByName(SH_NOTES);
  if (!notes) {
    notes = ss.insertSheet(SH_NOTES);
    const headers = ['clientId','notes','updatedAt'];
    notes.appendRow(headers);
    notes.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#c9a84c')
      .setFontColor('#1a1610');
    notes.autoResizeColumns(1, headers.length);
  }

  // ── PaymentMethods sheet ──
  let pm = ss.getSheetByName(SH_PM);
  if (!pm) {
    pm = ss.insertSheet(SH_PM);
    const headers = ['clientId','method','updatedAt'];
    pm.appendRow(headers);
    pm.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#c9a84c')
      .setFontColor('#1a1610');
    pm.autoResizeColumns(1, headers.length);
  }

  SpreadsheetApp.getUi().alert('✅ Setup complete! All sheets created and seeded.');
}

// ════════════════════════════════════════════════════════
//  ACTIONS
// ════════════════════════════════════════════════════════

function getClients() {
  const sh = getSheet(SH_CLIENTS);
  return sheetToObjects(sh).map(r => ({
    id:       Number(r.id),
    name:     r.name,
    policy:   r.policy,
    plan:     r.plan,
    planType: r.planType,
    batch:    Number(r.batch),
    dob:      r.dob,
  }));
}

function adminLogin({ password }) {
  if (password !== ADMIN_PASSWORD) throw new Error('Incorrect password.');
  return { ok: true };
}

function clientLogin({ clientId, dob }) {
  const clients = getClients();
  const client = clients.find(c => c.id === Number(clientId));
  if (!client) throw new Error('Client not found.');
  const normalize = s => String(s).replace(/\D/g, '');
  if (normalize(dob) !== normalize(client.dob)) throw new Error('Date of birth does not match.');
  return client;
}

// Returns { [monthKey]: dateString } for a given client
function getPayments({ clientId }) {
  const sh = getSheet(SH_PAYMENTS);
  const rows = sheetToObjects(sh);
  const result = {};
  rows.filter(r => Number(r.clientId) === Number(clientId))
      .forEach(r => { result[r.monthKey] = r.paymentDate; });
  return result;
}

function setPayment({ clientId, monthKey, paymentDate }) {
  const sh = getSheet(SH_PAYMENTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cidIdx = headers.indexOf('clientId');
  const mkIdx  = headers.indexOf('monthKey');
  const dateIdx = headers.indexOf('paymentDate');
  const tsIdx  = headers.indexOf('recordedAt');

  // Find existing row
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][cidIdx]) === Number(clientId) && data[i][mkIdx] === monthKey) {
      sh.getRange(i + 1, dateIdx + 1).setValue(paymentDate);
      sh.getRange(i + 1, tsIdx + 1).setValue(new Date().toISOString());
      return { updated: true };
    }
  }
  // Insert new row
  sh.appendRow([clientId, monthKey, paymentDate, new Date().toISOString()]);
  return { inserted: true };
}

function clearPayment({ clientId, monthKey }) {
  const sh = getSheet(SH_PAYMENTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cidIdx = headers.indexOf('clientId');
  const mkIdx  = headers.indexOf('monthKey');

  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][cidIdx]) === Number(clientId) && data[i][mkIdx] === monthKey) {
      sh.deleteRow(i + 1);
      return { deleted: true };
    }
  }
  return { notFound: true };
}

function getNotes({ clientId }) {
  const sh = getSheet(SH_NOTES);
  const rows = sheetToObjects(sh);
  const row = rows.find(r => Number(r.clientId) === Number(clientId));
  return { notes: row ? row.notes : '' };
}

function setNotes({ clientId, notes }) {
  const sh = getSheet(SH_NOTES);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cidIdx   = headers.indexOf('clientId');
  const notesIdx = headers.indexOf('notes');
  const tsIdx    = headers.indexOf('updatedAt');

  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][cidIdx]) === Number(clientId)) {
      sh.getRange(i + 1, notesIdx + 1).setValue(notes);
      sh.getRange(i + 1, tsIdx + 1).setValue(new Date().toISOString());
      return { updated: true };
    }
  }
  sh.appendRow([clientId, notes, new Date().toISOString()]);
  return { inserted: true };
}

function getPaymentMethod({ clientId }) {
  const sh = getSheet(SH_PM);
  const rows = sheetToObjects(sh);
  const row = rows.find(r => Number(r.clientId) === Number(clientId));
  return { method: row ? row.method : '' };
}

function setPaymentMethod({ clientId, method }) {
  const sh = getSheet(SH_PM);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cidIdx    = headers.indexOf('clientId');
  const methodIdx = headers.indexOf('method');
  const tsIdx     = headers.indexOf('updatedAt');

  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][cidIdx]) === Number(clientId)) {
      sh.getRange(i + 1, methodIdx + 1).setValue(method);
      sh.getRange(i + 1, tsIdx + 1).setValue(new Date().toISOString());
      return { updated: true };
    }
  }
  sh.appendRow([clientId, method, new Date().toISOString()]);
  return { inserted: true };
}

// Bulk fetch all data for admin dashboard in one call
function getAllData({ password }) {
  if (password !== ADMIN_PASSWORD) throw new Error('Unauthorized');
  const clients = getClients();

  const paySh = getSheet(SH_PAYMENTS);
  const payRows = sheetToObjects(paySh);
  const payments = {};
  payRows.forEach(r => {
    const cid = Number(r.clientId);
    if (!payments[cid]) payments[cid] = {};
    payments[cid][r.monthKey] = r.paymentDate;
  });

  const notesSh = getSheet(SH_NOTES);
  const notes = {};
  sheetToObjects(notesSh).forEach(r => { notes[Number(r.clientId)] = r.notes; });

  const pmSh = getSheet(SH_PM);
  const methods = {};
  sheetToObjects(pmSh).forEach(r => { methods[Number(r.clientId)] = r.method; });

  return { clients, payments, notes, methods };
}

// ════════════════════════════════════════════════════════
//  MENU (run from spreadsheet UI)
// ════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⛪ St. Peter Tracker')
    .addItem('🔧 Setup Sheets (run first)', 'setupSheets')
    .addItem('🌐 Open Web App', 'openWebApp')
    .addToUi();
}

function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(`<script>window.open('${url}','_blank');google.script.host.close();</script>`);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening…');
}
