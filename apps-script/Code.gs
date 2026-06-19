// ══════════════════════════════════════════════════════════════════════════════
// M-MOLA 10 ANOS — EVENT MANAGEMENT SYSTEM
// Google Apps Script Backend
// ══════════════════════════════════════════════════════════════════════════════
//
// SETUP:
//  1. Tạo Google Sheet mới → Extensions → Apps Script → dán code này
//  2. Deploy → New deployment → Web App
//     - Execute as: Me
//     - Who has access: Anyone
//  3. Copy Web App URL → dán vào APPS_SCRIPT_URL trong các file HTML
//  4. Chạy hàm setup() một lần để tạo sheet + điền 211 vé
// ══════════════════════════════════════════════════════════════════════════════

// ── CONFIG ────────────────────────────────────────────────────────────────────
const ADMIN_PIN     = '2026';   // PIN cho ban tổ chức (đổi trước ngày sự kiện)
const TOTAL_TICKETS = 211;      // Tổng số khách mời

// Sheet names
const SH_GUESTS  = 'Convidados';
const SH_WINNERS = 'Sorteio';

// ── HTTP: GET ──────────────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'dashboard';

  if (action === 'stats')    return jsonOut(getStats());
  if (action === 'checkins') return jsonOut(getCheckins());
  if (action === 'guest')    return jsonOut(getGuest(e.parameter.ticket));

  // Default → serve Dashboard
  return HtmlService
    .createHtmlOutputFromFile('Dashboard')
    .setTitle('M-Mola · Dashboard do Evento')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── HTTP: POST ─────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'rsvp':    return jsonOut(handleRSVP(data));
      case 'checkin': return jsonOut(handleCheckin(data));
      case 'winner':  return jsonOut(handleWinner(data));
      default:        return jsonOut({ error: 'Acção desconhecida' });
    }
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ── RSVP ──────────────────────────────────────────────────────────────────────
function handleRSVP(data) {
  const sheet  = getSheet(SH_GUESTS);
  const ticket = pad(data.ticket);
  const status = data.confirmed ? 'Confirmado' : 'Não vai';

  const row = findRow(sheet, ticket);
  if (!row) return { error: 'Bilhete não encontrado: ' + ticket };

  const current = sheet.getRange(row, 3).getValue();
  if (current !== 'Pendente') {
    return { success: true, alreadyResponded: true, status: current };
  }

  sheet.getRange(row, 3).setValue(status);
  sheet.getRange(row, 4).setValue(formatDate(new Date()));
  if (data.name) sheet.getRange(row, 2).setValue(data.name);

  return { success: true, ticket, status };
}

// ── CHECK-IN ──────────────────────────────────────────────────────────────────
function handleCheckin(data) {
  if (String(data.pin) !== String(ADMIN_PIN)) {
    return { error: 'PIN incorreto' };
  }

  const sheet  = getSheet(SH_GUESTS);
  const ticket = pad(data.ticket);
  const row    = findRow(sheet, ticket);

  if (!row) return { error: 'Bilhete não encontrado: ' + ticket };

  const values     = sheet.getRange(row, 1, 1, 7).getValues()[0];
  const alreadyIn  = values[4] === 'Sim';
  const name       = values[1] || '—';
  const rsvp       = values[2] || 'Pendente';

  if (alreadyIn) {
    return { success: true, ticket, name, rsvp, alreadyCheckedIn: true, checkinTime: values[5] };
  }

  sheet.getRange(row, 5).setValue('Sim');
  sheet.getRange(row, 6).setValue(formatDate(new Date()));

  return { success: true, ticket, name, rsvp, checkedIn: true };
}

// ── WINNER ────────────────────────────────────────────────────────────────────
function handleWinner(data) {
  if (String(data.pin) !== String(ADMIN_PIN)) {
    return { error: 'PIN incorreto' };
  }
  const sheet = getSheet(SH_WINNERS);
  sheet.appendRow([
    formatDate(new Date()),
    pad(data.ticket),
    data.tier,
    data.prize
  ]);
  return { success: true };
}

// ── GET CHECKINS ──────────────────────────────────────────────────────────────
function getCheckins() {
  const sheet  = getSheet(SH_GUESTS);
  const data   = sheet.getDataRange().getValues();
  const tickets = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === 'Sim') tickets.push(String(data[i][0]));
  }
  return { tickets, count: tickets.length };
}

// ── STATS ─────────────────────────────────────────────────────────────────────
function getStats() {
  const sheet = getSheet(SH_GUESTS);
  const data  = sheet.getDataRange().getValues();

  let total = 0, confirmed = 0, declined = 0, pending = 0, checkedIn = 0;
  const recent = [];

  for (let i = 1; i < data.length; i++) {
    total++;
    const rsvp = data[i][2];
    if      (rsvp === 'Confirmado') confirmed++;
    else if (rsvp === 'Não vai')    declined++;
    else                             pending++;
    if (data[i][4] === 'Sim') checkedIn++;
  }

  // Last 10 check-ins (reverse order)
  for (let i = data.length - 1; i >= 1 && recent.length < 10; i--) {
    if (data[i][4] === 'Sim') {
      recent.push({
        ticket: data[i][0],
        name:   data[i][1] || '—',
        rsvp:   data[i][2] || 'Pendente',
        time:   data[i][5] || ''
      });
    }
  }

  // Winners
  const ws = getSheet(SH_WINNERS);
  const wd = ws.getDataRange().getValues();
  const winners = [];
  for (let i = 1; i < wd.length; i++) {
    winners.push({ time: wd[i][0], ticket: wd[i][1], tier: wd[i][2], prize: wd[i][3] });
  }

  return { total, confirmed, declined, pending, checkedIn, recent, winners };
}

// ── GET SINGLE GUEST ──────────────────────────────────────────────────────────
function getGuest(ticket) {
  if (!ticket) return { error: 'Sem bilhete' };
  const sheet = getSheet(SH_GUESTS);
  const row   = findRow(sheet, pad(ticket));
  if (!row) return { error: 'Não encontrado' };
  const v = sheet.getRange(row, 1, 1, 7).getValues()[0];
  return {
    ticket:      v[0],
    name:        v[1] || '—',
    rsvp:        v[2] || 'Pendente',
    rsvpTime:    v[3] || '',
    checkedIn:   v[4] === 'Sim',
    checkinTime: v[5] || ''
  };
}

// ── SETUP (run once) ──────────────────────────────────────────────────────────
function setup() {
  setupGuestsSheet();
  setupWinnersSheet();
  Logger.log('✅ Setup completo! ' + TOTAL_TICKETS + ' bilhetes criados.');
}

function setupGuestsSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SH_GUESTS);
  if (sheet)  sheet.clear();
  else        sheet = ss.insertSheet(SH_GUESTS, 0);

  const headers = ['Nº Bilhete', 'Nome', 'RSVP', 'Data RSVP', 'Check-in', 'Data Check-in', 'Observações'];
  const hRange  = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight('bold').setBackground('#0D1B2E').setFontColor('#C9A84C').setFontSize(11);

  const rows = [];
  for (let i = 1; i <= TOTAL_TICKETS; i++) {
    rows.push([pad(i), '', 'Pendente', '', 'Não', '', '']);
  }
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);

  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 200);
  sheet.setFrozenRows(1);

  // Conditional formatting — RSVP status colors
  const rsvpRange  = sheet.getRange('C2:C' + (TOTAL_TICKETS + 1));
  const rules      = sheet.getConditionalFormatRules();
  const green      = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Confirmado').setBackground('#d4edda').setFontColor('#155724').setRanges([rsvpRange]).build();
  const red        = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Não vai').setBackground('#f8d7da').setFontColor('#721c24').setRanges([rsvpRange]).build();
  const checkinRange = sheet.getRange('E2:E' + (TOTAL_TICKETS + 1));
  const gold       = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Sim').setBackground('#fff3cd').setFontColor('#856404').setRanges([checkinRange]).build();
  sheet.setConditionalFormatRules([...rules, green, red, gold]);

  Logger.log('✅ Sheet "' + SH_GUESTS + '" criada com ' + TOTAL_TICKETS + ' bilhetes.');
}

function setupWinnersSheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let sheet  = ss.getSheetByName(SH_WINNERS);
  if (sheet) sheet.clear();
  else       sheet = ss.insertSheet(SH_WINNERS);

  const headers = ['Data/Hora', 'Nº Bilhete', 'Categoria', 'Prémio'];
  const hRange  = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight('bold').setBackground('#0D1B2E').setFontColor('#C9A84C').setFontSize(11);
  sheet.setColumnWidths(1, 4, 160);
  sheet.setFrozenRows(1);

  Logger.log('✅ Sheet "' + SH_WINNERS + '" criada.');
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function getSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(name);
  if (!sheet) {
    if (name === SH_GUESTS)  { setupGuestsSheet();  sheet = ss.getSheetByName(name); }
    if (name === SH_WINNERS) { setupWinnersSheet(); sheet = ss.getSheetByName(name); }
  }
  return sheet;
}

function findRow(sheet, ticket) {
  const col = sheet.getRange('A:A').getValues();
  for (let i = 1; i < col.length; i++) {
    if (String(col[i][0]) === String(ticket)) return i + 1;
  }
  return null;
}

function pad(n) { return String(n).padStart(3, '0'); }

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function jsonOut(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
