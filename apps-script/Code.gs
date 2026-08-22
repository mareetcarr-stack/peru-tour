/**
 * Peru Tour Manager — Apps Script backend.
 *
 * Bound to the trip Google Sheet. Deployed as a Web App ("Execute as: Me",
 * "Who has access: Anyone with the link") this gives the static PWA a
 * free JSON API with no service-account key needed at runtime.
 *
 * SETUP (one-time):
 *   1. Open the Sheet → Extensions → Apps Script.
 *   2. Delete the placeholder code, paste this whole file in.
 *   3. Run `setPin` once from the editor toolbar (select it in the function
 *      dropdown → Run) to set the PIN. Change '1234' below first if you want
 *      something else — you can also change it later without redeploying,
 *      see setPin().
 *   4. Deploy → New deployment → type "Web app" →
 *        Execute as: Me
 *        Who has access: Anyone
 *      Click Deploy, authorize the permissions prompt (that's you granting
 *      your own script access to your own Sheet — normal and expected).
 *   5. Copy the "Web app URL" (ends in /exec) into CONFIG.APPS_SCRIPT_URL
 *      in index.html.
 *
 * Re-deploying: if you edit this file later, use Deploy → Manage deployments
 * → edit (pencil) → New version, so the existing /exec URL keeps working.
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────
const SHEET_NAMES = {
  checkin: 'CHECK-IN STATUS',
  checklist: 'DAILY CHECKLIST',
  tours: 'TOURS',
  tickets: 'TICKETS',
  info: 'B-DAYS, AGES, DIET',
  flights: 'FLIGHT STATUS',
  recs: 'RECOMMENDATIONS',
};

function setPin() {
  PropertiesService.getScriptProperties().setProperty('PIN', '1234');
}

// ─── ENTRY POINTS ────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'login') return respond_(login_(e.parameter.pin));
    if (action === 'data') {
      requireToken_(e.parameter.token);
      return respond_({ ok: true, data: getAllData_() });
    }
    return respond_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    // Body is sent as text/plain (not application/json) on purpose — a
    // JSON content-type triggers a CORS preflight OPTIONS request, which
    // Apps Script web apps cannot answer, and the browser would block the
    // real request. Parsing the JSON ourselves from postData.contents
    // sidesteps that entirely.
    const body = JSON.parse(e.postData.contents);
    requireToken_(body.token);
    const result = handleWrite_(body);
    return respond_({ ok: true, result: result });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── AUTH ────────────────────────────────────────────────────────────────
function login_(pin) {
  const props = PropertiesService.getScriptProperties();
  const realPin = props.getProperty('PIN') || '1234';
  if (String(pin) !== String(realPin)) return { ok: false, error: 'Wrong PIN' };
  const token = Utilities.getUuid();
  const sessions = JSON.parse(props.getProperty('SESSIONS') || '{}');
  sessions[token] = Date.now();
  // Keep the sessions list from growing forever — drop anything older than 180 days.
  const cutoff = Date.now() - 180 * 24 * 60 * 60 * 1000;
  Object.keys(sessions).forEach((t) => { if (sessions[t] < cutoff) delete sessions[t]; });
  props.setProperty('SESSIONS', JSON.stringify(sessions));
  return { ok: true, token: token };
}

function requireToken_(token) {
  if (!token) throw new Error('Missing token');
  const sessions = JSON.parse(PropertiesService.getScriptProperties().getProperty('SESSIONS') || '{}');
  if (!sessions[token]) throw new Error('Invalid or expired session — please enter the PIN again');
}

// ─── HELPERS ─────────────────────────────────────────────────────────────
function sheet_(key) {
  const name = SHEET_NAMES[key];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet tab not found: ' + name);
  return sh;
}

function bool_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
}

function parseMoney_(v) {
  if (!v) return 0;
  const n = parseFloat(String(v).replace(/[^0-9.,-]/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function fullName_(first, second) {
  return [first, second].filter(Boolean).join(' ').trim();
}

// ─── READ: full data blob ────────────────────────────────────────────────
function getAllData_() {
  return {
    passengers: getCheckin_(),
    flights: getFlights_(),
    checklist: getChecklist_(),
    tours: getTours_(),
    tickets: getTickets_(),
    info: getInfo_(),
    places: getRecs_(),
  };
}

function getCheckin_() {
  const rows = sheet_('checkin').getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (row[1] === '' && row[2] === '') continue; // no name → skip
    out.push({
      id: String(row[0]),
      first: row[1], second: row[2], name: fullName_(row[1], row[2]),
      arrived: bool_(row[3]),
      pnr: row[4], passport: row[5], passportExpiry: row[6], orderId: row[7],
      checkinLink: row[8], boardingStatus: row[9] || '',
      actualArrival: row[10] || '', errorFlag: row[11] || '', sclLimDone: row[12] || '',
    });
  }
  return out;
}

function getFlights_() {
  const rows = sheet_('flights').getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row[1]) continue; // no flight number → skip
    out.push({
      date: row[0], flightNumber: row[1], link: row[2] || '',
      status: row[3] || '', departure: row[4] || '', arrival: row[5] || '',
    });
  }
  return out;
}

function getChecklist_() {
  const rows = sheet_('checklist').getDataRange().getValues();
  const days = [];
  let current = null;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const a = String(row[0] || '').trim();
    if (/^Day\s+\d+/i.test(a)) {
      current = { day: a, title: row[1] || '', items: [] };
      days.push(current);
    } else if ((a.toUpperCase() === 'TRUE' || a.toUpperCase() === 'FALSE') && current) {
      current.items.push({ row: r + 1, text: row[1] || '', done: a.toUpperCase() === 'TRUE' });
    }
  }
  return days;
}

// Tour column pairs start at column index 4 (0-based) and run until the
// "TOTAL" header is hit. Each pair is [TripADeal col, Valencia col].
function tourDefs_(headers) {
  const defs = [];
  let i = 4;
  while (i < headers.length && String(headers[i]).trim().toUpperCase() !== 'TOTAL') {
    const name = String(headers[i]).replace(/\s*TripA[Dd]eal\s*$/i, '').trim();
    defs.push({ name: name, colT: i, colV: i + 1 });
    i += 2;
  }
  return defs;
}

function getTours_() {
  const rows = sheet_('tours').getDataRange().getValues();
  const headers = rows[0];
  const priceRow = rows[1] || [];
  const defs = tourDefs_(headers).map((d) => ({
    name: d.name, colT: d.colT, colV: d.colV,
    price: parseMoney_(priceRow[d.colT]) || parseMoney_(priceRow[d.colV]) || 0,
  }));
  const totalCol = headers.findIndex((h) => String(h).trim().toUpperCase() === 'TOTAL');
  const paidCol = headers.findIndex((h) => String(h).trim().toUpperCase() === 'PAID');

  const out = [];
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    if (row[1] === '' && row[2] === '') continue;
    out.push({
      id: String(row[0]), name: fullName_(row[1], row[2]),
      tip: bool_(row[3]),
      tours: defs.map((d) => ({ name: d.name, t: bool_(row[d.colT]), v: bool_(row[d.colV]) })),
      total: totalCol >= 0 ? row[totalCol] : '',
      paid: paidCol >= 0 ? bool_(row[paidCol]) : false,
    });
  }
  return { tourDefs: defs.map((d) => ({ name: d.name, price: d.price })), rows: out, tipPrice: parseMoney_(priceRow[3]) };
}

function getTickets_() {
  const rows = sheet_('tickets').getDataRange().getValues();
  const headers = rows[0];
  const dateRow = rows[1] || [];
  const timeRow = rows[2] || [];
  const defs = [];
  for (let i = 4; i < headers.length; i++) {
    if (!headers[i]) continue;
    defs.push({ col: i, title: String(headers[i]).trim(), date: dateRow[i] || '', time: timeRow[i] || '' });
  }
  const out = [];
  for (let r = 3; r < rows.length; r++) {
    const row = rows[r];
    if (row[1] === '' && row[2] === '') continue;
    out.push({
      id: String(row[0]), name: fullName_(row[1], row[2]), passport: row[3] || '',
      checks: defs.map((d) => ({ col: d.col, checked: bool_(row[d.col]) })),
    });
  }
  return { ticketDefs: defs, rows: out };
}

function getInfo_() {
  const rows = sheet_('info').getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (row[1] === '' && row[2] === '') continue;
    out.push({
      id: String(row[0]), name: fullName_(row[1], row[2]),
      birthday: row[3] || '', age: row[4] || '', diet: row[5] || '',
    });
  }
  return out;
}

function getRecs_() {
  const rows = sheet_('recs').getDataRange().getValues();
  const restaurants = [];
  const sightseeing = [];
  const parse = (cell) => {
    const m = String(cell || '').match(/in\s+([^:]+):\s*(https?:\/\/\S+)/i);
    if (!m) return null;
    return { city: m[1].trim(), url: m[2].replace(/[￼ \s]+$/, '') };
  };
  for (let r = 1; r < rows.length; r++) {
    const rest = parse(rows[r][0]);
    if (rest) restaurants.push(rest);
    const sight = parse(rows[r][2]);
    if (sight) sightseeing.push(sight);
  }
  return { restaurants: restaurants, sightseeing: sightseeing };
}

// ─── WRITE ───────────────────────────────────────────────────────────────
function handleWrite_(body) {
  switch (body.action) {
    case 'toggleChecklist': return toggleChecklistItem_(body.row, body.done);
    case 'toggleTourTick': return toggleTourTick_(body.id, body.tourIndex, body.which, body.value);
    case 'toggleTip': return toggleTip_(body.id, body.value);
    case 'setPaid': return setPaid_(body.id, body.value);
    case 'toggleTicket': return toggleTicket_(body.id, body.col, body.value);
    case 'setBirthday': return setBirthday_(body.id, body.value);
    case 'setDiet': return setDiet_(body.id, body.value);
    default: throw new Error('Unknown write action: ' + body.action);
  }
}

function findRowById_(sh, id, startRow) {
  const rows = sh.getDataRange().getValues();
  for (let r = startRow; r < rows.length; r++) {
    if (String(rows[r][0]) === String(id)) return r + 1; // 1-based sheet row
  }
  throw new Error('Passenger id not found: ' + id);
}

function toggleChecklistItem_(row, done) {
  sheet_('checklist').getRange(row, 1).setValue(done ? 'TRUE' : 'FALSE');
  return { row: row, done: done };
}

function toggleTourTick_(id, tourIndex, which, value) {
  const sh = sheet_('tours');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const defs = tourDefs_(headers);
  const def = defs[tourIndex];
  if (!def) throw new Error('Unknown tour index: ' + tourIndex);
  const col = which === 't' ? def.colT : def.colV;
  const row = findRowById_(sh, id, 2);
  sh.getRange(row, col + 1).setValue(value ? 'TRUE' : 'FALSE');
  return { id: id, tourIndex: tourIndex, which: which, value: value };
}

function toggleTip_(id, value) {
  const sh = sheet_('tours');
  const row = findRowById_(sh, id, 2);
  sh.getRange(row, 4).setValue(value ? 'TRUE' : 'FALSE'); // col D = Tips
  return { id: id, value: value };
}

function setPaid_(id, value) {
  const sh = sheet_('tours');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const paidCol = headers.findIndex((h) => String(h).trim().toUpperCase() === 'PAID');
  if (paidCol < 0) throw new Error('PAID column not found');
  const row = findRowById_(sh, id, 2);
  sh.getRange(row, paidCol + 1).setValue(value ? 'TRUE' : 'FALSE');
  return { id: id, value: value };
}

function toggleTicket_(id, col, value) {
  const sh = sheet_('tickets');
  const row = findRowById_(sh, id, 3);
  sh.getRange(row, col + 1).setValue(value ? 'TRUE' : 'FALSE');
  return { id: id, col: col, value: value };
}

function setBirthday_(id, value) {
  const sh = sheet_('info');
  const row = findRowById_(sh, id, 1);
  sh.getRange(row, 4).setValue(value); // col D = Birthday
  return { id: id, value: value };
}

function setDiet_(id, value) {
  const sh = sheet_('info');
  const row = findRowById_(sh, id, 1);
  sh.getRange(row, 6).setValue(value); // col F = Dietary Requirements
  return { id: id, value: value };
}
