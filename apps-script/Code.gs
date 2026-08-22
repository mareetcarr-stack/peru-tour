/**
 * Peru Tour Manager — Apps Script backend.
 *
 * Bound to the trip Google Sheet. Deployed as a Web App ("Execute as: Me",
 * "Who has access: Anyone with the link") this gives the static PWA a
 * free JSON API with no service-account key needed at runtime.
 *
 * No PIN / auth — anyone with the deployed URL can read and write. That's a
 * deliberate simplification (removed at the owner's request); the URL
 * itself is the only thing gating access.
 *
 * SETUP (one-time):
 *   1. Open the Sheet → Extensions → Apps Script.
 *   2. Delete the placeholder code, paste this whole file in.
 *   3. Deploy → New deployment → type "Web app" →
 *        Execute as: Me
 *        Who has access: Anyone
 *      Click Deploy, authorize the permissions prompt (that's you granting
 *      your own script access to your own Sheet — normal and expected).
 *   4. Copy the "Web app URL" (ends in /exec) into CONFIG.APPS_SCRIPT_URL
 *      in index.html.
 *
 * Re-deploying: if you edit this file later, use Deploy → Manage deployments
 * → edit (pencil) → New version, so the existing /exec URL keeps working.
 *
 * IMPORTANT sheet change needed once: the "Valencia" (V) column and "Tips"
 * column in the TOURS tab used to be plain checkboxes (TRUE/FALSE). They now
 * store a currency code + amount instead (e.g. "USD" or "PEN:120.00"), so
 * remove the checkbox validation on those columns in the Sheet (select the
 * columns → Data → Data validation → Remove rule) so Sheets doesn't fight
 * the app over what's a valid value there. The app clears it automatically
 * on write too, as a safety net.
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
  receipts: 'RECEIPTS',
};

// ─── ENTRY POINTS ────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'data') return respond_({ ok: true, data: getAllData_() });
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

// Sheets silently promotes some typed cells (e.g. "24/02/36") to real Date
// objects. JSON.stringify then serializes those via Date#toJSON, which
// applies UTC + JS's 2-digit-year rule (36 → 1936, not 2036) — garbage for
// display. Format any Date back to a plain d/M/yyyy string; everything else
// passes through untouched.
function cellStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return v == null ? '' : v;
}

// Same idea, but formatted for an HTML <input type="date"> (needs yyyy-MM-dd).
// Used only for the Birthday column, which round-trips through that input.
function isoDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (!v) return '';
  return String(v);
}

// V (Valencia) and Tips cells store a currency code + optional amount as
// plain text: "" (not booked), "USD" (auto-priced, no amount needed),
// "PEN:120.00" (manual amount). Old boolean TRUE values from before this
// change are treated as "booked, currency not yet chosen" → USD default.
function parseCC_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s || s.toUpperCase() === 'FALSE') return { currency: '', amount: 0 };
  if (s.toUpperCase() === 'TRUE') return { currency: 'USD', amount: 0 };
  const m = s.match(/^([A-Za-z]{3})(?::([0-9.]+))?$/);
  if (m) return { currency: m[1].toUpperCase(), amount: parseFloat(m[2]) || 0 };
  return { currency: '', amount: 0 };
}
function encodeCC_(currency, amount) {
  if (!currency) return '';
  if (currency === 'USD') return 'USD';
  return currency + ':' + (parseFloat(amount) || 0).toFixed(2);
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
    receipts: getReceipts_(),
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
      pnr: row[4], passport: row[5], passportExpiry: cellStr_(row[6]), orderId: row[7],
      checkinLink: row[8], boardingStatus: row[9] || '',
      actualArrival: cellStr_(row[10]) || '', errorFlag: row[11] || '', sclLimDone: row[12] || '',
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
    const flightNumber = String(row[1]).trim();
    out.push({
      date: cellStr_(row[0]), flightNumber: flightNumber,
      // Column C is a HYPERLINK() formula whose display text is just "View" —
      // getValues() only returns that display text, losing the URL. The
      // underlying link always follows this exact pattern, so build it
      // directly rather than trying to pull rich-text/formula data out.
      link: flightNumber ? 'https://es.flightaware.com/live/flight/' + encodeURIComponent(flightNumber) : '',
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
    const aRaw = row[0];
    const a = String(aRaw == null ? '' : aRaw).trim();
    const text = String(row[1] || '').trim();
    if (/^Day\s+\d+/i.test(a)) {
      current = { day: a, title: text, items: [] };
      days.push(current);
    } else if (current && text) {
      // A checkbox that has never been touched comes back as an empty
      // string, not FALSE — checking strictly for "TRUE"/"FALSE" text was
      // silently dropping every never-checked item (which was most of the
      // future days). Any row with item text under a day counts as an
      // item regardless of what's literally in the checkbox cell; bool_()
      // already treats blank/anything-but-TRUE as not done.
      current.items.push({ row: r + 1, text: text, done: bool_(aRaw) });
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
      tip: parseCC_(row[3]),
      tours: defs.map((d) => ({ name: d.name, t: bool_(row[d.colT]), v: parseCC_(row[d.colV]) })),
      total: totalCol >= 0 ? row[totalCol] : '',
      paid: paidCol >= 0 ? bool_(row[paidCol]) : false,
    });
  }
  return {
    tourDefs: defs.map((d) => ({ name: d.name, price: d.price })),
    rows: out,
    tipPrice: parseMoney_(priceRow[3]),
  };
}

function getTickets_() {
  const rows = sheet_('tickets').getDataRange().getValues();
  const headers = rows[0];
  const dateRow = rows[1] || [];
  const timeRow = rows[2] || [];
  const defs = [];
  for (let i = 4; i < headers.length; i++) {
    if (!headers[i]) continue;
    defs.push({ col: i, title: String(headers[i]).trim(), date: cellStr_(dateRow[i]) || '', time: cellStr_(timeRow[i]) || '' });
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
      birthday: isoDate_(row[3]), age: row[4] || '', diet: row[5] || '',
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
    return { city: m[1].trim(), url: m[2].replace(/[￼ \s]+$/, '') };
  };
  for (let r = 1; r < rows.length; r++) {
    const rest = parse(rows[r][0]);
    if (rest) restaurants.push(rest);
    const sight = parse(rows[r][2]);
    if (sight) sightseeing.push(sight);
  }
  return { restaurants: restaurants, sightseeing: sightseeing };
}

// Reads the actual dropdown list off an existing validated cell, so the
// app always offers exactly the same options as the Sheet itself — no
// hardcoded list to fall out of sync if someone edits the validation rule.
function getDropdownOptions_(sheet, row, col, fallback) {
  try {
    const dv = sheet.getRange(row, col).getDataValidation();
    if (!dv) return fallback;
    const values = dv.getCriteriaValues();
    return Array.isArray(values[0]) && values[0].length ? values[0] : fallback;
  } catch (e) {
    return fallback;
  }
}

function getReceipts_() {
  const sh = sheet_('receipts');
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const description = String(row[0] || '').trim();
    if (!description) continue; // skip blank rows
    out.push({
      row: r + 1,
      description: description,
      observation: String(row[1] || '').trim(),
      paymentMethod: String(row[2] || '').trim(),
      currency: String(row[3] || '').trim(),
      amount: parseMoney_(row[4]),
    });
  }
  return {
    rows: out,
    paymentMethods: getDropdownOptions_(sh, 2, 3, ['Deposito en cuenta', 'Efectivo - en los demas casos']),
    currencies: getDropdownOptions_(sh, 2, 4, ['PEN', 'USD']),
  };
}

// ─── WRITE ───────────────────────────────────────────────────────────────
function handleWrite_(body) {
  switch (body.action) {
    case 'toggleChecklist': return toggleChecklistItem_(body.row, body.done);
    case 'toggleArrived': return toggleArrived_(body.id, body.value);
    case 'toggleTourTick': return toggleTourTick_(body.id, body.tourIndex, body.value);
    case 'setTourValencia': return setTourValencia_(body.id, body.tourIndex, body.currency, body.amount);
    case 'setTip': return setTip_(body.id, body.currency, body.amount);
    case 'setTourPrice': return setTourPrice_(body.tourIndex, body.price);
    case 'setTipPrice': return setTipPrice_(body.price);
    case 'setPaid': return setPaid_(body.id, body.value);
    case 'toggleTicket': return toggleTicket_(body.id, body.col, body.value);
    case 'setBirthday': return setBirthday_(body.id, body.value);
    case 'setDiet': return setDiet_(body.id, body.value);
    case 'addReceipt': return addReceipt_(body.description, body.observation, body.paymentMethod, body.currency, body.amount);
    case 'setReceiptField': return setReceiptField_(body.row, body.field, body.value);
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

function toggleArrived_(id, value) {
  const sh = sheet_('checkin');
  const row = findRowById_(sh, id, 1);
  sh.getRange(row, 4).setValue(value ? 'TRUE' : 'FALSE'); // col D = Arrived
  return { id: id, value: value };
}

function toggleTourTick_(id, tourIndex, value) {
  const sh = sheet_('tours');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const defs = tourDefs_(headers);
  const def = defs[tourIndex];
  if (!def) throw new Error('Unknown tour index: ' + tourIndex);
  const row = findRowById_(sh, id, 2);
  sh.getRange(row, def.colT + 1).setValue(value ? 'TRUE' : 'FALSE');
  return { id: id, tourIndex: tourIndex, value: value };
}

function setTourValencia_(id, tourIndex, currency, amount) {
  const sh = sheet_('tours');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const defs = tourDefs_(headers);
  const def = defs[tourIndex];
  if (!def) throw new Error('Unknown tour index: ' + tourIndex);
  const row = findRowById_(sh, id, 2);
  const cell = sh.getRange(row, def.colV + 1);
  cell.clearDataValidations(); // this column used to be a checkbox; stop Sheets fighting the new text value
  cell.setValue(encodeCC_(currency, amount));
  return { id: id, tourIndex: tourIndex, currency: currency, amount: amount };
}

function setTip_(id, currency, amount) {
  const sh = sheet_('tours');
  const row = findRowById_(sh, id, 2);
  const cell = sh.getRange(row, 4); // col D = Tips
  cell.clearDataValidations();
  cell.setValue(encodeCC_(currency, amount));
  return { id: id, currency: currency, amount: amount };
}

function setTourPrice_(tourIndex, price) {
  const sh = sheet_('tours');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const defs = tourDefs_(headers);
  const def = defs[tourIndex];
  if (!def) throw new Error('Unknown tour index: ' + tourIndex);
  sh.getRange(2, def.colV + 1).setValue(parseFloat(price) || 0);
  return { tourIndex: tourIndex, price: price };
}

function setTipPrice_(price) {
  sheet_('tours').getRange(2, 4).setValue(parseFloat(price) || 0);
  return { price: price };
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

const RECEIPT_COLS_ = { description: 1, observation: 2, paymentMethod: 3, currency: 4, amount: 5 };

function addReceipt_(description, observation, paymentMethod, currency, amount) {
  const sh = sheet_('receipts');
  const rows = sh.getDataRange().getValues();
  const row = rows.length + 1; // first fully blank row after the existing data
  sh.getRange(row, 1, 1, 5).setValues([[description || '', observation || '', paymentMethod || '', currency || '', amount || '']]);
  // Row 2 already carries the dropdown validation for Medio de Pago / Tipo de
  // Moneda — copy it onto the new row so it stays a proper dropdown in the
  // Sheet, not just plain text that happens to match.
  ['paymentMethod', 'currency'].forEach((key) => {
    const col = RECEIPT_COLS_[key];
    const dv = sh.getRange(2, col).getDataValidation();
    if (dv) sh.getRange(row, col).setDataValidation(dv);
  });
  return { row: row, description: description, observation: observation, paymentMethod: paymentMethod, currency: currency, amount: amount };
}

function setReceiptField_(row, field, value) {
  const col = RECEIPT_COLS_[field];
  if (!col) throw new Error('Unknown receipt field: ' + field);
  const sh = sheet_('receipts');
  sh.getRange(row, col).setValue(field === 'amount' ? (parseFloat(value) || 0) : value);
  return { row: row, field: field, value: value };
}
