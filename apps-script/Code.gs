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
  travellerData: 'Traveller Data',
  recs: 'RECOMMENDATIONS',
  receipts: 'RECEIPTS',
  sunat: 'SUNAT_TRIGGER',
  boardingpass: 'BOARDINGPASS_TRIGGER',
  ticketcheck: 'TICKETCHECK_TRIGGER',
  flightschedule: 'FLIGHTSCHEDULE_TRIGGER',
};

// The "TripADeal" Drive folder (Daily Itineraries, Tour Leader Reports,
// etc.) — this script runs as the folder owner, so no extra sharing or
// credentials are needed to read it.
const DOCS_FOLDER_ID = '18quQBR8nf_eSVYKCA0ThIJXZoVJZYTbQ';
// Run this once from the Apps Script editor (select it in the function
// dropdown → Run) to give Yenrri edit access to everything in that folder,
// so tapping a document in the app actually lets him edit it rather than
// just view it.
function shareDocsFolderWithYenrri() {
  const folder = DriveApp.getFolderById(DOCS_FOLDER_ID);
  folder.addEditor('yenrrichacon@gmail.com');
  Logger.log('Shared "%s" with yenrrichacon@gmail.com as an editor.', folder.getName());
}

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

// For timestamps that need minute/second precision (SUNAT watcher
// heartbeat/status) rather than cellStr_'s date-only display format.
// Returns epoch milliseconds, or null if the cell has nothing usable.
function epochMs_(v) {
  if (v instanceof Date) return v.getTime();
  if (!v) return null;
  const t = Date.parse(String(v));
  return isNaN(t) ? null : t;
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
    sunat: getSunatStatus_(),
    boardingPass: getBoardingPassStatus_(),
    ticketCheck: getTicketCheckStatus_(),
    flightSchedule: getFlightScheduleStatus_(),
    documents: getDocuments_(),
    wise: getWiseInfo_(),
  };
}

// ─── CANCELLED TRAVELLERS ──────────────────────────────────────────────
// A traveller who dropped out before the tour (but was never removed from
// the sheet's row layout — every tab is keyed by passenger ID, and
// reshuffling rows would break that) gets marked here instead. Column N
// of CHECK-IN STATUS (the master passenger tab) holds this — every other
// tab that's keyed by passenger ID looks it up from here rather than
// duplicating a "Cancelled" column on each of its own sheets.
const CANCELLED_COL_ = 14; // column N, 1-based
function ensureCancelledHeader_() {
  const sh = sheet_('checkin');
  const header = sh.getRange(1, CANCELLED_COL_).getValue();
  if (!header) sh.getRange(1, CANCELLED_COL_).setValue('Cancelled');
}
function getCancelledIds_() {
  const rows = sheet_('checkin').getDataRange().getValues();
  const ids = new Set();
  for (let r = 1; r < rows.length; r++) {
    if (bool_(rows[r][CANCELLED_COL_ - 1])) ids.add(String(rows[r][0]));
  }
  return ids;
}
function toggleCancelled_(id, value) {
  ensureCancelledHeader_();
  const sh = sheet_('checkin');
  const row = findRowById_(sh, id, 1);
  sh.getRange(row, CANCELLED_COL_).setValue(value ? 'TRUE' : 'FALSE');
  return { id: id, value: value };
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
      cancelled: bool_(row[CANCELLED_COL_ - 1]),
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

  const cancelledIds = getCancelledIds_();
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
      cancelled: cancelledIds.has(String(row[0])),
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
  const cancelledIds = getCancelledIds_();
  const out = [];
  for (let r = 3; r < rows.length; r++) {
    const row = rows[r];
    if (row[1] === '' && row[2] === '') continue;
    out.push({
      id: String(row[0]), name: fullName_(row[1], row[2]), passport: row[3] || '',
      checks: defs.map((d) => ({ col: d.col, checked: bool_(row[d.col]) })),
      cancelled: cancelledIds.has(String(row[0])),
    });
  }
  return { ticketDefs: defs, rows: out };
}

function getInfo_() {
  const rows = sheet_('info').getDataRange().getValues();
  const cancelledIds = getCancelledIds_();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (row[1] === '' && row[2] === '') continue;
    out.push({
      id: String(row[0]), name: fullName_(row[1], row[2]),
      birthday: isoDate_(row[3]), age: row[4] || '', diet: row[5] || '',
      cancelled: cancelledIds.has(String(row[0])),
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

// ─── DOCUMENTS (Drive folder) ──────────────────────────────────────────────
function docType_(mimeType) {
  if (mimeType === MimeType.GOOGLE_DOCS) return 'doc';
  if (mimeType === MimeType.GOOGLE_SHEETS) return 'sheet';
  if (mimeType === MimeType.GOOGLE_SLIDES) return 'slides';
  if (mimeType === MimeType.PDF) return 'pdf';
  return 'file';
}

// "Day 2: Lima city tour" / "Day 13: Puno" sort in proper day order (not
// lexicographic, which would put "Day 10" before "Day 2").
function naturalKey_(name) {
  const m = String(name).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : Number.MAX_SAFE_INTEGER;
}

function listFolderFiles_(folder) {
  const tripSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const out = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getId() === tripSheetId) continue; // the trip Sheet lives in this folder too — already in the app
    if (/^image\//.test(f.getMimeType())) continue; // e.g. the Wise QR code — not a document to browse/edit
    out.push({
      id: f.getId(),
      name: f.getName(),
      url: f.getUrl(),
      type: docType_(f.getMimeType()),
      updated: f.getLastUpdated().getTime(),
    });
  }
  out.sort((a, b) => naturalKey_(a.name) - naturalKey_(b.name) || a.name.localeCompare(b.name));
  return out;
}

// Curation for the loose (non-subfolder) files directly inside the
// TripADeal folder: these two always come first, in this exact order,
// ahead of every subfolder; this one is hidden from the app entirely.
// Matched by Drive file ID (permanent) rather than name — names get
// renamed/retyped (this already happened once: "A 2 Briefing" -> "Briefing"),
// but the underlying file ID never changes.
const TOP_DOC_ORDER_ = [
  '1TGaOAj1h3jJU-jxa5s4wjnbfXB9hzbDp9jZ9B6Mhe-s', // Recommendations for Peru
  '196wurKBKoG-i6UG6x--fhxqIBBoKxkcuVHNu_GHH4d4', // Briefing (was "A 2 Briefing")
];
const EXCLUDED_DOC_IDS_ = [
  '1Zilb4HarBOgqC2YBQubscLTH-0jdkfXXtfp8O7opTp4', // Andean Wings Sotupa Eco Lodge Menu Selection
];

// The merged boarding-pass PDFs land here (written directly by
// boarding-pass-runner.js on the owner's Mac, synced up via Google Drive
// Desktop) — a different folder from the TripADeal one above, so it's
// looked up by name rather than a hardcoded ID. The name is set by that
// script, not something Yenrri or Maree would casually rename.
// User confirmed: each pass is set to "Anyone with the link can view" so
// the WhatsApp share button works for passengers without a Google login.
// Only show passes generated recently — otherwise this list just grows
// forever as old flights pile up. Anything older than this is still safe
// in Drive, it just won't clutter the app.
const BOARDINGPASS_MAX_AGE_MS_ = 3 * 24 * 60 * 60 * 1000; // 3 days

function getBoardingPassDocs_() {
  const it = DriveApp.getFoldersByName('LATAM Boarding Passes');
  if (!it.hasNext()) return [];
  const folder = it.next();
  const cutoff = Date.now() - BOARDINGPASS_MAX_AGE_MS_;
  const files = listFolderFiles_(folder).filter((f) => f.updated >= cutoff);
  files.forEach((f) => {
    try {
      DriveApp.getFileById(f.id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // sharing change failed (rare) — file still appears, just may not
      // be openable by the recipient until this is retried.
    }
  });
  files.sort((a, b) => b.updated - a.updated); // newest first
  return files;
}

function getDocuments_() {
  const root = DriveApp.getFolderById(DOCS_FOLDER_ID);
  const sections = [];

  const boardingPassFiles = getBoardingPassDocs_();
  if (boardingPassFiles.length) sections.push({ name: 'Boarding Passes', files: boardingPassFiles });

  const rootFiles = listFolderFiles_(root).filter((f) => EXCLUDED_DOC_IDS_.indexOf(f.id) === -1);

  const topFiles = TOP_DOC_ORDER_.map((id) => rootFiles.find((f) => f.id === id)).filter(Boolean);
  if (topFiles.length) sections.push({ name: 'Key Documents', files: topFiles });

  const subfolderSections = [];
  const subfolders = root.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    subfolderSections.push({ name: sub.getName(), files: listFolderFiles_(sub) });
  }
  subfolderSections.sort((a, b) => a.name.localeCompare(b.name));
  sections.push.apply(sections, subfolderSections);

  const otherRootFiles = rootFiles.filter((f) => TOP_DOC_ORDER_.indexOf(f.id) === -1);
  if (otherRootFiles.length) sections.push({ name: 'Other documents', files: otherRootFiles });

  return { folderName: root.getName(), sections: sections };
}

// ─── WISE PAYMENT DETAILS ──────────────────────────────────────────────────
// Deliberately NOT hardcoded here — this file is public (GitHub). The actual
// name/BSB/account number live only in a "WISE" tab in the private Sheet;
// this just auto-creates that tab (with blank values) the first time it's
// needed so there's somewhere to type them in.
function getWiseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('WISE');
  if (!sh) {
    sh = ss.insertSheet('WISE');
    sh.getRange(1, 1, 4, 2).setValues([
      ['Field', 'Value'],
      ['Name', ''],
      ['BSB Code', ''],
      ['Account Number', ''],
    ]);
  }
  return sh;
}

// The QR image itself also never touches git — it's found by name in the
// same TripADeal Drive folder the Docs tab already reads, and its bytes are
// embedded directly as a data URI in the JSON response (more reliable than
// linking to a Drive URL, which needs specific sharing/CORS to render inline).
function findWiseQrImage_() {
  try {
    const folder = DriveApp.getFolderById(DOCS_FOLDER_ID);
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (/^image\//.test(f.getMimeType()) && /wise/i.test(f.getName())) {
        const blob = f.getBlob();
        return 'data:' + f.getMimeType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
    }
  } catch (e) {
    // no access / folder missing — just show the text fields without a QR
  }
  return '';
}

function getWiseInfo_() {
  const sh = getWiseSheet_();
  const rows = sh.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < rows.length; r++) map[String(rows[r][0]).trim()] = String(rows[r][1] || '').trim();
  return {
    name: map['Name'] || '',
    bsb: map['BSB Code'] || '',
    account: map['Account Number'] || '',
    qrDataUri: findWiseQrImage_(),
  };
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

// ─── SUNAT submission trigger ─────────────────────────────────────────────
// This tab is a simple mailbox between the app and a watcher script running
// on the owner's own Mac (never anything in this Apps Script or the app
// itself talks to SUNAT directly, and no SUNAT credentials ever pass through
// here). Row 2 holds the current state:
//   A: Status  ('idle' | 'requested' | 'running' | 'done' | 'error')
//   B: RequestedAt   C: StartedAt   D: FinishedAt
//   E: Message (result summary or error text)
//   F: WatcherHeartbeat (written by the watcher on every poll, so the app
//      can tell whether anything is actually listening)
const SUNAT_HEADERS_ = ['Status', 'RequestedAt', 'StartedAt', 'FinishedAt', 'Message', 'WatcherHeartbeat'];

function sunatSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.sunat);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.sunat);
    sh.getRange(1, 1, 1, SUNAT_HEADERS_.length).setValues([SUNAT_HEADERS_]);
    sh.getRange(2, 1).setValue('idle');
  }
  return sh;
}

function getSunatStatus_() {
  const sh = sunatSheet_();
  const row = sh.getRange(2, 1, 1, SUNAT_HEADERS_.length).getValues()[0];
  return {
    status: String(row[0] || 'idle').trim() || 'idle',
    requestedAt: epochMs_(row[1]),
    startedAt: epochMs_(row[2]),
    finishedAt: epochMs_(row[3]),
    message: String(row[4] || ''),
    heartbeat: epochMs_(row[5]),
  };
}

function requestSunatRun_() {
  const sh = sunatSheet_();
  // B2=RequestedAt, C2=StartedAt, D2=FinishedAt, E2=Message
  sh.getRange(2, 1).setValue('requested');
  sh.getRange(2, 2).setValue(new Date());
  sh.getRange(2, 3, 1, 3).setValue(''); // clear StartedAt/FinishedAt/Message from any previous run
  return getSunatStatus_();
}

// ─── Boarding pass generation trigger ─────────────────────────────────────
// Same mailbox pattern as SUNAT above, but for boarding-pass-runner.js — a
// watcher on the owner's Mac polls this tab, runs `./deploy.sh && node
// boarding-pass-runner.js`, and writes the result back here. No credentials
// pass through this sheet or this script; the watcher already has its own
// Google OAuth token on disk.
// Column G (Progress) is written mid-run by the watcher as it streams the
// runner's output — "current/total" passengers processed so far — so the
// app can show a progress bar/ETA instead of just "Running…" for however
// long a full run takes.
const BOARDINGPASS_HEADERS_ = ['Status', 'RequestedAt', 'StartedAt', 'FinishedAt', 'Message', 'WatcherHeartbeat', 'Progress'];

function boardingPassSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.boardingpass);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.boardingpass);
    sh.getRange(1, 1, 1, BOARDINGPASS_HEADERS_.length).setValues([BOARDINGPASS_HEADERS_]);
    sh.getRange(2, 1).setValue('idle');
  } else if (sh.getLastColumn() < BOARDINGPASS_HEADERS_.length) {
    // Tab already existed from before the Progress column was added —
    // "create if missing" alone never backfills new columns onto an
    // existing sheet, so do that explicitly here.
    sh.getRange(1, 1, 1, BOARDINGPASS_HEADERS_.length).setValues([BOARDINGPASS_HEADERS_]);
  }
  return sh;
}

function getBoardingPassStatus_() {
  const sh = boardingPassSheet_();
  const row = sh.getRange(2, 1, 1, BOARDINGPASS_HEADERS_.length).getValues()[0];
  return {
    status: String(row[0] || 'idle').trim() || 'idle',
    requestedAt: epochMs_(row[1]),
    startedAt: epochMs_(row[2]),
    finishedAt: epochMs_(row[3]),
    message: String(row[4] || ''),
    heartbeat: epochMs_(row[5]),
    progress: String(row[6] || ''),
  };
}

function requestBoardingPassRun_() {
  const sh = boardingPassSheet_();
  sh.getRange(2, 1).setValue('requested');
  sh.getRange(2, 2).setValue(new Date());
  sh.getRange(2, 3, 1, 3).setValue(''); // clear StartedAt/FinishedAt/Message from any previous run
  sh.getRange(2, 7).setValue(''); // clear Progress from any previous run
  return getBoardingPassStatus_();
}

// ─── Ticket check trigger ──────────────────────────────────────────────
// Same mailbox pattern again, for ticket-checker.js — it cross-checks the
// TICKETS sheet against whatever PDFs are sitting in the Drive ticket-inbox
// folder and writes the results straight back to that sheet. No files or
// credentials pass through here.
const TICKETCHECK_HEADERS_ = ['Status', 'RequestedAt', 'StartedAt', 'FinishedAt', 'Message', 'WatcherHeartbeat'];

function ticketCheckSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.ticketcheck);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.ticketcheck);
    sh.getRange(1, 1, 1, TICKETCHECK_HEADERS_.length).setValues([TICKETCHECK_HEADERS_]);
    sh.getRange(2, 1).setValue('idle');
  }
  return sh;
}

function getTicketCheckStatus_() {
  const sh = ticketCheckSheet_();
  const row = sh.getRange(2, 1, 1, TICKETCHECK_HEADERS_.length).getValues()[0];
  return {
    status: String(row[0] || 'idle').trim() || 'idle',
    requestedAt: epochMs_(row[1]),
    startedAt: epochMs_(row[2]),
    finishedAt: epochMs_(row[3]),
    message: String(row[4] || ''),
    heartbeat: epochMs_(row[5]),
  };
}

function requestTicketCheckRun_() {
  const sh = ticketCheckSheet_();
  sh.getRange(2, 1).setValue('requested');
  sh.getRange(2, 2).setValue(new Date());
  sh.getRange(2, 3, 1, 3).setValue(''); // clear StartedAt/FinishedAt/Message from any previous run
  return getTicketCheckStatus_();
}

// ─── Flight monitor scheduling trigger ────────────────────────────────────
// Same mailbox pattern again, for schedule-jobs.js — a one-shot setup step
// (not a continuously-running monitor itself) that reads the current
// tour's flight data and (re)installs the two actual background monitors
// (flight-monitor.js, flight-status-monitor.js) with the right future
// start times, plus a calendar reminder on the owner's own calendar. Safe
// to run repeatedly — it recomputes and reinstalls cleanly each time.
// Column G (DataSnapshot) holds a fingerprint of the Traveller Data /
// FLIGHT STATUS cells schedule-jobs.js actually reads, captured at the
// moment a run is requested — so the app can tell when that data has
// since changed and a re-run is actually needed (see
// flightScheduleDataSnapshot_() below).
const FLIGHTSCHEDULE_HEADERS_ = ['Status', 'RequestedAt', 'StartedAt', 'FinishedAt', 'Message', 'WatcherHeartbeat', 'DataSnapshot'];

function flightScheduleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.flightschedule);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.flightschedule);
    sh.getRange(1, 1, 1, FLIGHTSCHEDULE_HEADERS_.length).setValues([FLIGHTSCHEDULE_HEADERS_]);
    sh.getRange(2, 1).setValue('idle');
  } else if (sh.getLastColumn() < FLIGHTSCHEDULE_HEADERS_.length) {
    // Tab already existed from before DataSnapshot was added — "create if
    // missing" alone never backfills new columns onto an existing sheet.
    sh.getRange(1, 1, 1, FLIGHTSCHEDULE_HEADERS_.length).setValues([FLIGHTSCHEDULE_HEADERS_]);
  }
  return sh;
}

// A fingerprint of exactly the cells schedule-jobs.js reads to compute
// monitor start times: Traveller Data's Lima-arrival column (including
// its header, which itself encodes the arrival date) and FLIGHT STATUS's
// date + scheduled-time columns. Any change to those cells changes this
// string, regardless of anything else happening elsewhere in the sheet.
function flightScheduleDataSnapshot_() {
  // Only the exact substrings schedule-jobs.js's own regexes extract are
  // included — not the raw cell text. FLIGHT STATUS column D in
  // particular carries live-updating status text ("EN ROUTE — arriving in
  // 12 minutes"), most of which schedule-jobs.js ignores entirely (it
  // only reads a *leading* HH:MM). Fingerprinting the raw cell would flag
  // "changed" on every such status tick even though nothing that actually
  // feeds the schedule computation moved.
  const traveller = sheet_('travellerData');
  const travellerLastRow = traveller.getLastRow();
  const limaArrivals = travellerLastRow > 0
    ? traveller.getRange(1, 8, travellerLastRow, 1).getValues().map((r) => {
        const cell = String(r[0] || '');
        const m = cell.match(/(\d{1,2}):(\d{2})/);
        return m ? m[0] : cell; // row 1 is the header (carries the date, no time) — keep as-is
      }).join('|')
    : '';

  const flights = sheet_('flights');
  const flightsLastRow = flights.getLastRow();
  const flightRows = flightsLastRow > 1
    ? flights.getRange(2, 1, flightsLastRow - 1, 4).getValues().map((r) => {
        const dateStr = String(r[0] || '');
        const m = String(r[3] || '').match(/^(\d{1,2}):(\d{2})/);
        return dateStr + '~' + (m ? m[0] : '');
      }).join('|')
    : '';

  return limaArrivals + '##' + flightRows;
}

function getFlightScheduleStatus_() {
  const sh = flightScheduleSheet_();
  const row = sh.getRange(2, 1, 1, FLIGHTSCHEDULE_HEADERS_.length).getValues()[0];
  const storedSnapshot = String(row[6] || '');
  let dataChangedSinceLastRun = false;
  if (storedSnapshot) {
    try {
      dataChangedSinceLastRun = flightScheduleDataSnapshot_() !== storedSnapshot;
    } catch (e) {
      // Can't read the source sheets right now — don't false-alarm.
      dataChangedSinceLastRun = false;
    }
  }
  return {
    status: String(row[0] || 'idle').trim() || 'idle',
    requestedAt: epochMs_(row[1]),
    startedAt: epochMs_(row[2]),
    finishedAt: epochMs_(row[3]),
    message: String(row[4] || ''),
    heartbeat: epochMs_(row[5]),
    dataChangedSinceLastRun: dataChangedSinceLastRun,
  };
}

function requestFlightScheduleRun_() {
  const sh = flightScheduleSheet_();
  sh.getRange(2, 1).setValue('requested');
  sh.getRange(2, 2).setValue(new Date());
  sh.getRange(2, 3, 1, 3).setValue(''); // clear StartedAt/FinishedAt/Message from any previous run
  try {
    sh.getRange(2, 7).setValue(flightScheduleDataSnapshot_());
  } catch (e) {
    sh.getRange(2, 7).setValue(''); // couldn't read source data — leave unset rather than fail the request
  }
  return getFlightScheduleStatus_();
}

// ─── WRITE ───────────────────────────────────────────────────────────────
function handleWrite_(body) {
  switch (body.action) {
    case 'toggleChecklist': return toggleChecklistItem_(body.row, body.done);
    case 'toggleArrived': return toggleArrived_(body.id, body.value);
    case 'toggleCancelled': return toggleCancelled_(body.id, body.value);
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
    case 'requestSunatRun': return requestSunatRun_();
    case 'requestBoardingPassRun': return requestBoardingPassRun_();
    case 'requestTicketCheckRun': return requestTicketCheckRun_();
    case 'requestFlightScheduleRun': return requestFlightScheduleRun_();
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
