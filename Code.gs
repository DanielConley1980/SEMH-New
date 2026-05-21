// SEMH Assessment Tracker — Google Apps Script backend
// Deploy as: Web App → Execute as: Me → Who has access: Anyone
//
// 1. Open this spreadsheet in Google Sheets
// 2. Extensions → Apps Script → paste this entire file → Save
// 3. Deploy → New deployment → Web app → Execute as: Me → Anyone → Deploy
// 4. Copy the web app URL into the SEMH Tracker "Apps Script URL" field

const SPREADSHEET_ID = '1w588riFpf2HVkzDM3UvAnEyBhCn6nKw3tskz8sQLyu4';
const STUDENTS_GID   = 1402887730;   // The sheet tab GID

// Column headers written to the sheet
// Baseline block: "Baseline Date", "Baseline Total", "B01"–"B32"
// Review block:   "Review Date",   "Review Total",   "% Change", "R01"–"R32"
function baselineItemCol(id) { return 'B' + String(id).padStart(2, '0'); }
function reviewItemCol(id)   { return 'R' + String(id).padStart(2, '0'); }

// ── One-time data cleanup ────────────────────────────────────────────────
// Run this ONCE from the Apps Script editor (select cleanupStudentData →
// click Run). It strips "Year " prefixes from the Year column and keeps
// only the first form entry if a cell contains multiple values.
function cleanupStudentData() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sheetByGid(ss, STUDENTS_GID);
  const lastR = sheet.getLastRow();
  const lastC = sheet.getLastColumn();
  if (lastR < 2) { Logger.log('No data rows found.'); return; }

  const all     = sheet.getRange(1, 1, lastR, lastC).getValues();
  const headers = all[0].map(h => String(h).trim().toLowerCase());

  const fi = colIndex(headers, ['form','class','form group','tutor group','registration group','group','set']);
  const yi = colIndex(headers, ['year','year group','yr','year grp']);

  if (fi < 0 && yi < 0) { Logger.log('Could not find Form or Year columns.'); return; }

  let changes = 0;
  for (let r = 1; r < all.length; r++) {
    let changed = false;

    if (yi >= 0) {
      const raw = String(all[r][yi] || '').trim();
      const cleaned = raw.replace(/^year\s*/i, '');
      if (cleaned !== raw) {
        sheet.getRange(r + 1, yi + 1).setValue(cleaned);
        changed = true;
      }
    }

    if (fi >= 0) {
      const raw = String(all[r][fi] || '').trim();
      // Take only the first value when cells contain comma/semicolon/space separated forms
      const first = raw.split(/[,;\/\s]+/)[0].trim();
      if (first !== raw) {
        sheet.getRange(r + 1, fi + 1).setValue(first);
        changed = true;
      }
    }

    if (changed) changes++;
  }
  Logger.log('Done. ' + changes + ' rows updated.');
}

// ── GET: JSONP endpoint (fetchStudents, getReview) ───────────────────────
function doGet(e) {
  const cb     = e.parameter.callback || 'callback';
  const action = e.parameter.action   || '';
  let result;
  try {
    if (action === 'students') {
      result = getStudents();
    } else if (action === 'review') {
      result = getReview(decodeURIComponent(e.parameter.name || ''));
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ── POST: endpoint (saveReview) ──────────────────────────────────────────
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ error: 'Bad JSON: ' + err.message });
  }
  let result;
  try {
    if (payload.action === 'saveReview') {
      result = saveReview(payload);
    } else {
      result = { error: 'Unknown action: ' + payload.action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return respond(result);
}

// ── Return the student list ──────────────────────────────────────────────
function getStudents() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sheetByGid(ss, STUDENTS_GID);
  const lastR = sheet.getLastRow();
  const lastC = sheet.getLastColumn();
  if (lastR < 1 || lastC < 1) return { students: [] };

  const all     = sheet.getRange(1, 1, lastR, lastC).getValues();
  const headers = all[0].map(h => String(h).trim().toLowerCase());

  const ni = colIndexOrFallback(headers, ['name','student name','student','full name','pupil','pupil name','first name'], 0);
  const fi = colIndexOrFallback(headers, ['form','class','form group','tutor group','registration group','group','set'], 1);
  const yi = colIndexOrFallback(headers, ['year','year group','yr','year grp'], 2);

  // Skip header row when at least the name column was recognised
  const start = (colIndex(headers, ['name','student name','student','full name','pupil','pupil name','first name']) >= 0) ? 1 : 0;

  const students = [];
  for (let r = start; r < all.length; r++) {
    const name = String(all[r][ni] || '').trim();
    if (!name) continue;
    const s = { name };
    const form = String(all[r][fi] || '').trim();
    const year = String(all[r][yi] || '').trim();
    if (form) s.form = form;
    if (year) s.year = year;
    students.push(s);
  }
  return { students };
}

// ── Read all SEMH data for a student ────────────────────────────────────
function getReview(name) {
  if (!name) return { error: 'No name provided' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sheetByGid(ss, STUDENTS_GID);
  const lastR = sheet.getLastRow();
  if (lastR < 2) return { found: false };

  const lastC  = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastC).getValues()[0].map(h => String(h).trim());
  const headersL = headers.map(h => h.toLowerCase());

  const ni = colIndexOrFallback(headersL, ['name','student name','student','full name','pupil','pupil name','first name'], 0);

  // Find student row
  const nameData = sheet.getRange(2, ni + 1, lastR - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < nameData.length; i++) {
    if (String(nameData[i][0]).trim().toLowerCase() === name.toLowerCase()) {
      targetRow = i + 2;
      break;
    }
  }
  if (targetRow < 0) return { found: false };

  const row = sheet.getRange(targetRow, 1, 1, lastC).getValues()[0];

  // Helper to read a named column value
  function val(label) {
    const idx = headers.indexOf(label);
    return idx >= 0 ? row[idx] : null;
  }
  function dateVal(label) {
    const v = val(label);
    if (!v) return '';
    try { return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
    catch(e) { return ''; }
  }
  function numVal(label) {
    const v = val(label);
    return (v !== null && v !== '') ? Number(v) : null;
  }

  const result = { found: true };

  result.baselineDate  = dateVal('Baseline Date');
  result.baseline      = numVal('Baseline Total');
  result.reviewDate    = dateVal('Review Date');
  result.current       = numVal('Review Total');
  const pctRaw = val('% Change');
  result.percentChange = pctRaw !== null ? String(pctRaw).replace('%', '') : null;

  // Individual item scores (32 items each)
  result.baselineItems = [];
  result.reviewItems   = [];
  for (let id = 1; id <= 32; id++) {
    result.baselineItems.push(numVal(baselineItemCol(id)));
    result.reviewItems.push(numVal(reviewItemCol(id)));
  }

  return result;
}

// ── Write SEMH scores back to the student's row ──────────────────────────
// payload.mode === 'baseline'  → writes Baseline Date, Baseline Total, B01–B32
// payload.mode === 'review'    → writes Review Date, Review Total, % Change, R01–R32
// Both modes append any missing column headers automatically.
function saveReview(p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sheetByGid(ss, STUDENTS_GID);
  const lastR = sheet.getLastRow();
  if (lastR < 2) return { error: 'Sheet has no data rows' };

  // Read current headers (re-read after any appends)
  function readHeaders() {
    const lc = sheet.getLastColumn();
    return sheet.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h).trim());
  }

  let headers = readHeaders();
  const headersL = headers.map(h => h.toLowerCase());
  const ni = colIndexOrFallback(headersL, ['name','student name','student','full name','pupil','pupil name','first name'], 0);

  // Find student row
  const nameData = sheet.getRange(2, ni + 1, lastR - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < nameData.length; i++) {
    if (String(nameData[i][0]).trim().toLowerCase() === String(p.name || '').toLowerCase()) {
      targetRow = i + 2;
      break;
    }
  }
  if (targetRow < 0) {
    return { error: 'Student not found: "' + p.name + '". Check the name matches exactly.' };
  }

  const mode  = p.mode  || 'review';
  const items = p.items || [];  // array of { id, label, baseline, current }

  // Build ordered list of [columnHeader, value] pairs
  const updates = [];
  if (mode === 'baseline') {
    updates.push(['Baseline Date',  p.baselineDate || '']);
    updates.push(['Baseline Total', p.baseline != null ? Number(p.baseline) : '']);
    for (const it of items) {
      updates.push([baselineItemCol(it.id), it.baseline != null ? Number(it.baseline) : '']);
    }
  } else {
    updates.push(['Review Date',  p.reviewDate  || '']);
    updates.push(['Review Total', p.current     != null ? Number(p.current)      : '']);
    updates.push(['% Change',     p.percentChange != null ? Number(p.percentChange) : '']);
    for (const it of items) {
      updates.push([reviewItemCol(it.id), it.current != null ? Number(it.current) : '']);
    }
  }

  // Ensure every needed column header exists (append if missing)
  for (const [colName] of updates) {
    if (!headers.includes(colName)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(colName);
      headers.push(colName); // keep local copy in sync
    }
  }

  // Refresh headers after any appends
  headers = readHeaders();

  // Find the contiguous column range to update, then do a single setValues call
  const colNums = updates
    .map(([h]) => headers.indexOf(h) + 1)   // 1-based; 0 → not found
    .filter(c => c > 0);
  if (colNums.length === 0) return { ok: true };

  const minCol = Math.min(...colNums);
  const maxCol = Math.max(...colNums);
  const width  = maxCol - minCol + 1;

  // Read existing row values for that span
  const rowBuf = sheet.getRange(targetRow, minCol, 1, width).getValues()[0];

  // Overlay the updates
  for (const [colName, value] of updates) {
    const idx = headers.indexOf(colName);
    if (idx >= 0) rowBuf[idx - (minCol - 1)] = value;
  }

  // Write back in one batch
  sheet.getRange(targetRow, minCol, 1, width).setValues([rowBuf]);

  return { ok: true, row: targetRow };
}

// ── Helpers ──────────────────────────────────────────────────────────────
function sheetByGid(ss, gid) {
  for (const s of ss.getSheets()) {
    if (s.getSheetId() === gid) return s;
  }
  return ss.getSheets()[0];
}

// Returns the first matching column index (0-based), or -1
function colIndex(headers, candidates) {
  for (const c of candidates) {
    const i = headers.indexOf(c);
    if (i >= 0) return i;
  }
  for (const c of candidates) {
    const i = headers.findIndex(h => h.includes(c) || c.includes(h));
    if (i >= 0) return i;
  }
  return -1;
}

// Like colIndex but returns `fallback` instead of -1
function colIndexOrFallback(headers, candidates, fallback) {
  const i = colIndex(headers, candidates);
  return i >= 0 ? i : fallback;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
