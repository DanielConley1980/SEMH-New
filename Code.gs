// SEMH Assessment Tracker — Google Apps Script backend
// Deploy as: Web App → Execute as: Me → Who has access: Anyone
//
// 1. Open this spreadsheet in Google Sheets
// 2. Extensions → Apps Script → paste this entire file → Save
// 3. Deploy → New deployment → Web app → Execute as: Me → Anyone → Deploy
// 4. Copy the web app URL into the SEMH Tracker "Apps Script URL" field

const SPREADSHEET_ID = '1GVmV099c3MVX3rzNWICyMVMU5GQxdMpn5hnNAgGoRp4';
const STUDENTS_GID   = 1402887730;   // The "Students" sheet tab

// ── GET: JSONP endpoint for fetchStudents ────────────────────────────────
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

// ── POST: endpoint for saveReview ────────────────────────────────────────
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
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet  = sheetByGid(ss, STUDENTS_GID);
  const lastR  = sheet.getLastRow();
  const lastC  = sheet.getLastColumn();
  if (lastR < 1 || lastC < 1) return { students: [] };

  const all     = sheet.getRange(1, 1, lastR, lastC).getValues();
  const headers = all[0].map(h => String(h).trim().toLowerCase());

  // Detect which columns hold name / form / year — try common UK school labels
  const nc = colIndex(headers, ['name','student name','student','full name','pupil','pupil name','first name']);
  const fc = colIndex(headers, ['form','class','form group','tutor group','registration group','group','set']);
  const yc = colIndex(headers, ['year','year group','yr','year grp','y']);

  // Fallback: col A = name, col B = form, col C = year
  const ni = nc >= 0 ? nc : 0;
  const fi = fc >= 0 ? fc : 1;
  const yi = yc >= 0 ? yc : 2;

  // Skip header row only when at least one header was actually recognised
  const start = nc >= 0 ? 1 : 0;

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

// ── Read existing SEMH scores for a student ─────────────────────────────
function getReview(name) {
  if (!name) return { error: 'No name provided' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sheetByGid(ss, STUDENTS_GID);
  const lastR = sheet.getLastRow();
  const lastC = Math.max(sheet.getLastColumn(), 1);

  const headerRow = sheet.getRange(1, 1, 1, lastC).getValues()[0];
  const headers   = headerRow.map(h => String(h).trim().toLowerCase());

  const nc = colIndex(headers, ['name','student name','student','full name','pupil','pupil name','first name']);
  const ni = nc >= 0 ? nc : 0;

  // Find result columns
  const basScoreIdx  = colIndex(headers, ['baseline score', 'baseline']);
  const basDateIdx   = colIndex(headers, ['baseline date']);
  const currScoreIdx = colIndex(headers, ['current score', 'current', 'review score']);
  const revDateIdx   = colIndex(headers, ['review date']);
  const pctIdx       = colIndex(headers, ['% change', 'percent change', 'pct change', 'change']);

  // Find student row
  if (lastR < 2) return { found: false };
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
  const result = { found: true };
  if (basScoreIdx  >= 0) result.baseline      = row[basScoreIdx];
  if (basDateIdx   >= 0) result.baselineDate  = row[basDateIdx]  ? Utilities.formatDate(new Date(row[basDateIdx]),  Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  if (currScoreIdx >= 0) result.current       = row[currScoreIdx];
  if (revDateIdx   >= 0) result.reviewDate    = row[revDateIdx]  ? Utilities.formatDate(new Date(row[revDateIdx]),   Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  if (pctIdx       >= 0) result.percentChange = String(row[pctIdx]).replace('%', '');
  return result;
}

// ── Write SEMH scores back to the student row ────────────────────────────
// Writes to the SAME sheet as the student list.
// Appends result columns (Baseline Score, Baseline Date, Current Score,
// Review Date, % Change) if they don't already exist.
function saveReview(p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sheetByGid(ss, STUDENTS_GID);
  const lastR = sheet.getLastRow();
  const lastC = Math.max(sheet.getLastColumn(), 1);

  const headerRow = sheet.getRange(1, 1, 1, lastC).getValues()[0];
  const headers   = headerRow.map(h => String(h).trim().toLowerCase());

  // Find the name column
  const nc = colIndex(headers, ['name','student name','student','full name','pupil','pupil name','first name']);
  const ni = nc >= 0 ? nc : 0;

  // Result column definitions: [display label, search aliases]
  const RESULT_DEFS = [
    ['Baseline Score', ['baseline score', 'baseline']],
    ['Baseline Date',  ['baseline date']],
    ['Current Score',  ['current score', 'current', 'review score']],
    ['Review Date',    ['review date']],
    ['% Change',       ['% change', 'percent change', 'pct change', 'change']],
  ];

  // Ensure each result column exists, collect their 1-based column numbers
  const resultCols = RESULT_DEFS.map(([label, aliases]) => {
    let idx = colIndex(headers, aliases);
    if (idx < 0) {
      // Append a new column header
      idx = sheet.getLastColumn(); // 0-based index of new col
      sheet.getRange(1, idx + 1).setValue(label);
      headers.push(label.toLowerCase());
    }
    return idx + 1; // 1-based
  });

  // Find the student row (search from row 2 downward)
  let targetRow = -1;
  if (lastR >= 2) {
    const nameData = sheet.getRange(2, ni + 1, lastR - 1, 1).getValues();
    for (let i = 0; i < nameData.length; i++) {
      if (String(nameData[i][0]).trim().toLowerCase() === String(p.name || '').toLowerCase()) {
        targetRow = i + 2; // 1-based row
        break;
      }
    }
  }

  if (targetRow < 0) {
    return { error: 'Student not found in sheet: "' + p.name + '". Check the name matches exactly.' };
  }

  const values = [
    p.baseline    != null ? p.baseline    : '',
    p.baselineDate || '',
    p.current     != null ? p.current     : '',
    p.reviewDate  || '',
    p.percentChange != null ? p.percentChange + '%' : '',
  ];

  for (let i = 0; i < resultCols.length; i++) {
    sheet.getRange(targetRow, resultCols[i]).setValue(values[i]);
  }

  return { ok: true, row: targetRow };
}

// ── Helpers ──────────────────────────────────────────────────────────────
function sheetByGid(ss, gid) {
  for (const s of ss.getSheets()) {
    if (s.getSheetId() === gid) return s;
  }
  return ss.getSheets()[0]; // fallback to first sheet
}

function colIndex(headers, candidates) {
  // Exact match first
  for (const c of candidates) {
    const i = headers.indexOf(c);
    if (i >= 0) return i;
  }
  // Partial/substring match
  for (const c of candidates) {
    const i = headers.findIndex(h => h.includes(c) || c.includes(h));
    if (i >= 0) return i;
  }
  return -1;
}

// ── JSON response helper ─────────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
