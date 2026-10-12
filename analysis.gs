/**
 * Jaroi Production Dashboard - Google Apps Script Web App
 * Server side: reads every visible sheet dynamically, finds the real header
 * row, cleans the data and returns it to index.html as JSON.
 */

const SPREADSHEET_ID = '1hzpQkHY_RVae-I0yRCPRkwg50bfJJbkxwysOoBANA5w';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Jaroi Production Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called from the browser with google.script.run.getDashboardData() */
function getDashboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sheets = [];

  ss.getSheets().forEach(function (sh) {
    if (sh.isSheetHidden()) return;                    // skip Read Me / Setup Guide
    const lastR = sh.getLastRow(), lastC = sh.getLastColumn();
    if (lastR < 2 || lastC < 1) return;

    const rng = sh.getRange(1, 1, lastR, lastC);       // dynamic range, no hardcoding
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();

    const hr = findHeaderRow_(disp, lastC);
    const headers = makeUnique_(disp[hr].map(function (h) { return String(h).trim(); }));

    // Data rows = rows below the header whose first column is filled.
    // (Formula-prefilled blank rows have an empty first column and are dropped.)
    const rows = [];
    for (let r = hr + 1; r < vals.length; r++) {
      if (vals[r][0] === '' || vals[r][0] === null) continue;
      if (String(disp[r][0]).trim() === headers[0]) continue;
      rows.push(vals[r].map(function (v, c) { return norm_(v, disp[r][c], tz, headers[c]); }));
    }
    if (!rows.length) return;

    const meta = headers.map(function (h, c) { return inferColumn_(rows.map(function (r) { return r[c]; })); });
    // blank out cells that don't match the column type (stray text in a numeric column)
    rows.forEach(function (row) {
      meta.forEach(function (m, c) {
        if (m.type === 'number' && typeof row[c] !== 'number') row[c] = null;
      });
    });

    sheets.push({
      name: sh.getName(),
      headers: headers,
      types: meta.map(function (m) { return m.type; }),
      cat: meta.map(function (m) { return m.cat; }),
      rows: rows
    });
  });

  return { sheets: sheets, timezone: tz, generatedAt: new Date().toISOString() };
}

/* ---------------- helpers ---------------- */

// Header = first row (within the top 10) where at least half of the columns hold text.
function findHeaderRow_(disp, ncols) {
  const need = Math.max(3, Math.ceil(ncols * 0.5));
  for (let r = 0; r < Math.min(10, disp.length); r++) {
    let n = 0;
    for (let c = 0; c < ncols; c++) if (String(disp[r][c]).trim() !== '') n++;
    if (n >= need) return r;
  }
  return 0;
}

function makeUnique_(arr) {
  const seen = {};
  return arr.map(function (h, i) {
    if (!h) h = 'Column ' + (i + 1);
    if (seen[h]) { seen[h]++; return h + ' (' + seen[h] + ')'; }
    seen[h] = 1;
    return h;
  });
}

// Normalise one cell: real dates -> 'yyyy-MM-dd', time-of-day -> 'HH:mm',
// minutes that were formatted as dates (1900 serials) -> numbers, numbers -> rounded.
function norm_(v, d, tz, header) {
  if (v === '' || v === null) return null;
  if (v instanceof Date) {
    const y = +Utilities.formatDate(v, tz, 'yyyy');
    if (y >= 1950) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const isDuration = /\(min\)/i.test(String(header));   // minutes stored in a date/time-formatted cell
    const t = isDuration ? null : /(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i.exec(String(d));
    if (t) {
      let h = +t[1];
      if (t[3]) { const pm = t[3].toUpperCase() === 'PM'; if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
      return (h < 10 ? '0' : '') + h + ':' + t[2];
    }
    const p = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss').split(/[- :]/).map(Number);
    let s = (Date.UTC(p[0], p[1] - 1, p[2], p[3], p[4], p[5]) - Date.UTC(1899, 11, 30)) / 864e5;
    s = Math.round(s * 100) / 100;
    return Math.abs(s - Math.round(s)) < 0.02 ? Math.round(s) : s;
  }
  if (typeof v === 'number') return Math.round(v * 10000) / 10000;
  if (typeof v === 'boolean') return v ? 'Yes' : 'No';
  const s = String(v).trim();
  return s === '' ? null : s;
}

function inferColumn_(col) {
  const vals = col.filter(function (v) { return v !== null; });
  const n = vals.length;
  if (!n) return { type: 'empty', cat: false };
  const share = function (fn) { return vals.filter(fn).length / n; };
  if (share(function (v) { return typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v); }) >= 0.9) return { type: 'date', cat: false };
  if (share(function (v) { return typeof v === 'string' && /^\d{1,2}:\d{2}$/.test(v); }) >= 0.9) return { type: 'time', cat: false };
  if (share(function (v) { return typeof v === 'number'; }) >= 0.9) return { type: 'number', cat: false };
  const uniq = {};
  vals.forEach(function (v) { uniq[v] = 1; });
  const u = Object.keys(uniq).length;
  return { type: 'string', cat: u >= 2 && u <= 25 };
}
