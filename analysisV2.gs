/**
 * Jaroi Production Dashboard - Google Apps Script Web App
 * Server side: reads every visible sheet dynamically, finds the real header
 * row, cleans the data and returns it to index.html as JSON.
 * Authentication added; dashboard data is exposed only through the
 * authenticated wrapper, not through a client-callable getDashboardData().
 */

const SPREADSHEET_ID = '1hzpQkHY_RVae-I0yRCPRkwg50bfJJbkxwysOoBANA5w';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Jaroi Production Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============ Authentication ============ */

function loginUser(username, password) {
  username = String(username || '').trim();
  password = String(password || '');

  const props = PropertiesService.getScriptProperties();
  const savedUsername = props.getProperty('DASHBOARD_USERNAME');
  const savedPassword = props.getProperty('DASHBOARD_PASSWORD');

  if (!savedUsername || !savedPassword) {
    throw new Error('Login is not configured. Contact the administrator.');
  }

  if (username !== savedUsername || password !== savedPassword) {
    return { success: false, message: 'Invalid username or password.' };
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('jaroi_session_' + token, username, 21600);

  return { success: true, token: token, username: username };
}

function validateSession_(token) {
  if (!token || typeof token !== 'string') {
    throw new Error('Please log in to access the dashboard.');
  }

  const username = CacheService.getScriptCache().get('jaroi_session_' + token);
  if (!username) {
    throw new Error('Your session has expired. Please log in again.');
  }
  return username;
}

function logoutUser(token) {
  if (token && typeof token === 'string') {
    CacheService.getScriptCache().remove('jaroi_session_' + token);
  }
  return { success: true };
}

/** Browser-callable data endpoint. Always validates the session first. */
function getDashboardDataAuthenticated(token) {
  validateSession_(token);
  return getDashboardData_();
}

/** Private server-only implementation; trailing underscore blocks google.script.run calls. */
function getDashboardData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sheets = [];

  ss.getSheets().forEach(function (sh) {
    if (sh.isSheetHidden()) return;
    const lastR = sh.getLastRow(), lastC = sh.getLastColumn();
    if (lastR < 2 || lastC < 1) return;

    const rng = sh.getRange(1, 1, lastR, lastC);
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();

    const hr = findHeaderRow_(disp, lastC);
    const headers = makeUnique_(disp[hr].map(function (h) { return String(h).trim(); }));

    const rows = [];
    for (let r = hr + 1; r < vals.length; r++) {
      if (vals[r][0] === '' || vals[r][0] === null) continue;
      if (String(disp[r][0]).trim() === headers[0]) continue;
      rows.push(vals[r].map(function (v, c) {
        return norm_(v, disp[r][c], tz, headers[c]);
      }));
    }
    if (!rows.length) return;

    const meta = headers.map(function (h, c) {
      return inferColumn_(rows.map(function (r) { return r[c]; }));
    });
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

function norm_(v, d, tz, header) {
  if (v === '' || v === null) return null;
  if (v instanceof Date) {
    const y = +Utilities.formatDate(v, tz, 'yyyy');
    if (y >= 1950) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const isDuration = /\(min\)/i.test(String(header));
    const t = isDuration ? null : /(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i.exec(String(d));
    if (t) {
      let h = +t[1];
      if (t[3]) {
        const pm = t[3].toUpperCase() === 'PM';
        if (pm && h < 12) h += 12;
        if (!pm && h === 12) h = 0;
      }
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
