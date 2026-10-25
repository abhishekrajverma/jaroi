/**
 * TASK MANAGER - AUTOMATION SCRIPT
 *  1. Pick "Assigned To"            -> Assign Date fills in.
 *  2. Set "Send Mail?" to Send Mail  -> the assignee gets an email.
 *  3. Assignee sets Status on their own tab -> Tasks sheet updates.
 *     In Progress -> Start Date fills in. Done -> Completed On fills in (Delivery shows On time / Late by N days).
 *  4. "Update Log" sheet = every update made from the person tabs (full history of notes).
 *  5. "Task Log" sheet    = every change made to the Tasks table (who, what, old value, new value).
 *
 * Setup: Extensions > Apps Script > delete old code > paste this > Save.
 * Then choose  setup  in the function dropdown > Run > approve permissions (one time).
 * (If you already ran setup before, just paste and Save. No need to run it again.)
 */

const CFG = {
  TIMEZONE: 'Asia/Kolkata',          // used for the time stamps. Change if you are in another time zone.
  TASKS_SHEET: 'Tasks',
  LISTS_SHEET: 'Lists',
  LOG_SHEET: 'Update Log',           // created automatically
  TASKLOG_SHEET: 'Task Log',         // created automatically
  FIRST_ROW: 5,                      // first task row on the Tasks sheet
  HEADER_ROW: 4,
  NAMES_RANGE: 'E5:E30',             // Assigned To names on Lists sheet
  EMAILS_RANGE: 'H5:H30',            // matching emails on Lists sheet
  // Tasks sheet columns (A=1, B=2, ...)
  T: { ID: 1, TASK: 2, DESC: 3, PROJECT: 4, CATEGORY: 5, OWNER: 6, ASSIGNED: 7, PRIORITY: 8, STATUS: 9,
       START: 10, DUE: 11, COMPLETED: 12, EST: 14, ACTUAL: 15, PCT: 16, LINK: 23, UPDATED: 24, COMMENTS: 25,
       SEND: 26, MAILSTATUS: 27 },
  // Columns whose changes are written to the Task Log (typed-in columns only)
  TRACK: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 20, 21, 22, 23, 25, 26],
  // Person tabs
  P: { FIRST_ROW: 7, IN_FIRST: 19, IN_LAST: 22, HELPER: 23 },
};

/** Run once: installs the edit trigger. */
function setup() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onTaskEdit');
  if (!exists) ScriptApp.newTrigger('onTaskEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getUi().alert('Done. Automation is on.');
}

/** Fires on every manual edit and routes it. */
function onTaskEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() === CFG.TASKS_SHEET) return tasksEdit_(sh, e);
  if (sh.getRange(6, CFG.P.IN_FIRST).getValue() === 'Update Status') return personEdit_(sh, e);
}

/* ---------- Admin edits on the Tasks sheet ---------- */
function tasksEdit_(sh, e) {
  const T = CFG.T, out = [];
  const by = (e.user && e.user.getEmail && e.user.getEmail()) || 'Admin';
  const r1 = e.range.getRow(), r2 = e.range.getLastRow(), c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  const hit = col => col >= c1 && col <= c2;
  const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;

  // 1) log what the admin typed (up to 60 cells at once)
  if (e.range.getNumRows() * e.range.getNumColumns() <= 60) {
    for (let r = Math.max(r1, CFG.FIRST_ROW); r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (CFG.TRACK.indexOf(c) === -1) continue;
        const nv = sh.getRange(r, c).getDisplayValue();
        const ov = single ? fmtOld_(c, e.oldValue) : '';
        if (single && ov === nv) continue;
        out.push(logRow_(sh, r, c, ov, nv, by));
      }
    }
  }

  // 2) automation
  for (let r = r1; r <= r2; r++) {
    if (r < CFG.FIRST_ROW) continue;
    if (hit(T.OWNER) && !hit(T.SEND)) {                       // Assigned To changed
      const owner = String(sh.getRange(r, T.OWNER).getValue()).trim();
      setAndLog_(sh, r, T.ASSIGNED, owner ? todayDate_() : '', 'Auto (script)', out);
      sh.getRange(r, T.SEND).clearContent();                  // reset mail for the new assignee
      sh.getRange(r, T.MAILSTATUS).clearContent();
    }
    if (hit(T.STATUS)) {                                      // status typed by admin
      const st = String(sh.getRange(r, T.STATUS).getValue()).trim();
      if (st) applyStatusDates_(sh, r, st, out, 'Auto (script)');
    }
    if (hit(T.SEND) && String(sh.getRange(r, T.SEND).getValue()).trim() === 'Send Mail') {
      sendRow_(sh, r);
      out.push(logRow_(sh, r, T.MAILSTATUS, '', sh.getRange(r, T.MAILSTATUS).getDisplayValue(), 'Auto (script)', 'Email'));
    }
  }
  writeLog_(CFG.TASKLOG_SHEET, out);
}

/** Start Date on In Progress, Completed On on Done. */
function applyStatusDates_(tasks, tr, status, out, by) {
  const T = CFG.T, today = todayDate_();
  if (status.indexOf('In Progress') > -1) {
    if (tasks.getRange(tr, T.START).getValue() === '') setAndLog_(tasks, tr, T.START, today, by, out);
    setAndLog_(tasks, tr, T.COMPLETED, '', by, out);
  } else if (status.indexOf('Done') > -1) {
    setAndLog_(tasks, tr, T.COMPLETED, today, by, out);
    if (tasks.getRange(tr, T.PCT).getValue() === '') setAndLog_(tasks, tr, T.PCT, 1, by, out);
  } else {
    setAndLog_(tasks, tr, T.COMPLETED, '', by, out);
  }
}

function sendRow_(sh, row) {
  const T = CFG.T, status = sh.getRange(row, T.MAILSTATUS);
  const owner = String(sh.getRange(row, T.OWNER).getValue()).trim();
  const task  = String(sh.getRange(row, T.TASK).getValue()).trim();
  if (!task)  { status.setValue('Not sent: enter the task first'); return; }
  if (!owner) { status.setValue('Not sent: select Assigned To first'); return; }
  const email = lookupEmail_(owner);
  if (!email) { status.setValue('Not sent: no email set for ' + owner); return; }

  const g = c => sh.getRange(row, c).getDisplayValue() || '-';
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:540px">' +
    '<h2 style="color:#1F2A44;margin:0 0 8px">New task assigned to you</h2>' +
    '<p style="font-size:16px;margin:0 0 12px"><b>' + esc_(task) + '</b></p>' +
    '<table style="border-collapse:collapse;font-size:14px">' +
    row_('Task ID', g(T.ID)) + row_('Project', g(T.PROJECT)) + row_('Category', g(T.CATEGORY)) +
    row_('Priority', g(T.PRIORITY)) + row_('Assigned on', g(T.ASSIGNED)) + row_('Due date', g(T.DUE)) +
    row_('Est. hours', g(T.EST)) + row_('Details', g(T.DESC)) + row_('Reference', g(T.LINK)) + '</table>' +
    '<p><a href="' + SpreadsheetApp.getActive().getUrl() + '">Open the task manager</a></p></div>';
  try {
    MailApp.sendEmail({ to: email, subject: 'New task assigned: ' + task, htmlBody: html });
    status.setValue('Sent to ' + owner + ' | ' + Utilities.formatDate(new Date(), CFG.TIMEZONE, 'dd MMM yyyy HH:mm'));
  } catch (err) { status.setValue('Failed: ' + err.message); }
}

/* ---------- Assignee updates on their own tab -> Tasks sheet ---------- */
function personEdit_(sh, e) {
  const P = CFG.P, T = CFG.T;
  if (e.range.getLastColumn() < P.IN_FIRST || e.range.getColumn() > P.IN_LAST) return;
  const tasks = SpreadsheetApp.getActive().getSheetByName(CFG.TASKS_SHEET);
  const now = new Date();
  const updateRows = [], taskLog = [];
  for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
    if (r < P.FIRST_ROW) continue;
    const idx = sh.getRange(r, P.HELPER).getValue();
    if (typeof idx !== 'number') continue;                    // empty row
    const tr = CFG.FIRST_ROW + idx - 1;                       // matching row on Tasks sheet
    const inputs = sh.getRange(r, P.IN_FIRST, 1, 4);
    const [status, hours, pct, note] = inputs.getValues()[0];
    if (status === '' && hours === '' && pct === '' && note === '') continue;

    const owner = String(tasks.getRange(tr, T.OWNER).getValue());
    const by = owner + ' (own tab)';
    const oldStatus = String(tasks.getRange(tr, T.STATUS).getValue());
    if (status !== '') {
      setAndLog_(tasks, tr, T.STATUS, status, by, taskLog);
      applyStatusDates_(tasks, tr, String(status), taskLog, by);
    }
    if (hours !== '') setAndLog_(tasks, tr, T.ACTUAL, hours, by, taskLog);
    if (pct !== '')   setAndLog_(tasks, tr, T.PCT, pct, by, taskLog);
    if (note !== '')  setAndLog_(tasks, tr, T.COMMENTS,
      Utilities.formatDate(now, CFG.TIMEZONE, 'dd MMM HH:mm') + ' - ' + note, by, taskLog);
    tasks.getRange(tr, T.UPDATED).setValue(todayDate_());

    const change = status === '' ? '-' : (oldStatus && oldStatus !== String(status) ? oldStatus + ' → ' + status : String(status));
    updateRows.push([
      Utilities.formatDate(now, CFG.TIMEZONE, 'dd MMM yyyy HH:mm'),
      tasks.getRange(tr, T.ID).getDisplayValue(), tasks.getRange(tr, T.TASK).getValue(), owner,
      change, hours, pct, note
    ]);
    inputs.clearContent();                                    // ready for the next update
  }
  writeLog_(CFG.LOG_SHEET, updateRows);
  writeLog_(CFG.TASKLOG_SHEET, taskLog);
}

/* ---------- logging ---------- */
const LOGS = {
  'Update Log': { headers: ['Date & Time', 'Task ID', 'Task', 'Assigned To', 'Status Change', 'Actual Hours', '% Complete', 'Update Note'],
                  widths: [150, 80, 260, 130, 240, 100, 100, 420], wrap: [8], pct: 7, text: [1, 8], tab: '#F59E0B' },
  'Task Log':   { headers: ['Date & Time', 'Task ID', 'Task', 'Field', 'Old Value', 'New Value', 'Changed By'],
                  widths: [150, 80, 260, 150, 220, 300, 200], wrap: [5, 6], pct: 0, text: [1, 5, 6, 7], tab: '#06B6D4' },
};

/** Turns the raw 'old value' Google gives us into readable text (dates as 25 Oct 2026, % as 40%). */
function fmtOld_(col, ov) {
  if (ov === undefined || ov === null || ov === '') return '';
  const T = CFG.T, n = Number(ov);
  if ([T.ASSIGNED, T.START, T.DUE, T.COMPLETED, T.UPDATED].indexOf(col) > -1 && !isNaN(n))
    return Utilities.formatDate(new Date(Math.round((n - 25569) * 86400000)), 'UTC', 'dd MMM yyyy');
  if (col === T.PCT && !isNaN(n)) return Math.round(n * 100) + '%';
  return String(ov);
}

/** One Task Log row for a change in a Tasks-sheet cell. */
function logRow_(tasks, tr, col, oldV, newV, by, fieldName) {
  return [
    Utilities.formatDate(new Date(), CFG.TIMEZONE, 'dd MMM yyyy HH:mm'),
    tasks.getRange(tr, CFG.T.ID).getDisplayValue() || '-',
    tasks.getRange(tr, CFG.T.TASK).getValue(),
    fieldName || String(tasks.getRange(CFG.HEADER_ROW, col).getValue()),
    oldV, newV, by
  ];
}

/** Writes a value to the Tasks sheet and records the change if it differs. */
function setAndLog_(tasks, tr, col, value, by, out) {
  const cell = tasks.getRange(tr, col), oldV = cell.getDisplayValue();
  if (value === '') cell.clearContent(); else cell.setValue(value);
  const newV = cell.getDisplayValue();
  if (oldV !== newV) out.push(logRow_(tasks, tr, col, oldV, newV, by));
}

function writeLog_(name, rows) {
  if (!rows.length) return;
  const log = logSheet_(name);
  const first = log.getLastRow() + 1;
  LOGS[name].text.forEach(c => log.getRange(first, c, rows.length, 1).setNumberFormat('@'));   // keep dates as typed text
  log.getRange(first, 1, rows.length, rows[0].length).setValues(rows);
}

/** Finds (or creates and styles) a log sheet. */
function logSheet_(name) {
  const ss = SpreadsheetApp.getActive(), L = LOGS[name], n = L.headers.length;
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  sh.getRange(1, 1, 1, n).setValues([L.headers]).setBackground('#0F172A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);
  L.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 1, sh.getMaxRows() - 1, n).setVerticalAlignment('middle');
  if (L.pct) sh.getRange(2, L.pct, sh.getMaxRows() - 1, 1).setNumberFormat('0%');
  L.wrap.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setWrap(true));
  sh.getRange(1, 1, sh.getMaxRows(), n).createFilter();
  sh.setTabColor(L.tab);
  return sh;
}

/* ---------- helpers ---------- */
/** Today's date at midnight in the spreadsheet's time zone (so date maths is exact). */
function todayDate_() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const s = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return Utilities.parseDate(s, tz, 'yyyy-MM-dd');
}
function lookupEmail_(owner) {
  const ls = SpreadsheetApp.getActive().getSheetByName(CFG.LISTS_SHEET);
  const names = ls.getRange(CFG.NAMES_RANGE).getValues();
  const emails = ls.getRange(CFG.EMAILS_RANGE).getValues();
  for (let i = 0; i < names.length; i++) if (String(names[i][0]).trim() === owner) return String(emails[i][0]).trim();
  return '';
}
function row_(k, v) {
  return '<tr><td style="padding:4px 12px 4px 0;color:#6B7280;vertical-align:top">' + k + '</td><td style="padding:4px 0">' + esc_(v) + '</td></tr>';
}
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
