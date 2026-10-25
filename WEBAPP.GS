/**
 * TASK MANAGER - WEB APP (server side)
 * Add as a new script file named  WebApp  in the same Apps Script project (next to Code and Dashboard).
 * It reuses the functions from your automation script, so history logs, dates and emails all keep working.
 *
 * Login: people sign in with an ID and password stored in the "Users" sheet (created by setupUsers).
 * Passwords are turned into secure hashes automatically the first time each person signs in.
 *
 * Setup: add this file + an HTML file named  App  (paste App.html), run  setupUsers  once,
 *        then Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 */

const WEB = { USERS_SHEET: 'Users', ROWS: 200, RECUR: 20, BLOCK: 21, TAGS: 22, SESSION_SECS: 21600, MAX_FAILS: 5 };

/* ---------- serve the page ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('App').setTitle('Task Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- time zone ---------- */
/** Makes the spreadsheet use the time zone set in CFG.TIMEZONE (fixes a wrong "today", Overdue and Days Left). */
function ensureTimeZone_() {
  const ss = SpreadsheetApp.getActive();
  if (CFG.TIMEZONE && ss.getSpreadsheetTimeZone() !== CFG.TIMEZONE) ss.setSpreadsheetTimeZone(CFG.TIMEZONE);
}
/** Run from the editor to fix the time zone right now. */
function fixTimeZone() {
  ensureTimeZone_();
  SpreadsheetApp.getUi().alert('Time zone is now ' + SpreadsheetApp.getActive().getSpreadsheetTimeZone() + '.');
}

/* ---------- users and login ---------- */
const ROLES_ = ['Admin', 'Manager', 'Member'];

/** Run once (and again after adding people): creates the Users sheet with an ID and starting password for everyone. */
function setupUsers() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(WEB.USERS_SHEET), rows = [];
  if (!sh) sh = ss.insertSheet(WEB.USERS_SHEET);
  else if (sh.getLastRow() > 1) rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().filter(r => r[0] !== '');
  const pw = () => Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const ids = () => rows.map(r => String(r[1]).toLowerCase());
  if (!rows.some(r => String(r[2]) === 'Admin')) rows.unshift(['Admin', 'admin', 'Admin', pw(), '']);
  const lists = ss.getSheetByName(CFG.LISTS_SHEET);
  const names = lists.getRange(CFG.NAMES_RANGE).getValues(), emails = lists.getRange(CFG.EMAILS_RANGE).getValues();
  names.forEach((n, i) => {
    const name = String(n[0]).trim();
    if (!name || rows.some(r => String(r[0]).trim() === name)) return;
    let id = name.toLowerCase().replace(/[^a-z0-9]/g, '') || 'user', k = id, c = 1;
    while (ids().indexOf(k) > -1) k = id + (++c);
    rows.push([name, k, 'Member', pw(), String(emails[i][0]).trim()]);
  });
  rows.forEach(r => { if (ROLES_.indexOf(String(r[2])) === -1) r[2] = 'Member'; if (r[3] === '') r[3] = pw(); });
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([['Name', 'User ID', 'Role', 'Password', 'Email']])
    .setBackground('#0F172A').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(2, 4, Math.max(rows.length, 1), 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 5).setValues(rows);
  [150, 130, 90, 330, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1); sh.setTabColor('#F59E0B');
  sh.getRange(1, 7).setValue('Name must match the person\'s name in Lists. Roles: Admin = everything, Manager = view only, Member = own tasks. ' +
    'To set or reset a password, just type a new one in the Password column. After a person signs in, it is replaced by a secure hash automatically.')
    .setFontColor('#64748B');
  SpreadsheetApp.getUi().alert('Done. Open the "Users" sheet: each person has an ID and a starting password. Share them privately.');
}

function cleanKey_(k) { return String(k || '').replace(/[^A-Za-z0-9]/g, ''); }
function hashPw_(salt, pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw)
    .map(b => ((b < 0 ? b + 256 : b)).toString(16).padStart(2, '0')).join('');
}
function userRow_(sh, id) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] !== '' && String(v[i][1]).trim().toLowerCase() === id) return i + 2;
  return 0;
}
function checkPw_(sh, row, pw) {
  const cell = sh.getRange(row, 4), stored = String(cell.getValue());
  if (stored.indexOf('h:') === 0) {
    const p = stored.split(':');
    return p.length === 3 && hashPw_(p[1], pw) === p[2];
  }
  if (stored !== '' && stored === pw) {                         // first sign-in: replace the plain password by a hash
    const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    cell.setNumberFormat('@').setValue('h:' + salt + ':' + hashPw_(salt, pw));
    return true;
  }
  return false;
}

/** Sign in with ID + password. Returns a session token. */
function api_login(userId, password) {
  const id = String(userId || '').trim().toLowerCase(), pw = String(password || '');
  if (!id || !pw) throw new Error('Enter your ID and password.');
  const cache = CacheService.getScriptCache(), fk = 'f_' + id.replace(/[^a-z0-9]/g, '');
  const fails = Number(cache.get(fk) || 0);
  if (fails >= WEB.MAX_FAILS) throw new Error('Too many wrong attempts. Please wait 10 minutes and try again.');
  const sh = SpreadsheetApp.getActive().getSheetByName(WEB.USERS_SHEET), row = userRow_(sh, id);
  if (!row || !checkPw_(sh, row, pw)) {
    cache.put(fk, String(fails + 1), 600); Utilities.sleep(700);
    throw new Error('Wrong ID or password.');
  }
  cache.remove(fk);
  const r = sh.getRange(row, 1, 1, 3).getValues()[0];
  const user = { name: String(r[0]).trim(), id: String(r[1]).trim(), role: ROLES_.indexOf(String(r[2])) > -1 ? String(r[2]) : 'Member' };
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('s_' + token, JSON.stringify(user), WEB.SESSION_SECS);
  return { token: token, user: user };
}
function api_logout(token) { CacheService.getScriptCache().remove('s_' + cleanKey_(token)); return true; }
function api_changePassword(token, current, next) {
  const u = need_(token);
  next = String(next || '');
  if (next.length < 6) throw new Error('New password must be at least 6 characters.');
  const sh = SpreadsheetApp.getActive().getSheetByName(WEB.USERS_SHEET), row = userRow_(sh, String(u.id).toLowerCase());
  if (!row || !checkPw_(sh, row, String(current || ''))) throw new Error('Current password is wrong.');
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  sh.getRange(row, 4).setNumberFormat('@').setValue('h:' + salt + ':' + hashPw_(salt, next));
  return true;
}

function getUser_(token) {
  token = cleanKey_(token); if (!token) return null;
  const s = CacheService.getScriptCache().get('s_' + token);
  return s ? JSON.parse(s) : null;
}
function need_(token, roles) {
  const u = getUser_(token);
  if (!u) throw new Error('Your session has expired. Please sign in again.');
  if (roles && roles.indexOf(u.role) === -1) throw new Error('You do not have permission to do that.');
  return u;
}

/* ---------- reading ---------- */
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : ''; }
function num_(v) { return v === '' || v === null || isNaN(Number(v)) ? '' : Number(v); }

function taskFromRow_(r, rowNum, counts) {
  return {
    files: (counts && counts[String(r[0])]) || 0,
    row: rowNum, id: String(r[0]), task: String(r[1]), desc: String(r[2]), project: String(r[3]), category: String(r[4]),
    assignee: String(r[5]), assignDate: dstr_(r[6]), priority: String(r[7]), status: String(r[8]),
    start: dstr_(r[9]), due: dstr_(r[10]), completed: dstr_(r[11]), delivery: String(r[12]),
    est: num_(r[13]), actual: num_(r[14]), pct: num_(r[15]), daysLeft: num_(r[17]), health: String(r[18]),
    recurring: String(r[19]), blocked: String(r[20]), tags: String(r[21]), link: String(r[22]),
    updated: dstr_(r[23]), comments: String(r[24]), mailStatus: String(r[26]),
  };
}
function readTasks_(onlyFor) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.TASKS_SHEET);
  const v = sh.getRange(CFG.FIRST_ROW, 1, WEB.ROWS, 27).getValues();
  const out = [], counts = fileCounts_();
  v.forEach((r, i) => {
    if (r[1] === '') return;
    if (onlyFor && String(r[5]).trim() !== onlyFor) return;
    out.push(taskFromRow_(r, CFG.FIRST_ROW + i, counts));
  });
  return out;
}
function readOne_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.TASKS_SHEET);
  return taskFromRow_(sh.getRange(row, 1, 1, 27).getValues()[0], row, fileCounts_());
}
function readLists_() {
  const v = SpreadsheetApp.getActive().getSheetByName(CFG.LISTS_SHEET).getRange('A5:F30').getValues();
  const col = i => v.map(r => String(r[i]).trim()).filter(x => x);
  return { status: col(0), priority: col(1), project: col(2), category: col(3), assignees: col(4), recurring: col(5) };
}

/* ---------- API called from the page ---------- */
function api_init(key) {
  ensureTimeZone_();
  const u = getUser_(key);
  if (!u) return { error: 'invalid' };
  return {
    user: u, lists: readLists_(), today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
    tasks: readTasks_(u.role === 'Member' ? u.name : null),
  };
}

function api_logs(key) {
  need_(key, ['Admin', 'Manager']);
  const ss = SpreadsheetApp.getActive(), out = { updates: [], changes: [] };
  [[CFG.LOG_SHEET, 'updates', 8, 60], [CFG.TASKLOG_SHEET, 'changes', 7, 100]].forEach(x => {
    const sh = ss.getSheetByName(x[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const last = sh.getLastRow(), start = Math.max(2, last - x[3] + 1);
    out[x[1]] = sh.getRange(start, 1, last - start + 1, x[2]).getDisplayValues().reverse();
  });
  return out;
}

/** Admin: create or edit a task. Does everything the sheet automation does. */
function api_saveTask(key, t) {
  const u = need_(key, ['Admin']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const tasks = SpreadsheetApp.getActive().getSheetByName(CFG.TASKS_SHEET), T = CFG.T, L = readLists_();
    const out = [], by = u.name + ' (web app)';
    const title = String(t.task || '').trim();
    if (!title) throw new Error('Task title is required.');
    if (t.assignee && L.assignees.indexOf(t.assignee) === -1) throw new Error('Unknown person: ' + t.assignee);
    if (t.status && L.status.indexOf(t.status) === -1) throw new Error('Unknown status.');
    if (t.priority && L.priority.indexOf(t.priority) === -1) throw new Error('Unknown priority.');

    let row = Number(t.row) || 0;
    if (!row) {
      const col = tasks.getRange(CFG.FIRST_ROW, 2, WEB.ROWS, 1).getValues();
      const idx = col.findIndex(r => r[0] === '');
      if (idx < 0) throw new Error('The task table is full (' + WEB.ROWS + ' tasks).');
      row = CFG.FIRST_ROW + idx;
    } else if (row < CFG.FIRST_ROW || row >= CFG.FIRST_ROW + WEB.ROWS) throw new Error('Bad task row.');

    const D = s => s ? Utilities.parseDate(s, tz_(), 'yyyy-MM-dd') : '';
    const set = (col, val) => setAndLog_(tasks, row, col, val, by, out);
    const est = t.est === '' || t.est == null ? '' : Number(t.est);
    if (est !== '' && (isNaN(est) || est < 0)) throw new Error('Est. hours must be a number.');

    set(T.TASK, title); set(T.DESC, t.desc || ''); set(T.PROJECT, t.project || ''); set(T.CATEGORY, t.category || '');
    const oldOwner = String(tasks.getRange(row, T.OWNER).getValue()).trim(), newOwner = t.assignee || '';
    set(T.OWNER, newOwner);
    if (newOwner !== oldOwner) {
      set(T.ASSIGNED, newOwner ? todayDate_() : '');
      tasks.getRange(row, T.SEND).clearContent(); tasks.getRange(row, T.MAILSTATUS).clearContent();
    }
    set(T.PRIORITY, t.priority || '');
    set(T.START, D(t.start)); set(T.DUE, D(t.due)); set(T.EST, est);
    set(WEB.RECUR, t.recurring || ''); set(WEB.BLOCK, t.blocked || ''); set(WEB.TAGS, t.tags || ''); set(T.LINK, t.link || '');
    const oldStatus = String(tasks.getRange(row, T.STATUS).getValue());
    set(T.STATUS, t.status || '');
    if (t.status && t.status !== oldStatus) applyStatusDates_(tasks, row, t.status, out, by);
    tasks.getRange(row, T.UPDATED).setValue(todayDate_());

    let mail = '';
    if (t.sendMail && newOwner) {
      tasks.getRange(row, T.SEND).setValue('Send Mail');
      sendRow_(tasks, row);
      mail = tasks.getRange(row, T.MAILSTATUS).getDisplayValue();
      out.push(logRow_(tasks, row, T.MAILSTATUS, '', mail, 'Auto (script)', 'Email'));
    }
    writeLog_(CFG.TASKLOG_SHEET, out);
    SpreadsheetApp.flush();
    return { task: readOne_(row), mail: mail };
  } finally { lock.releaseLock(); }
}

/** Assignee (or admin): log progress on a task. Same as the green cells on the person tab. */
function api_updateProgress(key, row, up) {
  const u = need_(key, ['Admin', 'Member']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const tasks = SpreadsheetApp.getActive().getSheetByName(CFG.TASKS_SHEET), T = CFG.T, L = readLists_();
    row = Number(row);
    if (row < CFG.FIRST_ROW || row >= CFG.FIRST_ROW + WEB.ROWS || tasks.getRange(row, T.TASK).getValue() === '') throw new Error('Task not found.');
    const owner = String(tasks.getRange(row, T.OWNER).getValue()).trim();
    if (u.role === 'Member' && owner !== u.name) throw new Error('This task is not assigned to you.');
    const status = String(up.status || ''), note = String(up.note || '').trim();
    if (status && L.status.indexOf(status) === -1) throw new Error('Unknown status.');
    const hours = up.hours === '' || up.hours == null ? '' : Number(up.hours);
    if (hours !== '' && (isNaN(hours) || hours < 0)) throw new Error('Hours must be a number.');
    const pct = up.pct === '' || up.pct == null ? '' : Math.min(100, Math.max(0, Number(up.pct))) / 100;

    const out = [], by = u.name + ' (web app)', now = new Date();
    const oldStatus = String(tasks.getRange(row, T.STATUS).getValue());
    if (status) { setAndLog_(tasks, row, T.STATUS, status, by, out); if (status !== oldStatus) applyStatusDates_(tasks, row, status, out, by); }
    if (hours !== '') setAndLog_(tasks, row, T.ACTUAL, hours, by, out);
    if (pct !== '') setAndLog_(tasks, row, T.PCT, pct, by, out);
    if (note) setAndLog_(tasks, row, T.COMMENTS, Utilities.formatDate(now, CFG.TIMEZONE, 'dd MMM HH:mm') + ' - ' + note, by, out);
    tasks.getRange(row, T.UPDATED).setValue(todayDate_());

    const change = !status ? '-' : (oldStatus && oldStatus !== status ? oldStatus + ' → ' + status : status);
    writeLog_(CFG.LOG_SHEET, [[Utilities.formatDate(now, CFG.TIMEZONE, 'dd MMM yyyy HH:mm'), tasks.getRange(row, T.ID).getDisplayValue(),
      tasks.getRange(row, T.TASK).getValue(), owner, change, hours, pct, note]]);
    writeLog_(CFG.TASKLOG_SHEET, out);
    SpreadsheetApp.flush();
    return { task: readOne_(row) };
  } finally { lock.releaseLock(); }
}


/* ================= ATTACHMENTS (photos, PDFs, any file) ================= */
const FILES = {
  SHEET: 'Attachments', ROOT_NAME: 'TaskFlow Files', MAX_B64: 14000000,   // about 10 MB per file
  BLOCK: ['exe', 'bat', 'cmd', 'com', 'scr', 'msi', 'js', 'vbs', 'jar', 'ps1', 'sh', 'apk', 'dll', 'app'],
};

/** Run once from the editor: creates the Drive folder and the Attachments sheet (asks for Drive permission). */
function setupFiles() {
  rootFolder_(); filesSheet_();
  SpreadsheetApp.getUi().alert('Done. File uploads are ready. Files are stored in your Google Drive folder "' + FILES.ROOT_NAME + '".');
}
function rootFolder_() {
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('ROOT_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ } }
  const f = DriveApp.createFolder(FILES.ROOT_NAME); p.setProperty('ROOT_FOLDER', f.getId()); return f;
}
function taskFolder_(taskId) {
  const root = rootFolder_(), it = root.getFoldersByName(taskId);
  return it.hasNext() ? it.next() : root.createFolder(taskId);
}
function filesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FILES.SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(FILES.SHEET);
  const h = ['Uploaded', 'Task ID', 'Task', 'File name', 'Type', 'Size (KB)', 'Uploaded by', 'File ID'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setBackground('#0F172A').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1); sh.setTabColor('#A3E635');
  [150, 80, 240, 280, 160, 90, 140, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 1, sh.getMaxRows() - 1, 8).setNumberFormat('@');
  return sh;
}
function fileCounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(FILES.SHEET), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) m[String(r[0])] = (m[String(r[0])] || 0) + 1; });
  return m;
}
function fileList_(taskId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(FILES.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getDisplayValues()
    .filter(r => r[1] === taskId)
    .map(r => ({ when: r[0], name: r[3], mime: r[4], kb: Number(r[5]) || 0, by: r[6], id: r[7] })).reverse();
}
function fileRow_(fileId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(FILES.SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][7] === fileId) return { sheet: sh, row: i + 2, taskId: v[i][1], name: v[i][3], by: v[i][6] };
  return null;
}
/** Checks the person may see (or change) this task and returns its basics. */
function taskAccess_(u, row, write) {
  row = Number(row);
  const tasks = SpreadsheetApp.getActive().getSheetByName(CFG.TASKS_SHEET), T = CFG.T;
  if (!(row >= CFG.FIRST_ROW && row < CFG.FIRST_ROW + WEB.ROWS) || tasks.getRange(row, T.TASK).getValue() === '') throw new Error('Task not found.');
  const owner = String(tasks.getRange(row, T.OWNER).getValue()).trim();
  if (u.role === 'Member' && owner !== u.name) throw new Error('This task is not assigned to you.');
  if (write && u.role === 'Manager') throw new Error('Managers have view-only access.');
  return { row: row, id: tasks.getRange(row, T.ID).getDisplayValue(), title: String(tasks.getRange(row, T.TASK).getValue()), owner: owner, tasks: tasks };
}
function taskRowFromId_(taskId) {
  const n = Number(String(taskId).replace(/\D/g, ''));
  return n ? CFG.FIRST_ROW + n - 1 : 0;
}

function api_files(key, row) {
  const u = need_(key), ta = taskAccess_(u, row, false);
  return { files: fileList_(ta.id) };
}

function api_uploadFile(key, row, f) {
  const u = need_(key, ['Admin', 'Member']), ta = taskAccess_(u, row, true);
  const name = String(f && f.name || '').replace(/[\\/:*?"<>|]/g, '_').trim().slice(0, 120);
  if (!name) throw new Error('File has no name.');
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (FILES.BLOCK.indexOf(ext) > -1) throw new Error('For safety, ".' + ext + '" files cannot be uploaded.');
  const data = String(f.data || '');
  if (!data) throw new Error('The file is empty.');
  if (data.length > FILES.MAX_B64) throw new Error('File is too large (max 10 MB).');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(data), String(f.mime || 'application/octet-stream'), name);
    const file = taskFolder_(ta.id).createFile(blob);
    const sh = filesSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setNumberFormat('@').setValues([[
      Utilities.formatDate(new Date(), CFG.TIMEZONE, 'dd MMM yyyy HH:mm'), ta.id, ta.title, name,
      String(f.mime || 'application/octet-stream'), Math.max(1, Math.round(blob.getBytes().length / 1024)), u.name, file.getId()]]);
    writeLog_(CFG.TASKLOG_SHEET, [logRow_(ta.tasks, ta.row, CFG.T.TASK, '', name, u.name + ' (web app)', 'Attachment added')]);
    SpreadsheetApp.flush();
    return { files: fileList_(ta.id) };
  } finally { lock.releaseLock(); }
}

function api_getFile(key, fileId) {
  const u = need_(key), info = fileRow_(String(fileId));
  if (!info) throw new Error('File not found.');
  const row = taskRowFromId_(info.taskId);
  taskAccess_(u, row, false);
  const blob = DriveApp.getFileById(String(fileId)).getBlob(), bytes = blob.getBytes();
  if (bytes.length > 10 * 1024 * 1024) throw new Error('File is too large to open here. Open it from Google Drive instead.');
  return { name: info.name, mime: blob.getContentType() || 'application/octet-stream', data: Utilities.base64Encode(bytes) };
}

function api_deleteFile(key, fileId) {
  const u = need_(key, ['Admin', 'Member']), info = fileRow_(String(fileId));
  if (!info) throw new Error('File not found.');
  const ta = taskAccess_(u, taskRowFromId_(info.taskId), true);
  if (u.role !== 'Admin' && info.by !== u.name) throw new Error('You can only remove files you uploaded.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (e) { /* already gone */ }
    info.sheet.deleteRow(info.row);
    writeLog_(CFG.TASKLOG_SHEET, [logRow_(ta.tasks, ta.row, CFG.T.TASK, info.name, '', u.name + ' (web app)', 'Attachment removed')]);
    SpreadsheetApp.flush();
    return { files: fileList_(ta.id) };
  } finally { lock.releaseLock(); }
}

/** History of progress updates for one task (members can see their own tasks). */
function api_taskHistory(key, row) {
  const u = need_(key), ta = taskAccess_(u, row, false);
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return { history: [] };
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getDisplayValues().filter(r => r[1] === ta.id).reverse().slice(0, 30);
  return { history: v };
}
