/**
 * Jaroi PO Tracking Only Web App
 * Deploy this as a SEPARATE Apps Script project.
 * Reads only PO-related fields from "Production Log" and machine labels from
 * "Machine Master"; it never returns other sheets or unrelated columns.
 *
 * Script Properties required:
 *   SPREADSHEET_ID = your source Google spreadsheet ID
 *   PO_USERS_JSON = JSON object with username keys and passwordHash + department
 *
 * Example PO_USERS_JSON:
 * {
 *   "cutting.team": {"passwordHash":"PUT_SHA256_HEX_HASH_HERE","department":"Cutting"},
 *   "purchase.team": {"passwordHash":"PUT_SHA256_HEX_HASH_HERE","department":"Purchase"}
 * }
 *
 * Generate a password hash with the included helper: run makePOPasswordHash("your password")
 * from the Apps Script editor, then copy the returned hash into PO_USERS_JSON.
 * Never put plain-text passwords in PO_USERS_JSON.
 */

const PO_STALL_DAYS = 3;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('POTracking')
    .setTitle('Jaroi | PO Tracking')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Public data endpoint. It returns only PO tracking fields. */
function getPOTrackingData() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  const ss = SpreadsheetApp.openById(id);
  const prod = ss.getSheetByName('Production Log');
  if (!prod) throw new Error('The source spreadsheet does not contain a "Production Log" sheet.');

  const production = readPORelevantTable_(prod);
  const machineMap = readMachineLabels_(ss.getSheetByName('Machine Master'));
  const headers = production.headers;
  const ix = {};
  headers.forEach((h, i) => ix[normaliseHeader_(h)] = i);

  const pick = (row, names) => {
    for (let i = 0; i < names.length; i++) {
      const j = ix[normaliseHeader_(names[i])];
      if (j != null && row[j] !== '' && row[j] != null) return row[j];
    }
    return null;
  };

  const entries = [];
  production.rows.forEach(row => {
    const po = pick(row, ['Po No', 'PO No', 'PO Number', 'Purchase Order No', 'Production Order']);
    const date = dateKey_(pick(row, ['Date', 'Production Date']));
    if (po == null || String(po).trim() === '' || !date) return;

    const machineId = String(pick(row, ['ID (Machine/Manual)', 'Machine ID', 'Resource ID']) || '');
    const machine = machineMap[machineId] || {};
    entries.push({
      po: String(po).trim(),
      date: date,
      startTime: safeString_(pick(row, ['Start Time'])),
      endTime: safeString_(pick(row, ['End Time'])),
      shift: safeString_(pick(row, ['Shift'])),
      machineId: machineId,
      machineName: machine.name || machineId,
      process: machine.process || '',
      operator: safeString_(pick(row, ['Operator'])),
      product: safeString_(pick(row, ['Product', 'Product / Part', 'Product Name'])),
      job: safeString_(pick(row, ['Job', 'Job No', 'Batch', 'Batch No'])),
      work: safeString_(pick(row, ['Work', 'Work Type'])),
      planned: numberValue_(pick(row, ['Planned Qty'])),
      actual: numberValue_(pick(row, ['Actual Qty'])),
      good: numberValue_(pick(row, ['Good Qty'])),
      rejected: numberValue_(pick(row, ['Rejected Qty', 'Reject Qty'])),
      rework: numberValue_(pick(row, ['Rework Qty']))
    });
  });

  entries.sort((a, b) => a.date.localeCompare(b.date) || a.startTime.localeCompare(b.startTime));
  return {
    entries: entries,
    generatedAt: new Date().toISOString(),
    stallDays: PO_STALL_DAYS
  };
}

function readPORelevantTable_(sheet) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };
  const values = sheet.getRange(1, 1, Math.min(lastRow, 12), lastCol).getDisplayValues();
  let headerRow = 0, best = -1;
  for (let r = 0; r < values.length; r++) {
    const filled = values[r].filter(v => String(v).trim()).length;
    const joined = values[r].map(normaliseHeader_).join('|');
    const score = filled + (/(po no|po number|production order)/.test(joined) ? 100 : 0) +
      (/(date|production date)/.test(joined) ? 20 : 0);
    if (score > best) { best = score; headerRow = r; }
  }
  const headers = values[headerRow].map(h => String(h || '').trim());
  const dataStart = headerRow + 2;
  if (dataStart > lastRow) return { headers: headers, rows: [] };
  const rows = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getValues();
  return { headers: headers, rows: rows };
}

function readMachineLabels_(sheet) {
  const map = {};
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return map;
  const range = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 12), sheet.getLastColumn()).getDisplayValues();
  let hr = 0, best = -1;
  for (let r = 0; r < range.length; r++) {
    const norm = range[r].map(normaliseHeader_);
    const score = range[r].filter(v => String(v).trim()).length +
      (norm.some(x => ['id (machine/manual)','machine id','resource id'].includes(x)) ? 100 : 0);
    if (score > best) { best = score; hr = r; }
  }
  const headers = range[hr].map(normaliseHeader_);
  const idCol = firstIndex_(headers, ['id (machine/manual)', 'machine id', 'resource id']);
  const nameCol = firstIndex_(headers, ['resource / machine name', 'machine name / model', 'machine name', 'resource name']);
  const processCol = firstIndex_(headers, ['process / department', 'process', 'department']);
  if (idCol < 0) return map;
  const start = hr + 2;
  if (start > sheet.getLastRow()) return map;
  sheet.getRange(start, 1, sheet.getLastRow() - start + 1, sheet.getLastColumn()).getDisplayValues().forEach(row => {
    const id = String(row[idCol] || '').trim();
    if (id) map[id] = { name: nameCol < 0 ? id : String(row[nameCol] || id), process: processCol < 0 ? '' : String(row[processCol] || '') };
  });
  return map;
}

function firstIndex_(arr, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const n = normaliseHeader_(candidates[i]);
    const ix = arr.indexOf(n);
    if (ix >= 0) return ix;
  }
  return -1;
}
function normaliseHeader_(value) { return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function safeString_(value) { return value == null ? '' : String(value).trim(); }
function numberValue_(value) {
  if (value == null || value === '') return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}
function dateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof value === 'number') {
    const d = new Date(Math.round((value - 25569) * 86400000));
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return '';
}
