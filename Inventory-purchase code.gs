/**
 * Inventory & Purchase System — dynamic Material ID dropdowns + reorder requirements
 * + automatic Material ID generation from Product name.
 * Assumptions: headers are on row 2, data starts on row 3.
 * Master sheet: Materials, Material ID in column A, Product header named "Product".
 * Stock Summary headers include Material ID, Available, MSL, Flag.
 * Requirements headers include Req ID, Req Dt, Material ID, Required Qty, Reason, Available stock.
 */

const INV_CFG = {
  masterSheet: 'Materials',
  stockSheet: 'Stock Summary',
  requirementsSheet: 'Requirements',
  headerRow: 2,
  firstDataRow: 3,
  dropdownRows: 4000
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Tools')
    .addItem('Refresh Material ID dropdowns', 'refreshMaterialIdDropdowns')
    .addItem('Sync reorder items to Requirements', 'syncReorderRequirements')
    .addItem('Generate missing Material IDs', 'generateMissingMaterialIds')
    .addSeparator()
    .addItem('Lock all Auto columns', 'protectAutoTaggedColumns')
    .addToUi();
}

function refreshMaterialIdDropdowns() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(INV_CFG.masterSheet);
  if (!master) throw new Error('Missing master sheet: ' + INV_CFG.masterSheet);

  const last = Math.max(master.getLastRow(), INV_CFG.firstDataRow);
  const source = master.getRange(INV_CFG.firstDataRow, 1, last - INV_CFG.firstDataRow + 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(false)
    .setHelpText('Select a Material ID from the Materials master.')
    .build();

  ss.getSheets().forEach(sh => {
    if (sh.getName() === INV_CFG.masterSheet) return;
    const headers = sh.getRange(INV_CFG.headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    headers.forEach((h, i) => {
      if (String(h).trim().toLowerCase() === 'material id') {
        const rows = Math.max(INV_CFG.dropdownRows, sh.getMaxRows() - INV_CFG.firstDataRow + 1);
        sh.getRange(INV_CFG.firstDataRow, i + 1, rows, 1).setDataValidation(rule);
      }
    });
  });
  ss.toast('Material ID dropdowns refreshed across sheets.', 'Inventory Tools', 5);
}

function onEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();

  // New product typed/pasted in Materials -> create its Material ID first,
  // so the dropdown refresh and reorder sync below already see it.
  if (name === INV_CFG.masterSheet) autoIdForEditedRange_(e.range);

  // Keep dropdowns current when the master list is edited, and update requirements
  // after a direct edit on Stock Summary. Formula-only changes need the menu or trigger below.
  if (name === INV_CFG.masterSheet) refreshMaterialIdDropdowns();
  if (name === INV_CFG.stockSheet || name === INV_CFG.masterSheet) syncReorderRequirements();
}

/* ------------------------------------------------------------------
 * AUTO MATERIAL ID  (Materials sheet, column A)
 * Format: PREFIX-001, built from the Product name:
 *   - same Product name already has IDs  -> reuse its prefix, next number (BIRP-004)
 *   - one word product                   -> first 3 letters   (Blade -> BLA-001)
 *   - several words                      -> first 3 letters of word 1 + first letter
 *                                           of word 2          (Air Gun -> AIRG-001)
 *   - prefix already used by another product -> one more letter is added until unique
 * Existing IDs are never changed. Text in (brackets) is ignored when building the prefix.
 * ------------------------------------------------------------------ */

function autoIdForEditedRange_(range) {
  const sh = range.getSheet();
  const pCol = getProductColumn_(sh);
  if (!pCol) return;
  if (range.getColumn() > pCol || range.getLastColumn() < pCol) return;   // Product column not touched

  const rows = [];
  for (let i = 0; i < range.getNumRows(); i++) rows.push(range.getRow() + i);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try {
    generateMaterialIds_(rows);
  } finally {
    lock.releaseLock();
  }
}

/** Menu: fill every blank Material ID that has a Product name. */
function generateMissingMaterialIds() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const count = generateMaterialIds_(null);
    SpreadsheetApp.getActive().toast(
      count ? count + ' Material ID(s) generated.' : 'No missing Material IDs found.',
      'Inventory Tools', 5);
    if (count) refreshMaterialIdDropdowns();
  } finally {
    lock.releaseLock();
  }
}

function getProductColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (!lastCol) return 0;
  const headers = sh.getRange(INV_CFG.headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const idx = headers.findIndex(h => String(h).trim().toLowerCase() === 'product');
  return idx + 1;   // 0 when there is no Product header
}

/** rowNumbers = sheet row numbers to fill, or null for all rows. Returns how many IDs were created. */
function generateMaterialIds_(rowNumbers) {
  const sh = SpreadsheetApp.getActive().getSheetByName(INV_CFG.masterSheet);
  if (!sh) throw new Error('Missing master sheet: ' + INV_CFG.masterSheet);
  const pCol = getProductColumn_(sh);
  if (!pCol) throw new Error('Materials needs a "Product" header on row ' + INV_CFG.headerRow + '.');

  const lastRow = sh.getLastRow();
  if (lastRow < INV_CFG.firstDataRow) return 0;
  const n = lastRow - INV_CFG.firstDataRow + 1;

  const ids = sh.getRange(INV_CFG.firstDataRow, 1, n, 1).getValues().map(r => r[0]);
  const products = sh.getRange(INV_CFG.firstDataRow, pCol, n, 1).getValues().map(r => r[0]);
  const only = rowNumbers
    ? new Set(rowNumbers.filter(r => r >= INV_CFG.firstDataRow && r <= lastRow).map(r => r - INV_CFG.firstDataRow))
    : null;

  const plan = planMaterialIds_(ids, products, only);
  if (!plan.changed.length) return 0;

  if (plan.changed.length <= 25) {
    plan.changed.forEach(i => sh.getRange(INV_CFG.firstDataRow + i, 1).setValue(plan.ids[i]));
  } else {
    sh.getRange(INV_CFG.firstDataRow, 1, n, 1).setValues(plan.ids.map(v => [v]));
  }
  SpreadsheetApp.flush();
  return plan.changed.length;
}

const ID_RULES = { prefixLetters: 3, separator: '-', digits: 3 };

function normalizeProductKey_(name) {
  return String(name || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Ordered list of prefix candidates for a product name. */
function prefixCandidates_(name) {
  const n = ID_RULES.prefixLetters;
  const toWords = text => String(text || '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean);
  let words = toWords(String(name || '').replace(/\([^)]*\)/g, ' '));
  if (!words.length) words = toWords(name);          // name was only "(something)"
  if (!words.length) return ['MAT'];
  const first = words[0];
  const rest = words.slice(1).join('');
  const out = [];
  if (rest) {
    const head = first.slice(0, n);
    for (let k = 1; k <= rest.length; k++) out.push(head + rest.slice(0, k));
  } else {
    out.push(first.slice(0, n));
    for (let k = n + 1; k <= first.length; k++) out.push(first.slice(0, k));
  }
  return out;
}

/**
 * Pure planner. ids/products are parallel arrays (one entry per Materials data row).
 * Fills blank IDs only (optionally just the row indexes in onlyIdx). Never changes an existing ID.
 */
function planMaterialIds_(ids, products, onlyIdx) {
  const owner = {};            // prefix  -> product key that owns it
  const prefixOfProduct = {};  // product key -> prefix
  const maxNo = {};            // prefix  -> highest number used
  ids.forEach((raw, i) => {
    const m = String(raw || '').trim().match(/^(.+)-(\d+)$/);
    if (!m) return;
    const p = m[1].toUpperCase();
    maxNo[p] = Math.max(maxNo[p] || 0, parseInt(m[2], 10));
    const key = normalizeProductKey_(products[i]);
    if (!key) return;
    if (!(p in owner)) owner[p] = key;
    if (!(key in prefixOfProduct)) prefixOfProduct[key] = p;
  });

  const out = ids.slice();
  const changed = [];
  products.forEach((prod, i) => {
    if (String(out[i] || '').trim()) return;
    const key = normalizeProductKey_(prod);
    if (!key) return;
    if (onlyIdx && !onlyIdx.has(i)) return;

    let prefix = prefixOfProduct[key];
    if (!prefix) {
      const cands = prefixCandidates_(prod);
      prefix = cands.find(c => !(c in owner) || owner[c] === key);
      for (let s = 2; !prefix; s++) {              // last resort: BASE2, BASE3 ...
        const c = cands[0] + s;
        if (!(c in owner) || owner[c] === key) prefix = c;
      }
      owner[prefix] = key;
      prefixOfProduct[key] = prefix;
    }
    maxNo[prefix] = (maxNo[prefix] || 0) + 1;
    out[i] = prefix + ID_RULES.separator + String(maxNo[prefix]).padStart(ID_RULES.digits, '0');
    changed.push(i);
  });
  return { ids: out, changed: changed };
}


function syncReorderRequirements() {
  const ss = SpreadsheetApp.getActive();
  const stock = ss.getSheetByName(INV_CFG.stockSheet);
  const req = ss.getSheetByName(INV_CFG.requirementsSheet);
  if (!stock || !req) throw new Error('Stock Summary or Requirements sheet is missing.');

  const stockData = stock.getDataRange().getValues();
  if (stockData.length <= INV_CFG.headerRow) return;
  const stockHeaders = stockData[INV_CFG.headerRow - 1].map(v => String(v).trim().toLowerCase());
  const col = header => stockHeaders.indexOf(header.toLowerCase());
  const idCol = col('Material ID');
  const availCol = col('Available');
  const mslCol = col('MSL');
  const flagCol = col('Flag');
  if ([idCol, availCol, mslCol, flagCol].some(i => i < 0)) {
    throw new Error('Stock Summary must contain Material ID, Available, MSL, and Flag headers on row 2.');
  }

  const reqLastCol = req.getLastColumn();
  const reqHeaders = req.getRange(INV_CFG.headerRow, 1, 1, reqLastCol).getDisplayValues()[0]
    .map(v => String(v).trim().toLowerCase());
  const rcol = header => reqHeaders.indexOf(header.toLowerCase());
  const rId = rcol('Material ID');
  const rReqId = rcol('Req ID');
  const rDate = rcol('Req Dt');
  const rQty = rcol('Required Qty');
  const rReason = rcol('Reason');
  const rAvail = rcol('Available stock');
  if ([rId, rReqId, rDate, rQty].some(i => i < 0)) {
    throw new Error('Requirements must contain Req ID, Req Dt, Material ID, and Required Qty headers on row 2.');
  }

  const reqLastRow = Math.max(req.getLastRow(), INV_CFG.firstDataRow - 1);
  const reqRows = reqLastRow >= INV_CFG.firstDataRow
    ? req.getRange(INV_CFG.firstDataRow, 1, reqLastRow - INV_CFG.firstDataRow + 1, reqLastCol).getValues()
    : [];
  const existingIds = new Set();
  reqRows.forEach(row => {
    const id = String(row[rId] || '').trim();
    if (id) existingIds.add(id);
  });

  const master = ss.getSheetByName(INV_CFG.masterSheet);
  const masterValues = master.getRange(INV_CFG.firstDataRow, 1,
    Math.max(1, master.getLastRow() - INV_CFG.firstDataRow + 1), 10).getValues();
  const materialDetails = new Map();
  masterValues.forEach(row => {
    const id = String(row[0] || '').trim();
    if (id) materialDetails.set(id, row.slice(1, 10));
  });

  const toAppend = [];
  for (let i = INV_CFG.headerRow; i < stockData.length; i++) {
    const row = stockData[i];
    const id = String(row[idCol] || '').trim();
    const flag = String(row[flagCol] || '').trim().toLowerCase();
    const available = Number(row[availCol]) || 0;
    const msl = Number(row[mslCol]) || 0;
    if (!id || flag !== 'reorder' || msl <= available || existingIds.has(id)) continue;

    const newRow = new Array(reqLastCol).fill('');
    newRow[rReqId] = 'REQ-' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd') +
      '-' + String(Math.max(req.getLastRow() + toAppend.length, INV_CFG.firstDataRow) - 1).padStart(5, '0');
    newRow[rDate] = new Date();
    newRow[rId] = id;
    newRow[rQty] = Math.max(0, msl - available); // replenish up to MSL, less currently available stock
    if (rReason >= 0) newRow[rReason] = 'Automatic reorder: stock below MSL';
    if (rAvail >= 0) newRow[rAvail] = available;

    // Populate any matching material detail columns in Requirements.
    const detailHeaders = ['product', 'finish', 'thickness', 'size', 'brand', 'code', 'stage', 'unit', 'category'];
    const details = materialDetails.get(id) || [];
    detailHeaders.forEach((h, j) => {
      const targetCol = rcol(h);
      if (targetCol >= 0 && !newRow[targetCol]) newRow[targetCol] = details[j] || '';
    });
    toAppend.push(newRow);
    existingIds.add(id);
  }

  if (toAppend.length) {
    const startRow = Math.max(req.getLastRow() + 1, INV_CFG.firstDataRow);
    req.getRange(startRow, 1, toAppend.length, reqLastCol).setValues(toAppend);
    ss.toast(toAppend.length + ' reorder item(s) added to Requirements.', 'Inventory Tools', 6);
  } else {
    ss.toast('No new reorder items found.', 'Inventory Tools', 4);
  }
}

/**
 * Optional: create an hourly trigger so formula-driven Stock Summary flags are checked
 * even when their source formulas recalculate without a direct edit to Stock Summary.
 * Run this function once manually and approve permissions.
 */
function installHourlyReorderSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncReorderRequirements') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncReorderRequirements').timeBased().everyHours(1).create();
}


/**
 * Protect every data column marked "Auto" in Column Guide.
 * Also protects dynamic Materials lookup columns whose headers are driven by Config.
 * Run once manually after installing, and again after changing the Column Guide/config.
 * Sheet owner can always edit protected ranges; other editors cannot unless explicitly
 * added as protection editors. Existing protections created by this function are replaced.
 */
function protectAutoTaggedColumns() {
  const ss = SpreadsheetApp.getActive();
  const guide = ss.getSheetByName('Column Guide');
  if (!guide) throw new Error('Missing sheet: Column Guide');

  const values = guide.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('Column Guide has no configuration rows.');
  const headers = values[0].map(v => String(v).trim().toLowerCase());
  const sheetIdx = headers.indexOf('sheet');
  const columnIdx = headers.indexOf('column');
  const autoIdx = headers.indexOf('typed or auto');
  if ([sheetIdx, columnIdx, autoIdx].some(i => i < 0)) {
    throw new Error('Column Guide must have Sheet, Column, and Typed or auto headers.');
  }

  // Sheet -> explicit auto column names, plus a flag for dynamic Material lookup columns.
  const autoNames = new Map();
  const dynamicFieldSheets = new Set();
  values.slice(1).forEach(row => {
    const sheetName = String(row[sheetIdx] || '').trim();
    const columnName = String(row[columnIdx] || '').trim();
    const type = String(row[autoIdx] || '').trim().toLowerCase();
    if (!sheetName || type !== 'auto') return;
    if (!autoNames.has(sheetName)) autoNames.set(sheetName, new Set());
    if (/^field\s*1\s*to\s*field\s*12$/i.test(columnName)) {
      dynamicFieldSheets.add(sheetName);
    } else {
      autoNames.get(sheetName).add(columnName.toLowerCase());
    }
  });

  // Remove only protections created by this function; preserve unrelated protections.
  ss.getSheets().forEach(sh => {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
      if (p.getDescription() === 'Inventory System - Auto column lock') p.remove();
    });
  });

  let protectedCount = 0;
  ss.getSheets().forEach(sh => {
    const sheetName = sh.getName();
    if (sheetName === 'Column Guide' || sheetName === 'Config') return;
    const names = autoNames.get(sheetName) || new Set();
    const lastCol = sh.getLastColumn();
    if (!lastCol) return;
    const row = sh.getRange(INV_CFG.headerRow, 1, 1, lastCol).getDisplayValues()[0];
    const formulas = sh.getRange(INV_CFG.headerRow, 1, 1, lastCol).getFormulas()[0];
    const columns = new Set();

    row.forEach((header, i) => {
      const normalized = String(header || '').trim().toLowerCase();
      if (normalized && names.has(normalized)) columns.add(i + 1);
      // Dynamic Material fields have header formulas referencing Config!$B$... .
      if (dynamicFieldSheets.has(sheetName) && formulas[i] && /config!\$b\$\d+/i.test(formulas[i])) {
        columns.add(i + 1);
      }
    });

    // Protect data rows only; keep headers available for administrators to maintain.
    columns.forEach(col => {
      const rowCount = Math.max(1, sh.getMaxRows() - INV_CFG.firstDataRow + 1);
      const range = sh.getRange(INV_CFG.firstDataRow, col, rowCount, 1);
      const protection = range.protect().setDescription('Inventory System - Auto column lock');
      protection.setWarningOnly(false);
      // Keep the script runner/owner as an editor, remove other editors from this protection.
      const editors = protection.getEditors();
      if (editors.length) protection.removeEditors(editors);
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      protectedCount++;
    });
  });
  ss.toast('Locked ' + protectedCount + ' Auto column(s).', 'Inventory Tools', 6);
}
