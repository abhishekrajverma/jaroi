function onOpen() {
 SpreadsheetApp.getUi()
   .createMenu('🤖 Get System Prompt')
   .addItem('📋 Generate & Copy System Prompt', 'runGenerateSystemPrompt')
   .addSeparator()
   .addItem('📄 Save Prompt to New Sheet', 'savePromptToSheet')
   .addItem('🔍 Preview Raw Metadata JSON', 'previewMetadataJson')
   .addToUi();
}


function runGenerateSystemPrompt() {
 try {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const metadata = buildSpreadsheetMetadata(ss);
   const systemPrompt = constructSystemPrompt(metadata);
   showSystemPromptDialog(systemPrompt, metadata);
 } catch (e) {
   SpreadsheetApp.getUi().alert('Error: ' + e.message + '\n\nStack:\n' + e.stack);
 }
}


function savePromptToSheet() {
 try {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const metadata = buildSpreadsheetMetadata(ss);
   const systemPrompt = constructSystemPrompt(metadata);


   const sheetName = '_SystemPrompt_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
   const outputSheet = ss.insertSheet(sheetName);


   // Write system prompt split into rows to avoid cell character limits
   const lines = systemPrompt.split('\n');
   const data = lines.map(line => [line]);
   outputSheet.getRange(1, 1, data.length, 1).setValues(data);
   outputSheet.setColumnWidth(1, 900);


   SpreadsheetApp.getUi().alert('✅ System prompt saved to sheet: ' + sheetName);
 } catch (e) {
   SpreadsheetApp.getUi().alert('Error: ' + e.message);
 }
}

function previewMetadataJson() {
 try {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const metadata = buildSpreadsheetMetadata(ss);
   const json = JSON.stringify(metadata, null, 2);
   showTextDialog(json, '📊 Raw Metadata JSON');
 } catch (e) {
   SpreadsheetApp.getUi().alert('Error: ' + e.message);
 }
}

// ─────────────────────────────────────────────
//  METADATA BUILDER — TOP LEVEL
// ─────────────────────────────────────────────


function buildSpreadsheetMetadata(ss) {
 const sheets = ss.getSheets();


 const metadata = {
   spreadsheetName: ss.getName(),
   spreadsheetId: ss.getId(),
   url: ss.getUrl(),
   locale: ss.getSpreadsheetLocale(),
   timezone: Session.getScriptTimeZone(),
   totalSheets: sheets.length,
   generatedAt: new Date().toISOString(),
   namedRanges: getNamedRanges(ss),
   sheets: []
 };


 sheets.forEach((sheet, idx) => {
   metadata.sheets.push(buildSheetMetadata(sheet, idx));
 });

 return metadata;
}

function buildSheetMetadata(sheet, sheetIndex) {
 const sheetName = sheet.getName();
 const lastRow = sheet.getLastRow();
 const lastCol = sheet.getLastColumn();


 const base = {
   sheetName,
   sheetIndex,
   tabColor: getTabColor(sheet),
   isHidden: sheet.isSheetHidden(),
   frozenRows: sheet.getFrozenRows(),
   frozenColumns: sheet.getFrozenColumns(),
   sheetType: sheet.getType() === SpreadsheetApp.SheetType.GRID ? 'GRID' : 'OBJECT'
 };


 if (lastRow === 0 || lastCol === 0) {
   return Object.assign(base, { isEmpty: true });
 }


 // ── Ranges ──
 const endCol = colIndexToLetter(lastCol);
 const fullDataRange = `A1:${endCol}${lastRow}`;
 const headerRowRange = `A1:${endCol}1`;
 const dataBodyRange = lastRow > 1 ? `A2:${endCol}${lastRow}` : null;


 // ── Headers ──
 const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
 const headerRow = rawHeaders.map((h, i) => (String(h).trim() !== '' ? String(h).trim() : `Column_${i + 1}`));


 // ── Sample rows (first 5 data rows) ──
 const sampleCount = Math.min(5, lastRow - 1);
 let sampleData = [];
 if (sampleCount > 0) {
   const rawSample = sheet.getRange(2, 1, sampleCount, lastCol).getValues();
   sampleData = rawSample.map(row =>
     Object.fromEntries(headerRow.map((h, i) => [h, formatCellValue(row[i])]))
   );
 }


 // ── Column analysis (cap rows for performance) ──
 const MAX_ANALYSIS_ROWS = 2000;
 const analysisCount = Math.min(lastRow - 1, MAX_ANALYSIS_ROWS);
 let analysisData = [];
 if (analysisCount > 0) {
   analysisData = sheet.getRange(2, 1, analysisCount, lastCol).getValues();
 }
 const columnMetadata = analyzeAllColumns(headerRow, analysisData, lastRow - 1);


 // ── Derived column-type summary ──
 const categoricalCols = columnMetadata.filter(c => c.isCategorical).map(c => c.columnName);
 const numericCols = columnMetadata.filter(c => c.dataType === 'number').map(c => c.columnName);
 const dateCols = columnMetadata.filter(c => c.dataType === 'date').map(c => c.columnName);
 const booleanCols = columnMetadata.filter(c => c.dataType === 'boolean').map(c => c.columnName);
 const textCols = columnMetadata
   .filter(c => c.dataType === 'string' && !c.isCategorical)
   .map(c => c.columnName);


 // ── Conditional formatting & protections ──
 const conditionalFormats = getConditionalFormatSummary(sheet);
 const protection = getProtectionSummary(sheet);

 return Object.assign(base, {
   isEmpty: false,
   ranges: {
     fullDataRange,
     headerRowRange,
     dataBodyRange,
     totalRows: lastRow,
     totalColumns: lastCol,
     totalDataRows: lastRow - 1
   },
   headerRow,
   columnTypeSummary: {
     categorical: categoricalCols,
     numeric: numericCols,
     date: dateCols,
     boolean: booleanCols,
     freeText: textCols
   },
   columnMetadata,
   sampleData,
   conditionalFormats,
   protection
 });
}


function analyzeAllColumns(headers, data, totalDataRows) {
 return headers.map((header, colIdx) => {
   const rawValues = data.map(row => row[colIdx]);
   return analyzeColumn(header, colIdx, rawValues, totalDataRows);
 });
}


function analyzeColumn(colName, colIdx, rawValues, totalDataRows) {
 const nonEmpty = rawValues.filter(v => v !== '' && v !== null && v !== undefined);
 const emptyCount = rawValues.length - nonEmpty.length;


 // ── Type profiling ──
 const typeCounts = { number: 0, date: 0, boolean: 0, string: 0 };
 nonEmpty.forEach(v => {
   if (typeof v === 'boolean') typeCounts.boolean++;
   else if (v instanceof Date) typeCounts.date++;
   else if (typeof v === 'number') typeCounts.number++;
   else typeCounts.string++;
 });


 const total = nonEmpty.length || 1;
 const dominantType = getDominantType(typeCounts);


 const typeProfile = {
   number: pct(typeCounts.number, total),
   date: pct(typeCounts.date, total),
   boolean: pct(typeCounts.boolean, total),
   string: pct(typeCounts.string, total)
 };


 const uniqueSet = new Set(nonEmpty.map(v => stringifyValue(v)));
 const uniqueCount = uniqueSet.size;
 const uniqueValues = [...uniqueSet].sort();


 const isCategorical = detectCategorical(uniqueCount, totalDataRows, dominantType, uniqueValues);


 const result = {
   columnIndex: colIdx,
   columnLetter: colIndexToLetter(colIdx + 1),
   columnName: colName,
   dataType: dominantType,
   typeProfile,
   fillRate: pct(nonEmpty.length, rawValues.length),
   totalValues: rawValues.length,
   nonEmptyCount: nonEmpty.length,
   emptyCount,
   uniqueCount,
   isCategorical,
   categoricalOptions: isCategorical ? uniqueValues.slice(0, 60) : null,
   categoricalOptionCount: isCategorical ? uniqueValues.length : null
 };


 // ── Type-specific enrichment ──
 if (dominantType === 'number') {
   result.numericStats = computeNumericStats(nonEmpty.filter(v => typeof v === 'number'));
 }


 if (dominantType === 'date') {
   result.dateRange = computeDateRange(nonEmpty.filter(v => v instanceof Date));
 }


 // ── Semantic role hints ──
 result.semanticRole = inferSemanticRole(colName, dominantType, isCategorical, uniqueValues);


 return result;
}


function detectCategorical(uniqueCount, totalDataRows, dominantType, uniqueValues) {
 // Dates and plain free-text are never categorical
 if (dominantType === 'date') return false;
 if (dominantType === 'boolean') return true;


 // Numbers CAN be categorical if very low cardinality (e.g. rating 1-5, year)
 const uniqueRatio = uniqueCount / Math.max(totalDataRows, 1);


 if (dominantType === 'number') {
   // Treat as categorical only if ≤10 distinct values and looks like a code/rating
   return uniqueCount <= 10 && uniqueRatio < 0.05;
 }


 // String type — use layered thresholds
 if (uniqueCount <= 2) return true;                        // binary flag
 if (uniqueCount <= 15) return true;                        // classic enum
 if (uniqueCount <= 40 && uniqueRatio < 0.05) return true;  // larger enums in big sheets
 if (uniqueCount <= 80 && uniqueRatio < 0.02) return true;  // very large sheets, still low ratio

 return false;
}


const SEMANTIC_PATTERNS = {
 id: /\b(id|uuid|key|ref|code|number|no\.?|#)\b/i,
 name: /\b(name|title|label|description|desc|subject|summary)\b/i,
 status: /\b(status|state|stage|phase|flag)\b/i,
 category: /\b(category|type|kind|class|group|segment|tier|level)\b/i,
 department: /\b(department|dept|team|division|unit|function|org)\b/i,
 date: /\b(date|time|day|month|year|created|updated|modified|start|end|due|deadline)\b/i,
 amount: /\b(amount|price|cost|revenue|salary|budget|total|sum|value|fee|rate)\b/i,
 quantity: /\b(qty|quantity|count|number|vol|units|hours|days)\b/i,
 email: /\b(email|mail|e-mail)\b/i,
 phone: /\b(phone|mobile|tel|fax|contact)\b/i,
 location: /\b(city|country|state|region|location|address|zip|postal|district)\b/i,
 person: /\b(owner|assignee|manager|lead|rep|person|employee|agent|user|contact)\b/i,
 priority: /\b(priority|severity|urgency|importance|rank)\b/i,
 percentage: /\b(%|percent|pct|ratio|rate|share)\b/i,
 boolean: /\b(is_|has_|active|enabled|visible|approved|verified|flagged)\b/i
};


function inferSemanticRole(colName, dataType, isCategorical, uniqueValues) {
 const lower = colName.toLowerCase();


 for (const [role, pattern] of Object.entries(SEMANTIC_PATTERNS)) {
   if (pattern.test(lower)) return role;
 }


 // Fall back to type-driven roles
 if (dataType === 'date') return 'date';
 if (dataType === 'number') return 'metric';
 if (dataType === 'boolean') return 'flag';
 if (isCategorical) return 'category';
 return 'freeText';
}

// ─────────────────────────────────────────────
//  STATS HELPERS
// ─────────────────────────────────────────────


function computeNumericStats(nums) {
 if (!nums.length) return null;
 const sorted = [...nums].sort((a, b) => a - b);
 const sum = nums.reduce((a, b) => a + b, 0);
 const n = nums.length;
 const avg = sum / n;
 const mid = Math.floor(n / 2);
 const median = n % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
 const variance = nums.reduce((acc, v) => acc + Math.pow(v - avg, 2), 0) / n;


 return {
   min: sorted[0],
   max: sorted[n - 1],
   sum: Math.round(sum * 100) / 100,
   avg: Math.round(avg * 100) / 100,
   median: Math.round(median * 100) / 100,
   stdDev: Math.round(Math.sqrt(variance) * 100) / 100,
   count: n
 };
}


function computeDateRange(dates) {
 if (!dates.length) return null;
 const sorted = dates.sort((a, b) => a - b);
 return {
   earliest: Utilities.formatDate(sorted[0], 'UTC', 'yyyy-MM-dd'),
   latest: Utilities.formatDate(sorted[sorted.length - 1], 'UTC', 'yyyy-MM-dd'),
   spanDays: Math.round((sorted[sorted.length - 1] - sorted[0]) / 86400000)
 };
}






function getNamedRanges(ss) {
 try {
   return ss.getNamedRanges().map(nr => ({
     name: nr.getName(),
     range: nr.getRange().getA1Notation(),
     sheet: nr.getRange().getSheet().getName()
   }));
 } catch (e) { return []; }
}


function getConditionalFormatSummary(sheet) {
 try {
   return sheet.getConditionalFormatRules().map((rule, i) => ({
     ruleIndex: i,
     ranges: rule.getRanges().map(r => r.getA1Notation())
   }));
 } catch (e) { return []; }
}


function getProtectionSummary(sheet) {
 try {
   const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
   return {
     sheetProtected: sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).length > 0,
     protectedRangeCount: protections.length,
     protectedRanges: protections.map(p => p.getRange().getA1Notation())
   };
 } catch (e) { return { sheetProtected: false, protectedRangeCount: 0, protectedRanges: [] }; }
}


function getTabColor(sheet) {
 try {
   const c = sheet.getTabColorObject();
   return c ? c.asRgbColor().asHexString() : null;
 } catch (e) { return null; }
}






function colIndexToLetter(idx) {
 let letter = '';
 while (idx > 0) {
   const rem = (idx - 1) % 26;
   letter = String.fromCharCode(65 + rem) + letter;
   idx = Math.floor((idx - 1) / 26);
 }
 return letter;
}


function getDominantType(counts) {
 const order = ['date', 'boolean', 'number', 'string'];
 return order.reduce((best, t) => counts[t] > counts[best] ? t : best, 'string');
}


function pct(count, total) {
 return Math.round((count / Math.max(total, 1)) * 10000) / 100 + '%';
}


function formatCellValue(v) {
 if (v instanceof Date) return Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
 if (typeof v === 'boolean') return v;
 return v;
}


function stringifyValue(v) {
 if (v instanceof Date) return Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
 return String(v);
}






function constructSystemPrompt(metadata) {
 const ss = metadata;
 const sheetList = ss.sheets.map((s, i) =>
   `  ${i + 1}. "${s.sheetName}" — ${s.isEmpty ? 'EMPTY' : `${s.ranges.totalRows} rows × ${s.ranges.totalColumns} cols | Range: ${s.ranges.fullDataRange}`}`
 ).join('\n');


 const namedRangesBlock = ss.namedRanges.length
   ? ss.namedRanges.map(nr => `  - ${nr.name} → ${nr.sheet}!${nr.range}`).join('\n')
   : '  (none)';


 const sheetsDetail = ss.sheets.map(s => buildSheetBlock(s)).join('\n\n' + '─'.repeat(72) + '\n\n');


 return `${'═'.repeat(72)}
 GOOGLE SHEETS → DASHBOARD  |  AI CODING AGENT SYSTEM PROMPT
${'═'.repeat(72)}


You are an expert Google Apps Script and frontend dashboard developer.
Your task is to analyse the structured spreadsheet metadata below and
produce a fully functional, interactive dashboard as a Google Apps Script
Web App (HTML Service).


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SPREADSHEET INFORMATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Name          : ${ss.spreadsheetName}
Spreadsheet ID: ${ss.spreadsheetId}
URL           : ${ss.url}
Locale        : ${ss.locale}
Timezone      : ${ss.timezone}
Total Sheets  : ${ss.totalSheets}
Generated At  : ${ss.generatedAt}


NAMED RANGES:
${namedRangesBlock}


SHEETS OVERVIEW:
${sheetList}


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DETAILED SHEET METADATA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


${sheetsDetail}


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DASHBOARD CONSTRUCTION INSTRUCTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


1. DATA ACCESS
  • Use SpreadsheetApp.openById("${ss.spreadsheetId}") in Apps Script.
  • Do NOT hardcode data ranges (e.g., "A1:L500"). Always fetch data dynamically using sheet.getDataRange().getValues() or dynamically calculated ranges (like sheet.getLastRow()) to accommodate future data growth.
  • Filter out empty rows programmatically after fetching.
  • The first row of each sheet is always the HEADER row.


2. CATEGORICAL COLUMNS (isCategorical: true)
  • Render as multi-select dropdown filter chips or pill buttons.
  • Use the provided categoricalOptions as the known option list.
  • Map options to distinct colours for colour-coded tags/badges.
  • Columns with semanticRole "status" or "priority" are especially
    suited for Kanban-style grouping.


3. NUMERIC COLUMNS (dataType: "number")
  • Use numericStats (min/max/avg/sum) for KPI summary cards.
  • Plot in bar charts (comparisons), line charts (trends over dates),
    or scatter plots (correlations between two numeric columns).
  • Columns with semanticRole "amount" or "quantity" should have a
    currency/unit-aware formatter.


4. DATE COLUMNS (dataType: "date")
  • Add a date-range picker filter (start / end date).
  • Use for the X-axis of time-series charts.
  • dateRange.spanDays tells you the natural granularity
    (days < 60 → daily; < 730 → weekly; else monthly).


5. BOOLEAN COLUMNS (dataType: "boolean")
  • Render as toggle filters (All / True / False).
  • Show true/false ratio as a donut chart or progress bar.


6. FREE-TEXT COLUMNS (isCategorical: false, dataType: "string")
  • Include in a global search box (case-insensitive substring match).
  • Columns with semanticRole "name" or "description" are primary
    display fields in the data table.


7. DASHBOARD COMPONENT CHECKLIST
  ✅ Global search bar (searches all free-text + categorical columns)
  ✅ Filter panel: one control per categorical / boolean column
  ✅ Date range picker for every date column
  ✅ KPI cards: count of rows, sum/avg of key numeric columns
  ✅ At least one chart per numeric column (recommend Chart.js via CDN)
  ✅ Time-series chart for every date+numeric pair
  ✅ Sortable, paginated data table (all columns, sticky header)
  ✅ "Clear all filters" reset button
  ✅ Responsive layout (desktop + tablet)
  ✅ Sheet tabs/selector if multiple non-empty sheets


8. TECHNOLOGY STACK
  • Backend  : Google Apps Script (Code.gs) — data fetch & serve
  • Frontend : HTML + CSS + vanilla JS in a single index.html
  • Charts   : Chart.js 4.x from cdnjs.cloudflare.com
  • Styling  : Tailwind CSS Play CDN OR hand-written CSS variables
  • No external auth needed — deploy as Web App with sheet owner access.


9. PERFORMANCE GUIDELINES
  • Fetch data once on load, cache in a JS variable.
  • Apply all filters client-side (avoid repeated server calls).
  • Paginate the data table at 50 rows per page.
  • For sheets > 5 000 rows use server-side aggregation and pass only
    summary data to the frontend.


10. OUTPUT EXPECTED FROM YOU (THE CODING AGENT)
   • Code.gs  — Apps Script server file (data fetching, doGet)
   • index.html — Full dashboard HTML/CSS/JS
   • Optionally utils.gs for helper functions
   • Step-by-step deployment instructions


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RAW METADATA JSON  (ground truth — use this for exact column details)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


\`\`\`json
${JSON.stringify(metadata, null, 2)}
\`\`\`


${'═'.repeat(72)}
`;
}






function buildSheetBlock(sheet) {
 if (sheet.isEmpty) {
   return `SHEET [${sheet.sheetIndex}]: "${sheet.sheetName}"
 Status   : EMPTY
 Hidden   : ${sheet.isHidden}`;
 }


 const r = sheet.ranges;
 const cts = sheet.columnTypeSummary;


 const colMetaLines = sheet.columnMetadata.map(c => {
   let line = `    [${c.columnLetter}] ${c.columnName.padEnd(28)} | type=${String(c.dataType).padEnd(7)} | fill=${String(c.fillRate).padStart(6)} | unique=${String(c.uniqueCount).padStart(5)} | role=${c.semanticRole}`;


   if (c.isCategorical && c.categoricalOptions) {
     const shown = c.categoricalOptions.slice(0, 15).join(' | ');
     const more = c.categoricalOptions.length > 15 ? ` … (+${c.categoricalOptions.length - 15} more)` : '';
     line += `\n        → OPTIONS: ${shown}${more}`;
   }
   if (c.numericStats) {
     const ns = c.numericStats;
     line += `\n        → min=${ns.min}  max=${ns.max}  avg=${ns.avg}  sum=${ns.sum}  stdDev=${ns.stdDev}`;
   }
   if (c.dateRange) {
     line += `\n        → ${c.dateRange.earliest} → ${c.dateRange.latest}  (${c.dateRange.spanDays} days)`;
   }
   return line;
 }).join('\n');


 const sampleJson = JSON.stringify(sheet.sampleData, null, 4)
   .split('\n').map(l => '    ' + l).join('\n');


 return `SHEET [${sheet.sheetIndex}]: "${sheet.sheetName}"
 Tab Color      : ${sheet.tabColor || '—'}
 Hidden         : ${sheet.isHidden}
 Frozen Rows    : ${sheet.frozenRows}   Frozen Cols: ${sheet.frozenColumns}
 Full Range     : ${r.fullDataRange}
 Header Range   : ${r.headerRowRange}
 Body Range     : ${r.dataBodyRange || '(no data rows)'}
 Total Rows     : ${r.totalRows}  (${r.totalDataRows} data + 1 header)
 Total Columns  : ${r.totalColumns}


 COLUMN TYPE SUMMARY:
   Categorical (filters) : ${cts.categorical.length ? cts.categorical.join(', ') : '—'}
   Numeric (charts/KPIs) : ${cts.numeric.length ? cts.numeric.join(', ') : '—'}
   Date (time-series)    : ${cts.date.length ? cts.date.join(', ') : '—'}
   Boolean (toggles)     : ${cts.boolean.length ? cts.boolean.join(', ') : '—'}
   Free-text (search)    : ${cts.freeText.length ? cts.freeText.join(', ') : '—'}


 COLUMN DETAIL:
${colMetaLines}


 SAMPLE DATA (first ${sheet.sampleData.length} rows):
${sampleJson}`;
}






function showSystemPromptDialog(prompt, metadata) {
 const charCount = prompt.length;
 const sheetCount = metadata.sheets.filter(s => !s.isEmpty).length;


 const escapedPrompt = prompt
   .replace(/\\/g, '\\\\')
   .replace(/`/g, '\\`')
   .replace(/\$/g, '\\$');


 const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
 * { box-sizing: border-box; margin: 0; padding: 0; }
 body {
   font-family: 'Google Sans', Arial, sans-serif;
   background: #0f1117;
   color: #e8eaf0;
   height: 100vh;
   display: flex;
   flex-direction: column;
   padding: 12px;
   gap: 10px;
 }
 .header {
   display: flex;
   align-items: center;
   gap: 10px;
   flex-shrink: 0;
 }
 .header h2 {
   font-size: 15px;
   font-weight: 600;
   color: #c8d6ff;
   letter-spacing: 0.01em;
 }
 .badge {
   background: #1e2a4a;
   color: #7c9fff;
   font-size: 11px;
   padding: 2px 8px;
   border-radius: 10px;
   font-weight: 500;
   border: 1px solid #2a3a60;
 }
 .toolbar {
   display: flex;
   gap: 8px;
   flex-shrink: 0;
   align-items: center;
   flex-wrap: wrap;
 }
 .btn {
   display: inline-flex;
   align-items: center;
   gap: 5px;
   padding: 7px 14px;
   border-radius: 7px;
   border: none;
   font-size: 13px;
   font-weight: 500;
   cursor: pointer;
   transition: all 0.15s ease;
   font-family: inherit;
 }
 .btn-primary {
   background: #3b6ef0;
   color: white;
 }
 .btn-primary:hover { background: #4d7ef8; }
 .btn-secondary {
   background: #1e2533;
   color: #a8b8d8;
   border: 1px solid #2c3a55;
 }
 .btn-secondary:hover { background: #252f44; color: #c8d6ff; }
 .btn-success {
   background: #166534;
   color: #86efac;
   border: 1px solid #15803d;
 }
 .btn-success:hover { background: #15803d; }
 .stats {
   font-size: 11px;
   color: #556;
   margin-left: auto;
 }
 .textarea-wrap {
   flex: 1;
   position: relative;
   min-height: 0;
 }
 textarea {
   width: 100%;
   height: 100%;
   background: #090c14;
   color: #b8cce0;
   border: 1px solid #1e2a40;
   border-radius: 8px;
   padding: 14px;
   font-family: 'JetBrains Mono', 'Fira Code', 'Cascadia Code', 'Courier New', monospace;
   font-size: 11.5px;
   line-height: 1.65;
   resize: none;
   outline: none;
   tab-size: 2;
   letter-spacing: 0.01em;
 }
 textarea:focus { border-color: #3b6ef0; }
 .copy-flash {
   display: none;
   position: absolute;
   top: 10px;
   right: 10px;
   background: #166534;
   color: #86efac;
   padding: 5px 12px;
   border-radius: 6px;
   font-size: 12px;
   font-weight: 500;
 }
 .footer {
   flex-shrink: 0;
   font-size: 11px;
   color: #44506a;
   text-align: center;
   line-height: 1.5;
 }
</style>
</head>
<body>


<div class="header">
 <h2>🤖 System Prompt</h2>
 <span class="badge">${sheetCount} sheet${sheetCount !== 1 ? 's' : ''}</span>
 <span class="badge">${charCount.toLocaleString()} chars</span>
</div>


<div class="toolbar">
 <button class="btn btn-primary" onclick="copyPrompt()">📋 Copy Full Prompt</button>
 <button class="btn btn-secondary" onclick="selectAll()">⬜ Select All</button>
 <button class="btn btn-secondary" onclick="copyJson()">{ } Copy JSON Only</button>
 <span class="stats">Ready to paste into Claude, GPT-4, Gemini, or any AI coding agent</span>
</div>


<div class="textarea-wrap">
 <textarea id="prompt" spellcheck="false" readonly></textarea>
 <div class="copy-flash" id="flash">✓ Copied!</div>
</div>


<div class="footer">
 Paste this prompt into your AI coding agent → ask it to build Code.gs + index.html for a Google Apps Script Web App dashboard.
</div>


<script>
 const PROMPT = \`${escapedPrompt}\`;


 document.getElementById('prompt').value = PROMPT;


 function copyPrompt() {
   navigator.clipboard.writeText(PROMPT).then(showFlash).catch(fallbackCopy);
 }


 function fallbackCopy() {
   const ta = document.getElementById('prompt');
   ta.select();
   document.execCommand('copy');
   showFlash();
 }


 function selectAll() {
   const ta = document.getElementById('prompt');
   ta.focus();
   ta.select();
 }


 function copyJson() {
   const start = PROMPT.indexOf('\`\`\`json');
   const end   = PROMPT.lastIndexOf('\`\`\`');
   if (start !== -1 && end > start) {
     const json = PROMPT.slice(start + 7, end).trim();
     navigator.clipboard.writeText(json).then(showFlash).catch(() => {
       const ta = document.getElementById('prompt');
       ta.focus();
       ta.setSelectionRange(start, end + 3);
       document.execCommand('copy');
       showFlash();
     });
   }
 }


 function showFlash() {
   const f = document.getElementById('flash');
   f.style.display = 'block';
   setTimeout(() => { f.style.display = 'none'; }, 2000);
 }
</script>
</body>
</html>`)
   .setWidth(900)
   .setHeight(620)
   .setTitle('AI Dashboard System Prompt');


 SpreadsheetApp.getUi().showModalDialog(html, '🤖 System Prompt Generator');
}


function showTextDialog(text, title) {
 const escaped = text
   .replace(/\\/g, '\\\\')
   .replace(/`/g, '\\`')
   .replace(/\$/g, '\\$');


 const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
<style>
 body { margin: 8px; font-family: monospace; font-size: 12px; background:#0f1117; color:#b8cce0; }
 textarea { width:100%; height: 500px; background:#090c14; color:#b8cce0;
            border:1px solid #1e2a40; border-radius:6px; padding:10px;
            font-family: monospace; font-size: 12px; }
 button { margin-bottom:8px; padding:6px 14px; background:#3b6ef0; color:white;
          border:none; border-radius:6px; cursor:pointer; font-size:13px; }
</style>
</head>
<body>
<button onclick="document.getElementById('t').select();document.execCommand('copy')">📋 Copy JSON</button>
<textarea id="t" readonly></textarea>
<script>document.getElementById('t').value = \`${escaped}\`;</script>
</body>
</html>`)
   .setWidth(860).setHeight(560).setTitle(title);
 SpreadsheetApp.getUi().showModalDialog(html, title);
}

