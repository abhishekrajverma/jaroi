/**
 * MANAGEMENT DASHBOARD ("Command Center")
 *
 * Setup:  Extensions > Apps Script > click  +  next to Files > Script > name it  Dashboard
 *         paste this whole file > Save.
 *         Choose  buildManagementDashboard  in the function dropdown > Run.
 *
 * It creates a "Management" sheet (first tab) plus a hidden "Calc" sheet that holds its numbers.
 * Everything is live formulas, so it updates by itself. Run it again any time to rebuild the look.
 */

const DASH = {
  FIRST: 5, LAST: 204,   // task rows on the Tasks sheet
  C: { bg: '#0B1020', panel: '#121A2E', line: '#1F2A44', text: '#E2E8F0', muted: '#94A3B8', dim: '#64748B',
       cyan: '#22D3EE', violet: '#A78BFA', pink: '#F472B6', lime: '#A3E635', amber: '#FBBF24',
       red: '#FB7185', green: '#34D399', blue: '#60A5FA' },
};

function buildManagementDashboard() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('Tasks')) { SpreadsheetApp.getUi().alert('Sheet "Tasks" not found.'); return; }
  ['Management', 'Calc'].forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
  const calc = ss.insertSheet('Calc');
  const ms = ss.insertSheet('Management');
  buildCalc_(calc);
  buildMgmt_(ms, calc);
  calc.hideSheet();
  ss.setActiveSheet(ms);
  ss.moveActiveSheet(1);
}

/* ---------------- helpers ---------------- */
const R_ = col => 'Tasks!$' + col + '$' + DASH.FIRST + ':$' + col + '$' + DASH.LAST;
const OPEN_ = () => R_('I') + ',"<>✅ Done",' + R_('I') + ',"<>⚫ Cancelled",' + R_('B') + ',"<>"';
const OVERDUE_ = '"🔴 Overdue"';

/* ---------------- hidden Calc sheet ---------------- */
function buildCalc_(c) {
  const R = R_, OPEN = OPEN_();
  const metrics = [
    ['Total',      '=COUNTA(' + R('B') + ')'],
    ['Done',       '=COUNTIF(' + R('I') + ',"✅ Done")'],
    ['Open',       '=COUNTIFS(' + OPEN + ')'],
    ['Overdue',    '=COUNTIF(' + R('S') + ',' + OVERDUE_ + ')'],
    ['OnTime',     '=COUNTIF(' + R('M') + ',"*On time*")'],
    ['Late',       '=COUNTIF(' + R('M') + ',"*Late by*")'],
    ['AvgCycle',   '=IFERROR(ROUND(AVERAGE(FILTER(' + R('L') + '-' + R('J') + ',' + R('L') + '<>"",' + R('J') + '<>"")),1),0)'],
    ['OpenHours',  '=SUMIFS(' + R('N') + ',' + OPEN + ')'],
    ['Unassigned', '=COUNTIFS(' + R('B') + ',"<>",' + R('F') + ',"")'],
    ['Blocked',    '=COUNTIF(' + R('I') + ',"🟠 Blocked")'],
    ['Due7',       '=COUNTIFS(' + R('K') + ',">="&TODAY(),' + R('K') + ',"<="&TODAY()+7,' + OPEN + ')'],
  ];
  c.getRange(1, 1, 1, 2).setValues([['Metric', 'Value']]);
  c.getRange(2, 1, metrics.length, 1).setValues(metrics.map(m => [m[0]]));
  c.getRange(2, 2, metrics.length, 1).setFormulas(metrics.map(m => [m[1]]));

  // weekly flow (last 8 weeks, Monday start)
  c.getRange('D1:H1').setValues([['Week start', 'Week', 'Assigned', 'Completed', 'Open backlog']]);
  const wk = [];
  for (let i = 0; i < 8; i++) {
    const r = 2 + i;
    wk.push([
      '=TODAY()-WEEKDAY(TODAY(),2)+1-7*' + (7 - i),
      '=TEXT(D' + r + ',"d mmm")',
      '=COUNTIFS(' + R('G') + ',">="&D' + r + ',' + R('G') + ',"<"&(D' + r + '+7))',
      '=COUNTIFS(' + R('L') + ',">="&D' + r + ',' + R('L') + ',"<"&(D' + r + '+7))',
      '=COUNTIFS(' + R('G') + ',"<"&(D' + r + '+7))-COUNTIFS(' + R('L') + ',"<"&(D' + r + '+7))',
    ]);
  }
  c.getRange(2, 4, 8, 5).setFormulas(wk);
  c.getRange('D2:D9').setNumberFormat('dd mmm yyyy');

  // status counts
  c.getRange('J1:K1').setValues([['Status', 'Tasks']]);
  const st = [];
  for (let i = 0; i < 6; i++) st.push(['=Lists!A' + (5 + i), '=COUNTIF(' + R('I') + ',J' + (2 + i) + ')']);
  c.getRange(2, 10, 6, 2).setFormulas(st);

  // aging of open tasks (days since Assign Date)
  c.getRange('M1:N1').setValues([['Open for', 'Tasks']]);
  const G = R('G');
  const ag = [
    ['0-3 days',   '=COUNTIFS(' + OPEN + ',' + G + ',">="&TODAY()-3)'],
    ['4-7 days',   '=COUNTIFS(' + OPEN + ',' + G + ',">="&TODAY()-7,' + G + ',"<"&TODAY()-3)'],
    ['8-14 days',  '=COUNTIFS(' + OPEN + ',' + G + ',">="&TODAY()-14,' + G + ',"<"&TODAY()-7)'],
    ['15-30 days', '=COUNTIFS(' + OPEN + ',' + G + ',">="&TODAY()-30,' + G + ',"<"&TODAY()-14)'],
    ['30+ days',   '=COUNTIFS(' + OPEN + ',' + G + ',"<"&TODAY()-30)'],
  ];
  c.getRange(2, 13, 5, 1).setValues(ag.map(a => [a[0]]));
  c.getRange(2, 14, 5, 1).setFormulas(ag.map(a => [a[1]]));

  // top overdue (task, assignee, due, days late, priority)
  c.getRange('P1:T1').setValues([['Task', 'Assignee', 'Due', 'Days late', 'Priority']]);
  c.getRange('P2').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + R('B') + ',' + R('F') + ',' + R('K') + ',TODAY()-' + R('K') + ',' + R('H') + '},' + R('S') + '=' + OVERDUE_ + '),4,FALSE),8,5),"")');
  c.getRange('R2:R9').setNumberFormat('dd mmm yyyy');

  // due in next 7 days (task, assignee, due, priority)
  c.getRange('V1:Y1').setValues([['Task', 'Assignee', 'Due', 'Priority']]);
  c.getRange('V2').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + R('B') + ',' + R('F') + ',' + R('K') + ',' + R('H') + '},' +
    R('K') + '>=TODAY(),' + R('K') + '<=TODAY()+7,' + R('I') + '<>"✅ Done",' + R('I') + '<>"⚫ Cancelled",' + R('B') + '<>""),3,TRUE),8,4),"")');
  c.getRange('X2:X9').setNumberFormat('dd mmm yyyy');

  // recent activity from the Update Log (newest first)
  c.getRange('AA1:AE1').setValues([['When', 'Task', 'Assigned to', 'Status change', 'Note']]);
  const ul = "'Update Log'!A2:H";
  c.getRange('AA2').setFormula('=IFERROR(ARRAY_CONSTRAIN(CHOOSECOLS(SORT(FILTER(' + ul + ',' + "'Update Log'!A2:A" + '<>""),ROW(FILTER(' + "'Update Log'!A2:A" + ',' + "'Update Log'!A2:A" + '<>"")),FALSE),1,3,4,5,8),6,5),"")');
}

/* ---------------- Management sheet ---------------- */
function buildMgmt_(ms, calc) {
  const C = DASH.C, R = R_, OPEN = OPEN_();
  const LASTROW = 63;
  if (ms.getMaxColumns() > 14) ms.deleteColumns(15, ms.getMaxColumns() - 14);
  if (ms.getMaxRows() > LASTROW) ms.deleteRows(LASTROW + 1, ms.getMaxRows() - LASTROW);
  if (ms.getMaxRows() < LASTROW) ms.insertRowsAfter(ms.getMaxRows(), LASTROW - ms.getMaxRows());
  ms.setHiddenGridlines(true);
  ms.setTabColor(C.cyan);
  ms.setColumnWidth(1, 24); ms.setColumnWidths(2, 12, 98); ms.setColumnWidth(14, 24);

  const all = ms.getRange(1, 1, LASTROW, 14);
  all.setBackground(C.bg).setFontFamily('Roboto').setFontColor(C.text).setFontSize(10).setVerticalAlignment('middle');

  const rh = (row, n, h) => ms.setRowHeights(row, n, h);
  const panel = a1 => ms.getRange(a1).setBackground(C.panel);
  const gap = a1 => ms.getRange(a1).setBorder(null, true, null, null, null, null, C.bg, SpreadsheetApp.BorderStyle.SOLID_THICK);
  const top = (a1, color) => ms.getRange(a1).setBorder(true, null, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  const mg = a1 => ms.getRange(a1).merge();
  const put = (a1, v, o) => {
    const r = ms.getRange(a1); if (String(a1).indexOf(':') > -1) r.merge();
    if (typeof v === 'string' && v.charAt(0) === '=') r.setFormula(v); else r.setValue(v);
    if (o) {
      if (o.size) r.setFontSize(o.size);
      if (o.color) r.setFontColor(o.color);
      if (o.bold) r.setFontWeight('bold');
      if (o.font) r.setFontFamily(o.font);
      if (o.h) r.setHorizontalAlignment(o.h);
      if (o.fmt) r.setNumberFormat(o.fmt);
      if (o.wrap) r.setWrap(true);
    }
    return r;
  };
  const title = (a1, text, color) => put(a1, '   ◆  ' + text, { size: 10, bold: true, color: color, font: 'Montserrat', h: 'left' });
  const head = (a1, text, h) => put(a1, text, { size: 8, bold: true, color: C.dim, h: h || 'center' });
  const sbar = (v, max, color) => 'IFERROR(SPARKLINE({' + v + ',' + max + '-' + v + '},{"charttype","bar";"color1","' + color + '";"color2","' + C.line + '"}),"")';
  const line = a1 => ms.getRange(a1).setBorder(null, null, true, null, null, null, '#1B2540', SpreadsheetApp.BorderStyle.SOLID);

  // ---- header ----
  rh(1, 1, 14); rh(2, 1, 46); rh(3, 1, 22); rh(4, 1, 4); rh(5, 1, 14);
  put('B2:H2', 'COMMAND CENTER', { size: 26, bold: true, color: '#FFFFFF', font: 'Montserrat', h: 'left' });
  put('I2:M2', '="●  LIVE   "&TEXT(NOW(),"dd mmm yyyy  HH:mm")', { size: 10, bold: true, color: C.green, h: 'right' });
  put('B3:M3', 'Task Manager overview  •  every number and chart updates automatically from the Tasks sheet', { size: 9, color: C.muted, h: 'left' });
  ms.getRange('B4:D4').setBackground(C.cyan); ms.getRange('E4:G4').setBackground(C.violet);
  ms.getRange('H4:J4').setBackground(C.pink); ms.getRange('K4:M4').setBackground(C.amber);

  // ---- KPI tiles (rows 6-8) ----
  rh(6, 1, 24); rh(7, 1, 54); rh(8, 1, 30); rh(9, 1, 12);
  panel('B6:M8');
  const tiles = [
    ['B', 'C', 'TOTAL TASKS',      '=Calc!B2',                                       '0',       C.cyan,   '=IFERROR(SPARKLINE(Calc!F2:F9,{"charttype","column";"color","' + C.cyan + '"}),"")'],
    ['D', 'E', 'COMPLETION',       '=IFERROR(Calc!B3/Calc!B2,0)',                    '0%',      C.green,  '=' + sbar('IFERROR(Calc!B3/Calc!B2,0)', '1', C.green)],
    ['F', 'G', 'OVERDUE',          '=Calc!B5',                                       '0',       C.red,    '="of "&Calc!B4&" open tasks"'],
    ['H', 'I', 'ON-TIME DELIVERY', '=IFERROR(Calc!B6/(Calc!B6+Calc!B7),"–")',        '0%',      C.violet, '=Calc!B6&" on time  •  "&Calc!B7&" late"'],
    ['J', 'K', 'AVG CYCLE TIME',   '=Calc!B8',                                       '0.0" d"', C.amber,  'start → done'],
    ['L', 'M', 'OPEN WORKLOAD',    '=Calc!B9',                                       '0" h"',   C.pink,   '=Calc!B4&" open tasks"'],
  ];
  tiles.forEach((t, i) => {
    const a = t[0], b = t[1];
    put(a + '6:' + b + '6', t[2], { size: 8, bold: true, color: C.muted, h: 'center' });
    put(a + '7:' + b + '7', t[3], { size: 30, bold: true, color: t[5], font: 'Montserrat', h: 'center', fmt: t[4] });
    put(a + '8:' + b + '8', t[6], { size: 9, color: C.muted, h: 'center' });
    top(a + '6:' + b + '6', t[5]);
    if (i > 0) gap(a + '6:' + a + '8');
  });

  // ---- smart insights (row 10) ----
  rh(10, 1, 52); rh(11, 1, 14);
  panel('B10:M10');
  const ins = [
    ['B10:E10', '=IF(Calc!B5=0,"✅  No overdue tasks. Delivery is on track.","🔴  "&Calc!B5&IF(Calc!B5=1," task is"," tasks are")&" overdue. Longest delay: "&Calc!S2&" days ("&Calc!P2&").")', C.red],
    ['F10:I10', '=IF(MAX($D$28:$D$32)=0,"👥  No open tasks are assigned yet.","👥  "&INDEX($B$28:$B$32,MATCH(MAX($D$28:$D$32),$D$28:$D$32,0))&" carries the most open work ("&MAX($D$28:$D$32)&" tasks).")', C.cyan],
    ['J10:M10', '=IF(Calc!B6+Calc!B7=0,"⏱  Delivery stats appear once tasks are completed.","⏱  "&TEXT(Calc!B6/(Calc!B6+Calc!B7),"0%")&" of completed tasks were delivered on time."&IF(Calc!B11>0,"  "&Calc!B11&" blocked.",""))', C.violet],
  ];
  ins.forEach((x, i) => {
    put(x[0], x[1], { size: 10, color: C.text, h: 'left', wrap: true });
    ms.getRange(x[0]).setBorder(null, true, null, null, null, null, x[2], SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  // ---- flow chart + status donut (rows 12-24) ----
  rh(12, 1, 30); rh(13, 12, 22); rh(25, 1, 14);
  panel('B12:I24'); panel('J12:M24'); gap('J12:J24');
  title('B12:I12', 'FLOW · LAST 8 WEEKS', C.cyan);
  title('J12:M12', 'STATUS MIX', C.violet);
  ms.insertChart(ms.newChart().setChartType(Charts.ChartType.COMBO)
    .addRange(calc.getRange('E1:H9')).setNumHeaders(1).setPosition(13, 2, 4, 2)
    .setOption('width', 776).setOption('height', 258)
    .setOption('backgroundColor', C.panel).setOption('fontName', 'Roboto')
    .setOption('colors', [C.cyan, C.violet, C.amber])
    .setOption('seriesType', 'bars').setOption('series', { 2: { type: 'line', pointSize: 7, lineWidth: 3 } })
    .setOption('legend', { position: 'top', textStyle: { color: C.muted } })
    .setOption('hAxis', { textStyle: { color: C.muted } })
    .setOption('vAxis', { textStyle: { color: C.muted }, gridlines: { color: C.line }, minValue: 0 })
    .build());
  ms.insertChart(ms.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(calc.getRange('J1:K7')).setNumHeaders(1).setPosition(13, 10, 4, 2)
    .setOption('width', 384).setOption('height', 258)
    .setOption('backgroundColor', C.panel).setOption('fontName', 'Roboto')
    .setOption('pieHole', 0.62).setOption('pieSliceBorderColor', C.panel)
    .setOption('colors', ['#94A3B8', C.blue, '#FB923C', '#FACC15', C.green, '#475569'])
    .setOption('legend', { position: 'right', textStyle: { color: C.muted } })
    .setOption('pieSliceText', 'value').setOption('pieSliceTextStyle', { color: '#0B1020', bold: true })
    .build());

  // ---- team workload + priority (rows 26-32) ----
  rh(26, 1, 30); rh(27, 1, 22); rh(28, 5, 28); rh(33, 1, 14);
  panel('B26:H32'); panel('I26:M32'); gap('I26:I32');
  title('B26:H26', 'TEAM WORKLOAD', C.green);
  title('I26:M26', 'PRIORITY PRESSURE', C.pink);
  head('B27:C27', 'TEAM MEMBER', 'left'); head('D27', 'OPEN'); head('E27', 'OVERDUE'); head('F27', 'ON-TIME'); head('G27', 'LOAD'); head('H27', 'CYCLE');
  for (let i = 0; i < 5; i++) {
    const r = 28 + i, F = R('F');
    put('B' + r + ':C' + r, '=IF(Lists!E' + (5 + i) + '="","",Lists!E' + (5 + i) + ')', { size: 11, bold: true, h: 'left' });
    put('D' + r, '=IF($B' + r + '="","",COUNTIFS(' + F + ',$B' + r + ',' + OPEN + '))', { size: 12, bold: true, h: 'center' });
    put('E' + r, '=IF($B' + r + '="","",COUNTIFS(' + F + ',$B' + r + ',' + R('S') + ',' + OVERDUE_ + '))', { size: 12, bold: true, h: 'center' });
    put('F' + r, '=IF($B' + r + '="","",IFERROR(COUNTIFS(' + F + ',$B' + r + ',' + R('M') + ',"*On time*")/(COUNTIFS(' + F + ',$B' + r + ',' + R('M') + ',"*On time*")+COUNTIFS(' + F + ',$B' + r + ',' + R('M') + ',"*Late by*")),"–"))', { h: 'center', fmt: '0%' });
    put('G' + r, '=IF($B' + r + '="","",' + sbar('$D' + r, 'MAX($D$28:$D$32)', C.cyan) + ')', { h: 'center' });
    put('H' + r, '=IF($B' + r + '="","",IFERROR(ROUND(AVERAGE(FILTER(' + R('L') + '-' + R('J') + ',' + F + '=$B' + r + ',' + R('L') + '<>"",' + R('J') + '<>"")),1),"–"))', { h: 'center', fmt: '0.0" d"', color: C.muted });
    line('B' + r + ':H' + r);
  }
  head('I27:J27', 'PRIORITY', 'left'); head('K27', 'OPEN'); head('L27', 'OVERDUE'); head('M27', 'SHARE');
  for (let i = 0; i < 4; i++) {
    const r = 28 + i;
    put('I' + r + ':J' + r, '=Lists!B' + (5 + i), { size: 11, bold: true, h: 'left' });
    put('K' + r, '=COUNTIFS(' + R('H') + ',$I' + r + ',' + OPEN + ')', { size: 12, bold: true, h: 'center' });
    put('L' + r, '=COUNTIFS(' + R('H') + ',$I' + r + ',' + R('S') + ',' + OVERDUE_ + ')', { size: 12, bold: true, h: 'center' });
    put('M' + r, '=' + sbar('$K' + r, 'MAX($K$28:$K$31)', C.pink), { h: 'center' });
    line('I' + r + ':M' + r);
  }
  put('I32:J32', '⚠  Unassigned', { size: 11, bold: true, color: C.amber, h: 'left' });
  put('K32', '=Calc!B10', { size: 12, bold: true, color: C.amber, h: 'center' });

  // ---- project health + aging (rows 34-41) ----
  rh(34, 1, 30); rh(35, 1, 22); rh(36, 6, 28); rh(42, 1, 14);
  panel('B34:H41'); panel('I34:M41'); gap('I34:I41');
  title('B34:H34', 'PROJECT HEALTH', C.blue);
  title('I34:M34', 'TASK AGING · OPEN TASKS', C.amber);
  head('B35:C35', 'PROJECT', 'left'); head('D35', 'TASKS'); head('E35', 'DONE'); head('F35:G35', 'PROGRESS'); head('H35', 'HEALTH');
  for (let i = 0; i < 6; i++) {
    const r = 36 + i, D = R('D');
    put('B' + r + ':C' + r, '=IF(Lists!C' + (5 + i) + '="","",Lists!C' + (5 + i) + ')', { size: 11, bold: true, h: 'left' });
    put('D' + r, '=IF($B' + r + '="","",COUNTIF(' + D + ',$B' + r + '))', { size: 12, bold: true, h: 'center' });
    put('E' + r, '=IF($B' + r + '="","",COUNTIFS(' + D + ',$B' + r + ',' + R('I') + ',"✅ Done"))', { size: 12, bold: true, h: 'center' });
    put('F' + r + ':G' + r, '=IF($B' + r + '="","",IF($D' + r + '=0,"",' + sbar('$E' + r, '$D' + r, C.green) + '))', { h: 'center' });
    put('H' + r, '=IF($B' + r + '="","",IF($D' + r + '=0,"⚪ –",IF(COUNTIFS(' + D + ',$B' + r + ',' + R('S') + ',' + OVERDUE_ + ')>=2,"🔴 At risk",IF(COUNTIFS(' + D + ',$B' + r + ',' + R('S') + ',' + OVERDUE_ + ')=1,"🟡 Watch","🟢 Healthy"))))', { size: 10, bold: true, h: 'center' });
    line('B' + r + ':H' + r);
  }
  head('I35:K35', 'OPEN FOR', 'left'); head('L35', 'TASKS'); head('M35', 'SHARE');
  for (let i = 0; i < 5; i++) {
    const r = 36 + i;
    put('I' + r + ':K' + r, '=Calc!M' + (2 + i), { size: 11, bold: true, h: 'left' });
    put('L' + r, '=Calc!N' + (2 + i), { size: 12, bold: true, h: 'center' });
    put('M' + r, '=' + sbar('$L' + r, 'MAX($L$36:$L$40)', C.amber), { h: 'center' });
    line('I' + r + ':M' + r);
  }

  // ---- needs attention + due soon (rows 43-52) ----
  rh(43, 1, 30); rh(44, 1, 22); rh(45, 8, 28); rh(53, 1, 14);
  panel('B43:H52'); panel('I43:M52'); gap('I43:I52');
  title('B43:H43', 'NEEDS ATTENTION · OVERDUE', C.red);
  title('I43:M43', 'DUE IN THE NEXT 7 DAYS', C.amber);
  head('B44:D44', 'TASK', 'left'); head('E44', 'ASSIGNEE'); head('F44', 'DUE'); head('G44', 'LATE'); head('H44', 'PRIORITY');
  head('I44:J44', 'TASK', 'left'); head('K44', 'ASSIGNEE'); head('L44', 'DUE'); head('M44', 'PRIORITY');
  for (let k = 0; k < 8; k++) {
    const r = 45 + k, cr = 2 + k;
    put('B' + r + ':D' + r, k === 0 ? '=IF(Calc!P' + cr + '="","🎉  Nothing is overdue",Calc!P' + cr + ')' : '=IF(Calc!P' + cr + '="","",Calc!P' + cr + ')', { size: 10, bold: true, h: 'left' });
    put('E' + r, '=IF(Calc!Q' + cr + '="","",Calc!Q' + cr + ')', { h: 'center', color: C.muted });
    put('F' + r, '=IF(Calc!R' + cr + '="","",Calc!R' + cr + ')', { h: 'center', fmt: 'dd mmm' });
    put('G' + r, '=IF(Calc!S' + cr + '="","",Calc!S' + cr + '&" d")', { h: 'center', bold: true, color: C.red });
    put('H' + r, '=IF(Calc!T' + cr + '="","",Calc!T' + cr + ')', { h: 'center' });
    line('B' + r + ':H' + r);
    put('I' + r + ':J' + r, k === 0 ? '=IF(Calc!V' + cr + '="","✨  Nothing due this week",Calc!V' + cr + ')' : '=IF(Calc!V' + cr + '="","",Calc!V' + cr + ')', { size: 10, bold: true, h: 'left' });
    put('K' + r, '=IF(Calc!W' + cr + '="","",Calc!W' + cr + ')', { h: 'center', color: C.muted });
    put('L' + r, '=IF(Calc!X' + cr + '="","",Calc!X' + cr + ')', { h: 'center', fmt: 'dd mmm' });
    put('M' + r, '=IF(Calc!Y' + cr + '="","",Calc!Y' + cr + ')', { h: 'center' });
    line('I' + r + ':M' + r);
  }

  // ---- recent activity (rows 54-61) ----
  rh(54, 1, 30); rh(55, 1, 22); rh(56, 6, 28); rh(62, 1, 14); rh(63, 1, 24);
  panel('B54:M61');
  title('B54:M54', 'RECENT ACTIVITY · FROM THE UPDATE LOG', C.lime);
  head('B55:C55', 'WHEN', 'left'); head('D55:E55', 'TASK', 'left'); head('F55:G55', 'ASSIGNED TO', 'left'); head('H55:I55', 'STATUS CHANGE', 'left'); head('J55:M55', 'NOTE', 'left');
  for (let k = 0; k < 6; k++) {
    const r = 56 + k, cr = 2 + k;
    put('B' + r + ':C' + r, k === 0 ? '=IF(Calc!AA' + cr + '="","No updates yet. They appear here as people post progress.",Calc!AA' + cr + ')' : '=IF(Calc!AA' + cr + '="","",Calc!AA' + cr + ')', { size: 9, color: C.muted, h: 'left' });
    put('D' + r + ':E' + r, '=IF(Calc!AB' + cr + '="","",Calc!AB' + cr + ')', { bold: true, h: 'left' });
    put('F' + r + ':G' + r, '=IF(Calc!AC' + cr + '="","",Calc!AC' + cr + ')', { h: 'left' });
    put('H' + r + ':I' + r, '=IF(Calc!AD' + cr + '="","",Calc!AD' + cr + ')', { size: 9, h: 'left' });
    put('J' + r + ':M' + r, '=IF(Calc!AE' + cr + '="","",Calc!AE' + cr + ')', { size: 9, color: C.muted, h: 'left' });
    line('B' + r + ':M' + r);
  }
  put('B63:M63', 'Live view  •  built from the Tasks, Update Log and Task Log sheets', { size: 8, color: C.dim, h: 'center' });

  // ---- colour rules ----
  const rules = [];
  const add = (a1, fn) => rules.push(fn(SpreadsheetApp.newConditionalFormatRule().setRanges([ms.getRange(a1)])).build());
  add('E28:E32', b => b.whenNumberGreaterThan(0).setFontColor(C.red));
  add('L28:L31', b => b.whenNumberGreaterThan(0).setFontColor(C.red));
  add('H36:H41', b => b.whenTextContains('At risk').setFontColor(C.red));
  add('H36:H41', b => b.whenTextContains('Watch').setFontColor(C.amber));
  add('H36:H41', b => b.whenTextContains('Healthy').setFontColor(C.green));
  ms.setConditionalFormatRules(rules);
}
