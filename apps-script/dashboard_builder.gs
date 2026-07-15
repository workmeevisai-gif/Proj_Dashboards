/**
 * ============================================================
 *  COMPANY EXPENSE DASHBOARD — GOOGLE SHEETS BUILDER
 * ============================================================
 *  Builds a complete expense-tracking workbook for the three
 *  active projects:
 *
 *    1. Aqua Bot
 *    2. IC Tester
 *    3. CPS Training Kit
 *
 *  Sheets created:
 *    • Dashboard                       — company-wide overview + charts
 *    • <Project> - Expenses            — expense log (one per project)
 *    • <Project> - Fund Utilization    — budget vs spend (one per project)
 *    • Lists (hidden)                  — dropdown sources
 *
 *  HOW TO RUN (one time):
 *    1. Open a new Google Sheet
 *    2. Extensions → Apps Script
 *    3. Paste this whole file, save
 *    4. Run the function  buildDashboard  (authorize when asked)
 *    5. Go back to the sheet — everything is generated
 *
 *  After the first run you also get a "⚙ Dashboard Tools" menu
 *  inside the spreadsheet for rebuilding.
 * ============================================================
 */

// ------------------------------------------------------------
// CONFIG — edit these to match your company
// ------------------------------------------------------------
const PROJECTS = ['Aqua Bot', 'IC Tester', 'CPS Training Kit'];

const CATEGORIES = [
  'Components & Electronics',
  'PCB & Fabrication',
  'Mechanical & 3D Printing',
  'Software & Cloud',
  'Tools & Equipment',
  'Testing & Certification',
  'Travel & Logistics',
  'Salaries & Stipends',
  'Miscellaneous',
];

const PAYMENT_MODES = ['UPI', 'Bank Transfer / NEFT', 'Credit / Debit Card', 'Cash', 'Cheque', 'Other'];

const CURRENCY_FORMAT = '₹#,##0';        // change to '$#,##0.00' etc. if needed
const PERCENT_FORMAT  = '0.0%';
const DATE_FORMAT     = 'dd-mmm-yyyy';

// Fiscal year start for the monthly tables (April = Indian FY).
// Change FY_START_MONTH to 0 for January.
const FY_START_MONTH = 3; // 0-indexed: 3 = April

// Chart palette (colour-blind-safe categorical order; grey = "Miscellaneous")
const PALETTE = ['#2a78d6', '#008300', '#e87ba4', '#eda100', '#1baf7a', '#eb6834', '#4a3aa7', '#e34948', '#9aa0a6'];
const PROJECT_COLORS = PALETTE.slice(0, PROJECTS.length);

// Styling
const HEADER_BG   = '#1f2937';
const HEADER_FG   = '#ffffff';
const KPI_BG      = '#eef3fb';
const TITLE_SIZE  = 18;

// Sample seed data so charts render immediately — delete the rows
// in each Expenses sheet and replace with real entries.
const SAMPLE_ALLOCATION = [80000, 40000, 30000, 25000, 35000, 30000, 20000, 200000, 15000];
const SAMPLE_EXPENSES = {
  'Aqua Bot': [
    ['2026-04-08', 'Waterproof thruster motors (x4)', 'Components & Electronics', 'RoboKits India', 'INV-1041', 18500, 'Bank Transfer / NEFT', 'Ravi', 'Sample row — delete'],
    ['2026-05-14', 'Hull enclosure 3D print', 'Mechanical & 3D Printing', 'Fracktal Works', 'INV-2210', 6200, 'UPI', 'Ravi', 'Sample row — delete'],
    ['2026-06-20', 'Pool testing session', 'Testing & Certification', 'AquaLab', 'RCPT-88', 4000, 'Cash', 'Priya', 'Sample row — delete'],
  ],
  'IC Tester': [
    ['2026-04-18', 'ZIF sockets & test probes', 'Components & Electronics', 'Element14', 'E14-77813', 9400, 'Credit / Debit Card', 'Arun', 'Sample row — delete'],
    ['2026-05-30', '4-layer PCB prototype run', 'PCB & Fabrication', 'PCBPower', 'PP-5521', 12800, 'Bank Transfer / NEFT', 'Arun', 'Sample row — delete'],
    ['2026-07-02', 'Bench multimeter', 'Tools & Equipment', 'Tequipment', 'TQ-3319', 15600, 'Bank Transfer / NEFT', 'Priya', 'Sample row — delete'],
  ],
  'CPS Training Kit': [
    ['2026-04-25', 'Sensor module bulk order', 'Components & Electronics', 'Robu.in', 'RB-99120', 21300, 'Bank Transfer / NEFT', 'Meena', 'Sample row — delete'],
    ['2026-06-05', 'Curriculum software licence', 'Software & Cloud', 'LabVIEW EDU', 'NI-4402', 11000, 'Credit / Debit Card', 'Meena', 'Sample row — delete'],
    ['2026-07-10', 'Kit carry cases (x20)', 'Mechanical & 3D Printing', 'CaseCraft', 'CC-105', 7800, 'UPI', 'Ravi', 'Sample row — delete'],
  ],
};

// ------------------------------------------------------------
// MENU
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ Dashboard Tools')
    .addItem('Build / Rebuild all sheets', 'buildDashboard')
    .addToUi();
}

// ------------------------------------------------------------
// MAIN
// ------------------------------------------------------------
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Guard: rebuilding wipes existing dashboard sheets.
  const existing = ss.getSheets().map(s => s.getName());
  const willTouch = allManagedSheetNames().filter(n => existing.includes(n));
  if (willTouch.length > 0) {
    const resp = ui.alert(
      'Rebuild dashboard?',
      'This will ERASE and rebuild these sheets:\n\n' + willTouch.join('\n') +
      '\n\nAny data typed into them will be lost. Continue?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  buildListsSheet(ss);
  PROJECTS.forEach(p => {
    buildExpenseSheet(ss, p);
    buildUtilizationSheet(ss, p);
  });
  buildOverviewSheet(ss);
  orderSheets(ss);
  removeDefaultSheet(ss);

  ss.toast('Dashboard built successfully ✔', 'Done', 5);
}

function allManagedSheetNames() {
  const names = ['Dashboard', 'Lists'];
  PROJECTS.forEach(p => names.push(expName(p), fuName(p)));
  return names;
}

function expName(project) { return project + ' - Expenses'; }
function fuName(project)  { return project + ' - Fund Utilization'; }

function getFreshSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getCharts().forEach(c => sh.removeChart(c));
    sh.getBandings().forEach(b => b.remove());
    const full = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    full.breakApart();
    full.clearDataValidations();
    full.clearNote();
    if (sh.isSheetHidden()) sh.showSheet();
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function fyStartDate() {
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() < FY_START_MONTH) year -= 1;
  return new Date(year, FY_START_MONTH, 1);
}

// ------------------------------------------------------------
// LISTS (hidden dropdown sources)
// ------------------------------------------------------------
function buildListsSheet(ss) {
  const sh = getFreshSheet(ss, 'Lists');
  sh.getRange(1, 1).setValue('Categories');
  sh.getRange(1, 2).setValue('Payment Modes');
  sh.getRange(1, 3).setValue('Projects');
  sh.getRange(2, 1, CATEGORIES.length, 1).setValues(CATEGORIES.map(c => [c]));
  sh.getRange(2, 2, PAYMENT_MODES.length, 1).setValues(PAYMENT_MODES.map(m => [m]));
  sh.getRange(2, 3, PROJECTS.length, 1).setValues(PROJECTS.map(p => [p]));
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  sh.hideSheet();
}

// ------------------------------------------------------------
// EXPENSE LOG — one per project
// Columns: A Date | B Description | C Category | D Vendor | E Invoice
//          F Amount | G Payment Mode | H Paid By | I Notes
// ------------------------------------------------------------
function buildExpenseSheet(ss, project) {
  const sh = getFreshSheet(ss, expName(project));
  const headers = ['Date', 'Description', 'Category', 'Vendor / Paid To', 'Invoice / Bill No.',
                   'Amount', 'Payment Mode', 'Paid By', 'Notes'];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);

  // Sample rows
  const rows = (SAMPLE_EXPENSES[project] || []).map(r => {
    const copy = r.slice();
    copy[0] = new Date(copy[0]);
    return copy;
  });
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const dataRows = 400; // pre-formatted entry space
  sh.getRange(2, 1, dataRows, 1).setNumberFormat(DATE_FORMAT);
  sh.getRange(2, 6, dataRows, 1).setNumberFormat(CURRENCY_FORMAT);

  // Dropdowns
  const lists = ss.getSheetByName('Lists');
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lists.getRange(2, 1, CATEGORIES.length, 1), true)
    .setAllowInvalid(false).build();
  const payRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lists.getRange(2, 2, PAYMENT_MODES.length, 1), true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 3, dataRows, 1).setDataValidation(catRule);
  sh.getRange(2, 7, dataRows, 1).setDataValidation(payRule);

  // Banding on the data area
  sh.getRange(2, 1, dataRows, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  const widths = [110, 260, 200, 170, 130, 110, 160, 110, 220];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

// ------------------------------------------------------------
// FUND UTILIZATION — one per project
// ------------------------------------------------------------
function buildUtilizationSheet(ss, project) {
  const sh = getFreshSheet(ss, fuName(project));
  const exp = "'" + expName(project) + "'";

  // Title
  sh.getRange('A1:I1').merge()
    .setValue(project + ' — Fund Utilization')
    .setFontSize(TITLE_SIZE).setFontWeight('bold');

  // KPI row (labels row 3, values row 4)
  const kpis = [
    ['TOTAL BUDGET',   '=$B$17'],
    ['TOTAL UTILIZED', '=SUM(' + exp + '!$F$2:$F)'],
    ['REMAINING',      '=A4-C4'],
    ['% UTILIZED',     '=IF(A4=0,0,C4/A4)'],
  ];
  const kpiCols = [1, 3, 5, 7]; // A, C, E, G
  kpis.forEach((k, i) => {
    const c = kpiCols[i];
    sh.getRange(3, c, 1, 2).merge().setValue(k[0])
      .setFontSize(9).setFontColor('#6b7280').setFontWeight('bold').setBackground(KPI_BG);
    const v = sh.getRange(4, c, 1, 2).merge().setFormula(k[1])
      .setFontSize(16).setFontWeight('bold').setBackground(KPI_BG);
    v.setNumberFormat(i === 3 ? PERCENT_FORMAT : CURRENCY_FORMAT);
  });
  sh.setRowHeight(4, 32);

  // Category table (rows 7-17)
  const tHead = ['Category', 'Allocated Budget', 'Utilized', 'Remaining', '% Utilized', 'Usage'];
  sh.getRange(7, 1, 1, tHead.length).setValues([tHead])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');

  const firstRow = 8;
  CATEGORIES.forEach((cat, i) => {
    const r = firstRow + i;
    sh.getRange(r, 1).setValue(cat);
    sh.getRange(r, 2).setValue(SAMPLE_ALLOCATION[i] || 0); // <-- editable budget
    sh.getRange(r, 3).setFormula('=SUMIF(' + exp + '!$C$2:$C,$A' + r + ',' + exp + '!$F$2:$F)');
    sh.getRange(r, 4).setFormula('=B' + r + '-C' + r);
    sh.getRange(r, 5).setFormula('=IF(B' + r + '=0,"",C' + r + '/B' + r + ')');
    sh.getRange(r, 6).setFormula(
      '=IF(B' + r + '>0,SPARKLINE(MIN(C' + r + ',B' + r + '),' +
      '{"charttype","bar";"max",B' + r + ';"color1",' +
      'IF(C' + r + '/B' + r + '>0.9,"#e34948",IF(C' + r + '/B' + r + '>0.75,"#eda100","#2a78d6"))}),"")'
    );
  });
  const totalRow = firstRow + CATEGORIES.length; // 17
  sh.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(totalRow, 2).setFormula('=SUM(B' + firstRow + ':B' + (totalRow - 1) + ')').setFontWeight('bold');
  sh.getRange(totalRow, 3).setFormula('=SUM(C' + firstRow + ':C' + (totalRow - 1) + ')').setFontWeight('bold');
  sh.getRange(totalRow, 4).setFormula('=B' + totalRow + '-C' + totalRow).setFontWeight('bold');
  sh.getRange(totalRow, 5).setFormula('=IF(B' + totalRow + '=0,"",C' + totalRow + '/B' + totalRow + ')').setFontWeight('bold');

  sh.getRange(firstRow, 2, CATEGORIES.length + 1, 3).setNumberFormat(CURRENCY_FORMAT);
  sh.getRange(firstRow, 5, CATEGORIES.length + 1, 1).setNumberFormat(PERCENT_FORMAT);

  // Highlight allocated-budget input column
  sh.getRange(firstRow, 2, CATEGORIES.length, 1).setBackground('#fffbe6')
    .setNote('Editable: enter the allocated budget for this category.');

  // Over-utilization warning colours on % column
  const pctRange = sh.getRange(firstRow, 5, CATEGORIES.length + 1, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1)
      .setFontColor('#e34948').setBold(true).setRanges([pctRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.75, 1)
      .setFontColor('#b45309').setRanges([pctRange]).build(),
  ];
  sh.setConditionalFormatRules(rules);

  // Monthly spend table (H7:I19)
  sh.getRange(7, 8, 1, 2).setValues([['Month', 'Spend']])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');
  sh.getRange(8, 8).setValue(fyStartDate());
  for (let i = 1; i < 12; i++) {
    sh.getRange(8 + i, 8).setFormula('=EOMONTH(H' + (7 + i) + ',0)+1');
  }
  for (let i = 0; i < 12; i++) {
    const r = 8 + i;
    sh.getRange(r, 9).setFormula(
      '=SUMIFS(' + exp + '!$F$2:$F,' + exp + '!$A$2:$A,">="&H' + r + ',' +
      exp + '!$A$2:$A,"<"&EOMONTH(H' + r + ',0)+1)'
    );
  }
  sh.getRange(8, 8, 12, 1).setNumberFormat('mmm yyyy');
  sh.getRange(8, 9, 12, 1).setNumberFormat(CURRENCY_FORMAT);

  // Column widths
  const widths = [210, 140, 120, 120, 100, 140, 20, 100, 110];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Charts
  const pie = sh.newChart().asPieChart()
    .addRange(sh.getRange(firstRow, 1, CATEGORIES.length, 1))
    .addRange(sh.getRange(firstRow, 3, CATEGORIES.length, 1))
    .setOption('title', 'Utilization by Category')
    .setOption('colors', PALETTE)
    .setOption('legend', { position: 'right' })
    .setPosition(20, 1, 0, 0)
    .build();
  sh.insertChart(pie);

  const col = sh.newChart().asColumnChart()
    .addRange(sh.getRange(7, 8, 13, 2))
    .setNumHeaders(1)
    .setOption('title', 'Monthly Spend')
    .setOption('colors', ['#2a78d6'])
    .setOption('legend', { position: 'none' })
    .setPosition(20, 6, 0, 0)
    .build();
  sh.insertChart(col);
}

// ------------------------------------------------------------
// COMPANY DASHBOARD
// ------------------------------------------------------------
function buildOverviewSheet(ss) {
  const sh = getFreshSheet(ss, 'Dashboard');

  sh.getRange('A1:H1').merge()
    .setValue('Company Expense Dashboard')
    .setFontSize(22).setFontWeight('bold');
  sh.getRange('A2:H2').merge()
    .setValue('Projects: ' + PROJECTS.join('  •  '))
    .setFontColor('#6b7280');

  const fu = PROJECTS.map(p => "'" + fuName(p) + "'");

  // KPI row
  const budgetSum   = fu.map(f => f + '!$B$17').join('+');
  const utilizedSum = fu.map(f => f + '!$C$4').join('+');
  const kpis = [
    ['TOTAL BUDGET',   '=' + budgetSum],
    ['TOTAL UTILIZED', '=' + utilizedSum],
    ['REMAINING',      '=A5-C5'],
    ['% UTILIZED',     '=IF(A5=0,0,C5/A5)'],
  ];
  const kpiCols = [1, 3, 5, 7];
  kpis.forEach((k, i) => {
    const c = kpiCols[i];
    sh.getRange(4, c, 1, 2).merge().setValue(k[0])
      .setFontSize(9).setFontColor('#6b7280').setFontWeight('bold').setBackground(KPI_BG);
    const v = sh.getRange(5, c, 1, 2).merge().setFormula(k[1])
      .setFontSize(16).setFontWeight('bold').setBackground(KPI_BG);
    v.setNumberFormat(i === 3 ? PERCENT_FORMAT : CURRENCY_FORMAT);
  });
  sh.setRowHeight(5, 32);

  // Per-project summary (rows 7-11)
  sh.getRange(7, 1, 1, 6)
    .setValues([['Project', 'Budget', 'Utilized', 'Remaining', '% Utilized', 'Usage']])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');

  PROJECTS.forEach((p, i) => {
    const r = 8 + i;
    const f = fu[i];
    sh.getRange(r, 1).setValue(p);
    sh.getRange(r, 2).setFormula('=' + f + '!$B$17');
    sh.getRange(r, 3).setFormula('=' + f + '!$C$4');
    sh.getRange(r, 4).setFormula('=B' + r + '-C' + r);
    sh.getRange(r, 5).setFormula('=IF(B' + r + '=0,"",C' + r + '/B' + r + ')');
    sh.getRange(r, 6).setFormula(
      '=IF(B' + r + '>0,SPARKLINE(MIN(C' + r + ',B' + r + '),' +
      '{"charttype","bar";"max",B' + r + ';"color1","' + PROJECT_COLORS[i] + '"}),"")'
    );
  });
  const tRow = 8 + PROJECTS.length; // 11
  sh.getRange(tRow, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(tRow, 2).setFormula('=SUM(B8:B' + (tRow - 1) + ')').setFontWeight('bold');
  sh.getRange(tRow, 3).setFormula('=SUM(C8:C' + (tRow - 1) + ')').setFontWeight('bold');
  sh.getRange(tRow, 4).setFormula('=B' + tRow + '-C' + tRow).setFontWeight('bold');
  sh.getRange(tRow, 5).setFormula('=IF(B' + tRow + '=0,"",C' + tRow + '/B' + tRow + ')').setFontWeight('bold');

  sh.getRange(8, 2, PROJECTS.length + 1, 3).setNumberFormat(CURRENCY_FORMAT);
  sh.getRange(8, 5, PROJECTS.length + 1, 1).setNumberFormat(PERCENT_FORMAT);

  // Monthly spend by project (rows 14-26)
  const mHead = ['Month'].concat(PROJECTS).concat(['Total']);
  sh.getRange(14, 1, 1, mHead.length).setValues([mHead])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');
  for (let i = 0; i < 12; i++) {
    const r = 15 + i;
    sh.getRange(r, 1).setFormula('=' + fu[0] + '!$H$' + (8 + i));
    PROJECTS.forEach((p, j) => {
      sh.getRange(r, 2 + j).setFormula('=' + fu[j] + '!$I$' + (8 + i));
    });
    sh.getRange(r, 2 + PROJECTS.length).setFormula('=SUM(B' + r + ':' +
      String.fromCharCode(65 + PROJECTS.length) + r + ')');
  }
  sh.getRange(15, 1, 12, 1).setNumberFormat('mmm yyyy');
  sh.getRange(15, 2, 12, PROJECTS.length + 1).setNumberFormat(CURRENCY_FORMAT);

  const widths = [180, 130, 130, 130, 100, 140, 20, 130];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Charts
  const budgetVsUtilized = sh.newChart().asColumnChart()
    .addRange(sh.getRange(7, 1, PROJECTS.length + 1, 3))
    .setNumHeaders(1)
    .setOption('title', 'Budget vs Utilized by Project')
    .setOption('colors', ['#2a78d6', '#008300'])
    .setOption('legend', { position: 'top' })
    .setPosition(28, 1, 0, 0)
    .build();
  sh.insertChart(budgetVsUtilized);

  const sharePie = sh.newChart().asPieChart()
    .addRange(sh.getRange(8, 1, PROJECTS.length, 1))
    .addRange(sh.getRange(8, 3, PROJECTS.length, 1))
    .setOption('title', 'Spend Share by Project')
    .setOption('colors', PROJECT_COLORS)
    .setOption('legend', { position: 'right' })
    .setPosition(28, 5, 0, 0)
    .build();
  sh.insertChart(sharePie);

  const monthly = sh.newChart().asColumnChart()
    .addRange(sh.getRange(14, 1, 13, 1 + PROJECTS.length))
    .setNumHeaders(1)
    .setOption('title', 'Monthly Spend by Project')
    .setOption('isStacked', true)
    .setOption('colors', PROJECT_COLORS)
    .setOption('legend', { position: 'top' })
    .setPosition(47, 1, 0, 0)
    .build();
  sh.insertChart(monthly);

  // Over-utilization highlight on the project % column
  const pctRange = sh.getRange(8, 5, PROJECTS.length + 1, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1)
      .setFontColor('#e34948').setBold(true).setRanges([pctRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.75, 1)
      .setFontColor('#b45309').setRanges([pctRange]).build(),
  ]);
}

// ------------------------------------------------------------
// HOUSEKEEPING
// ------------------------------------------------------------
function orderSheets(ss) {
  const order = ['Dashboard'];
  PROJECTS.forEach(p => order.push(expName(p), fuName(p)));
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));
}

function removeDefaultSheet(ss) {
  const managed = allManagedSheetNames();
  const def = ss.getSheets().find(s => !managed.includes(s.getName()) && s.getLastRow() === 0);
  if (def && ss.getSheets().length > managed.length) ss.deleteSheet(def);
}
