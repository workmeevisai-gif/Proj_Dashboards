/**
 * ============================================================
 *  COMPANY EXPENSE DASHBOARD — GOOGLE SHEETS BUILDER
 * ============================================================
 *  Tracks under development:
 *    1. Aqua Bot            (funded — ₹2,00,000 grant)
 *    2. IC Tester           (no budget yet)
 *    3. CPS Training Kit    (no budget yet)
 *    + Company (General)    (non-project company expenses;
 *                            founders' contribution ₹10,000)
 *
 *  Sheets created:
 *    • Dashboard                          — funding, spend & balance overview
 *    • Funding                            — money received (grants, founder money)
 *    • <Project> - Expenses               — simple expense log (one per project)
 *    • Company - Expenses                 — other company expenses
 *    • <Project> - Utilization Certificate — UC in the standard format,
 *                                            generated for funded projects
 *
 *  You only type into the Funding sheet and the Expenses sheets.
 *  The Dashboard computes everything else; the Utilization
 *  Certificate is regenerated from the menu.
 *
 *  Rebuilding is SAFE: rows already entered in Funding and the
 *  Expenses sheets are captured and written back after the rebuild.
 *
 *  HOW TO RUN (one time):
 *    1. Open a new Google Sheet
 *    2. Extensions → Apps Script
 *    3. Paste this whole file, save
 *    4. Run the function  buildDashboard  (authorize when asked)
 *    5. Go back to the sheet — everything is generated
 *
 *  Afterwards use the "⚙ Dashboard Tools" menu inside the sheet:
 *    • Build / Rebuild all sheets
 *    • Refresh Utilization Certificates   ← run after adding expenses
 * ============================================================
 */

// ------------------------------------------------------------
// CONFIG
// ------------------------------------------------------------
const PROJECTS = ['Aqua Bot', 'IC Tester', 'CPS Training Kit'];

// Non-project company expenses are tracked under this label
const COMPANY = 'Company (General)';
const TRACKS = PROJECTS.concat([COMPANY]);

const CURRENCY_FORMAT = '₹#,##0.00';
const PLAIN_AMOUNT_FORMAT = '#,##0.00';   // UC table (₹ already in the header)
const PERCENT_FORMAT  = '0.0%';
const DATE_FORMAT     = 'dd-mmm-yyyy';

// Fiscal year start for the monthly table (April = Indian FY; 0 = January)
const FY_START_MONTH = 3;

// Chart palette (colour-blind-safe categorical order; grey = company/general)
const TRACK_COLORS = ['#2a78d6', '#008300', '#e87ba4', '#9aa0a6'];

// Styling
const HEADER_BG = '#1f2937';
const HEADER_FG = '#ffffff';
const KPI_BG    = '#eef3fb';

// Funding received so far — seeded on first build only. Edit the names!
// (On rebuilds, whatever is on the Funding sheet is kept instead.)
const INITIAL_FUNDING = [
  ['', 'Grant — Aqua Bot development', 'Aqua Bot', 200000, ''],
  ['', 'Founder contribution — Founder 1', COMPANY, 5000, 'Bank account opening'],
  ['', 'Founder contribution — Founder 2', COMPANY, 5000, 'Bank account opening'],
];

// Aqua Bot expenditure already incurred (from the Utilization Certificate).
// Dates were not on the record — fill them in on the sheet when known.
const AQUABOT_EXPENSES = [
  ['Nema 34 Motor (2) + motor driver module (2)', 26547],
  ['Mechanical hull structure manufacturing', 25002.96],
  ['FlySky i6 RC', 4683],
  ['IMU 9250', 438],
  ['Propellers (2)', 8326],
  ['Steel Pipe & Welding Materials', 2570],
  ['Welding Charge', 1800],
  ['Bearing & Water Seal', 420],
  ['Lathe work', 20000],
  ['Self-locking screw x 16', 72],
  ['Buck convertor & capacitors', 2762],
  ['12x13 SS bolt', 160],
  ['Rubber washer x20', 200],
  ['5MM PP Rope', 96],
  ['Rubber matt & drill bit', 133],
  ['Bearing', 260],
];

// ------------------------------------------------------------
// MENU
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ Dashboard Tools')
    .addItem('Build / Rebuild all sheets', 'buildDashboard')
    .addItem('Refresh Utilization Certificates', 'refreshUtilizationCertificates')
    .addToUi();
}

// ------------------------------------------------------------
// MAIN
// ------------------------------------------------------------
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const existing = ss.getSheets().map(s => s.getName());
  const willTouch = allManagedSheetNames().filter(n => existing.includes(n));
  if (willTouch.length > 0) {
    const resp = ui.alert(
      'Rebuild dashboard?',
      'This will rebuild these sheets:\n\n' + willTouch.join('\n') +
      '\n\nRows already entered in Funding and the Expenses sheets are KEPT.\n' +
      'Dashboard and Utilization Certificate sheets are regenerated.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  buildFundingSheet(ss);
  TRACKS.forEach(t => buildExpenseSheet(ss, t));
  buildOverviewSheet(ss);
  refreshUtilizationCertificates();
  orderSheets(ss);
  removeDefaultSheet(ss);

  ss.toast('Dashboard built successfully ✔', 'Done', 5);
}

function allManagedSheetNames() {
  const names = ['Dashboard', 'Funding'];
  TRACKS.forEach(t => names.push(expName(t)));
  PROJECTS.forEach(p => names.push(ucName(p)));
  return names;
}

function expName(track) {
  return track === COMPANY ? 'Company - Expenses' : track + ' - Expenses';
}
function ucName(project) { return project + ' - Utilization Certificate'; }

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

// Reads existing data rows so a rebuild never loses entered data.
// Returns null if the sheet doesn't exist; otherwise an array of rows
// (numCols wide, from startCol) that have at least one non-empty cell.
function captureTableRows(ss, name, startCol, numCols) {
  const sh = ss.getSheetByName(name);
  if (!sh) return null;
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, startCol, last - 1, numCols).getValues();
  return vals.filter(r => r.some(c => c !== '' && c !== null));
}

function fyStartDate() {
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() < FY_START_MONTH) year -= 1;
  return new Date(year, FY_START_MONTH, 1);
}

// ------------------------------------------------------------
// FUNDING — money received (grants, founder contributions, etc.)
// Columns: A Date | B Source / Details | C Track | D Amount | E Remarks
// ------------------------------------------------------------
function buildFundingSheet(ss) {
  const preserved = captureTableRows(ss, 'Funding', 1, 5);
  const sh = getFreshSheet(ss, 'Funding');
  const headers = ['Date', 'Source / Details', 'Project / Track', 'Amount', 'Remarks'];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);

  const rows = (preserved && preserved.length) ? preserved : INITIAL_FUNDING;
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const dataRows = 100;
  sh.getRange(2, 1, dataRows, 1).setNumberFormat(DATE_FORMAT);
  sh.getRange(2, 4, dataRows, 1).setNumberFormat(CURRENCY_FORMAT);

  const trackRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TRACKS, true).setAllowInvalid(true).build();
  sh.getRange(2, 3, dataRows, 1).setDataValidation(trackRule);

  sh.getRange(2, 1, dataRows, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  const widths = [110, 320, 170, 130, 260];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

// ------------------------------------------------------------
// EXPENSE LOG — one per project + one for company expenses
// Columns: A Si No (auto) | B Date | C Particulars of Expenditure
//          D Amount | E Remarks
// ------------------------------------------------------------
function buildExpenseSheet(ss, track) {
  const preserved = captureTableRows(ss, expName(track), 2, 4); // B..E
  const sh = getFreshSheet(ss, expName(track));
  const headers = ['Si No', 'Date', 'Particulars of Expenditure', 'Amount', 'Remarks'];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);

  // Keep entered rows on rebuild; seed Aqua Bot on first build
  let rows = (preserved && preserved.length) ? preserved : null;
  if (!rows && track === 'Aqua Bot') {
    rows = AQUABOT_EXPENSES.map(r => ['', r[0], r[1], '']);
  }
  if (rows && rows.length) sh.getRange(2, 2, rows.length, 4).setValues(rows);

  const dataRows = 400;
  // Si No numbers itself when a Particulars entry exists
  const siFormulas = [];
  for (let r = 2; r < 2 + dataRows; r++) {
    siFormulas.push(['=IF($C' + r + '<>"",COUNTA($C$2:$C' + r + '),"")']);
  }
  sh.getRange(2, 1, dataRows, 1).setFormulas(siFormulas)
    .setHorizontalAlignment('center').setFontColor('#6b7280')
    .setNumberFormat('0'); // plain number — otherwise Sheets can render 1, 2, 3 as dates

  sh.getRange(2, 2, dataRows, 1).setNumberFormat(DATE_FORMAT);
  sh.getRange(2, 4, dataRows, 1).setNumberFormat(CURRENCY_FORMAT);

  sh.getRange(2, 1, dataRows, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  const widths = [60, 110, 380, 130, 280];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

// ------------------------------------------------------------
// DASHBOARD
// ------------------------------------------------------------
function buildOverviewSheet(ss) {
  const sh = getFreshSheet(ss, 'Dashboard');

  sh.getRange('A1:H1').merge()
    .setValue('Company Expense Dashboard')
    .setFontSize(22).setFontWeight('bold');
  sh.getRange('A2:H2').merge()
    .setValue('Tracking: ' + TRACKS.join('  •  '))
    .setFontColor('#6b7280');

  const exp = TRACKS.map(t => "'" + expName(t) + "'");
  const spentSum = exp.map(e => 'SUM(' + e + '!$D$2:$D)').join('+');

  // KPI row (labels row 4, values row 5)
  const kpis = [
    ['FUNDING RECEIVED', '=SUM(Funding!$D$2:$D)'],
    ['TOTAL SPENT',      '=' + spentSum],
    ['ACCOUNT BALANCE',  '=A5-C5'],
    ['% OF FUNDS USED',  '=IF(A5=0,0,C5/A5)'],
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

  // Per-track summary (rows 7-12)
  sh.getRange(7, 1, 1, 6)
    .setValues([['Project / Track', 'Funding Received', 'Spent', 'Balance', '% Used', 'Usage']])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');

  TRACKS.forEach((t, i) => {
    const r = 8 + i;
    sh.getRange(r, 1).setValue(t);
    sh.getRange(r, 2).setFormula('=SUMIF(Funding!$C$2:$C,$A' + r + ',Funding!$D$2:$D)');
    sh.getRange(r, 3).setFormula('=SUM(' + exp[i] + '!$D$2:$D)');
    sh.getRange(r, 4).setFormula('=IF(B' + r + '=0,"",B' + r + '-C' + r + ')');
    sh.getRange(r, 5).setFormula('=IF(B' + r + '=0,"",C' + r + '/B' + r + ')');
    sh.getRange(r, 6).setFormula(
      '=IF(B' + r + '>0,SPARKLINE(MIN(C' + r + ',B' + r + '),' +
      '{"charttype","bar";"max",B' + r + ';"color1","' + TRACK_COLORS[i] + '"}),"")'
    );
  });
  const tRow = 8 + TRACKS.length; // 12
  sh.getRange(tRow, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(tRow, 2).setFormula('=SUM(B8:B' + (tRow - 1) + ')').setFontWeight('bold');
  sh.getRange(tRow, 3).setFormula('=SUM(C8:C' + (tRow - 1) + ')').setFontWeight('bold');
  sh.getRange(tRow, 4).setFormula('=B' + tRow + '-C' + tRow).setFontWeight('bold');
  sh.getRange(tRow, 5).setFormula('=IF(B' + tRow + '=0,"",C' + tRow + '/B' + tRow + ')').setFontWeight('bold');

  sh.getRange(8, 2, TRACKS.length + 1, 3).setNumberFormat(CURRENCY_FORMAT);
  sh.getRange(8, 5, TRACKS.length + 1, 1).setNumberFormat(PERCENT_FORMAT);

  // Over-utilization highlight on % Used
  const pctRange = sh.getRange(8, 5, TRACKS.length + 1, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1)
      .setFontColor('#e34948').setBold(true).setRanges([pctRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.75, 1)
      .setFontColor('#b45309').setRanges([pctRange]).build(),
  ]);

  // Monthly spend table (rows 14-26). Uses the expense Date column, so
  // rows without a date are counted in totals but not in the monthly view.
  const mHead = ['Month'].concat(TRACKS).concat(['Total']);
  sh.getRange(14, 1, 1, mHead.length).setValues([mHead])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');
  sh.getRange(15, 1).setValue(fyStartDate());
  for (let i = 1; i < 12; i++) {
    sh.getRange(15 + i, 1).setFormula('=EOMONTH(A' + (14 + i) + ',0)+1');
  }
  const lastTrackCol = String.fromCharCode(65 + TRACKS.length); // e.g. 'E' for 4 tracks
  for (let i = 0; i < 12; i++) {
    const r = 15 + i;
    TRACKS.forEach((t, j) => {
      sh.getRange(r, 2 + j).setFormula(
        '=SUMIFS(' + exp[j] + '!$D$2:$D,' + exp[j] + '!$B$2:$B,">="&$A' + r + ',' +
        exp[j] + '!$B$2:$B,"<"&EOMONTH($A' + r + ',0)+1)'
      );
    });
    sh.getRange(r, 2 + TRACKS.length).setFormula('=SUM(B' + r + ':' + lastTrackCol + r + ')');
  }
  sh.getRange(15, 1, 12, 1).setNumberFormat('mmm yyyy');
  sh.getRange(15, 2, 12, TRACKS.length + 1).setNumberFormat(CURRENCY_FORMAT);

  const widths = [180, 140, 130, 130, 100, 140, 20, 130];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Charts
  const fundingVsSpent = sh.newChart().asColumnChart()
    .addRange(sh.getRange(7, 1, TRACKS.length + 1, 3))
    .setNumHeaders(1)
    .setOption('title', 'Funding vs Spent')
    .setOption('colors', ['#2a78d6', '#008300'])
    .setOption('legend', { position: 'top' })
    .setPosition(28, 1, 0, 0)
    .build();
  sh.insertChart(fundingVsSpent);

  const sharePie = sh.newChart().asPieChart()
    .addRange(sh.getRange(8, 1, TRACKS.length, 1))
    .addRange(sh.getRange(8, 3, TRACKS.length, 1))
    .setOption('title', 'Spend Share')
    .setOption('colors', TRACK_COLORS)
    .setOption('legend', { position: 'right' })
    .setPosition(28, 5, 0, 0)
    .build();
  sh.insertChart(sharePie);

  const monthly = sh.newChart().asColumnChart()
    .addRange(sh.getRange(14, 1, 13, 1 + TRACKS.length))
    .setNumHeaders(1)
    .setOption('title', 'Monthly Spend')
    .setOption('isStacked', true)
    .setOption('colors', TRACK_COLORS)
    .setOption('legend', { position: 'top' })
    .setPosition(47, 1, 0, 0)
    .build();
  sh.insertChart(monthly);
}

// ------------------------------------------------------------
// UTILIZATION CERTIFICATE — regenerated snapshot, print-ready,
// in the standard format:
//   Si No | Particulars of Expenditure | Amount (₹)
//   ...
//   Total Expenditure          | <sum>
//   Balance Amount (if any)    | <funding - sum>
// Built for every PROJECT that has funding recorded (company
// general expenses don't need a UC).
// ------------------------------------------------------------
function refreshUtilizationCertificates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let built = 0;
  PROJECTS.forEach(p => {
    const funding = getTrackFunding(ss, p);
    if (funding <= 0) return; // UC only needed for funded projects
    buildUCSheet(ss, p, funding, getExpenseRows(ss, p));
    built++;
  });
  ss.toast(built + ' Utilization Certificate(s) refreshed ✔', 'Done', 5);
}

function getTrackFunding(ss, track) {
  const sh = ss.getSheetByName('Funding');
  if (!sh || sh.getLastRow() < 2) return 0;
  const vals = sh.getRange(2, 3, sh.getLastRow() - 1, 2).getValues(); // C:D
  return vals.reduce((sum, r) => sum + (r[0] === track ? Number(r[1]) || 0 : 0), 0);
}

function getExpenseRows(ss, track) {
  const sh = ss.getSheetByName(expName(track));
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 3, sh.getLastRow() - 1, 2).getValues(); // C:D
  return vals.filter(r => String(r[0]).trim() !== '');
}

function buildUCSheet(ss, project, funding, rows) {
  const sh = getFreshSheet(ss, ucName(project));

  sh.getRange('A1:C1').merge().setValue('UTILIZATION CERTIFICATE')
    .setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:C2').merge().setValue('Project: ' + project)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A3:C3').merge()
    .setValue('Funds Received: ₹' + formatAmount(funding))
    .setHorizontalAlignment('center');
  sh.getRange('A5:C5').merge()
    .setValue('The expenditure incurred is summarized below:')
    .setFontWeight('bold');

  const headRow = 6;
  sh.getRange(headRow, 1, 1, 3)
    .setValues([['Si No', 'Particulars of Expenditure', 'Amount (₹)']])
    .setFontWeight('bold').setHorizontalAlignment('center');

  const total = rows.reduce((s, r) => s + (Number(r[1]) || 0), 0);
  const body = rows.map((r, i) => [i + 1, r[0], r[1]]);
  if (body.length) {
    sh.getRange(headRow + 1, 1, body.length, 3).setValues(body)
      .setHorizontalAlignment('center');
  }

  const totalRow = headRow + body.length + 1;
  sh.getRange(totalRow, 2).setValue('Total Expenditure').setFontWeight('bold');
  sh.getRange(totalRow, 3).setValue(total).setFontWeight('bold');
  sh.getRange(totalRow + 1, 2).setValue('Balance Amount (if any)').setFontWeight('bold');
  sh.getRange(totalRow + 1, 3).setValue(funding - total).setFontWeight('bold');
  sh.getRange(totalRow, 3, 2, 1).setHorizontalAlignment('center');

  const table = sh.getRange(headRow, 1, body.length + 3, 3);
  table.setBorder(true, true, true, true, true, true);
  table.setVerticalAlignment('middle');
  sh.getRange(headRow + 1, 3, body.length + 2, 1).setNumberFormat(PLAIN_AMOUNT_FORMAT);

  sh.getRange(totalRow + 3, 1, 1, 3).merge()
    .setValue('Generated on ' + Utilities.formatDate(new Date(),
      ss.getSpreadsheetTimeZone(), 'dd-MMM-yyyy') +
      ' — refresh via ⚙ Dashboard Tools → Refresh Utilization Certificates.')
    .setFontSize(8).setFontColor('#9ca3af');

  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 140);
}

// Indian digit grouping: 200000 -> "2,00,000.00"
function formatAmount(n) {
  const parts = Number(n).toFixed(2).split('.');
  let intPart = parts[0];
  const sign = intPart.startsWith('-') ? '-' : '';
  if (sign) intPart = intPart.slice(1);
  let last3 = intPart.slice(-3);
  let rest = intPart.slice(0, -3);
  if (rest) {
    rest = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',');
    last3 = rest + ',' + last3;
  }
  return sign + last3 + '.' + parts[1];
}

// ------------------------------------------------------------
// HOUSEKEEPING
// ------------------------------------------------------------
function orderSheets(ss) {
  const order = ['Dashboard', 'Funding'];
  PROJECTS.forEach(p => {
    order.push(expName(p));
    if (ss.getSheetByName(ucName(p))) order.push(ucName(p));
  });
  order.push(expName(COMPANY));
  let pos = 1;
  order.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos++);
  });
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));
}

function removeDefaultSheet(ss) {
  const managed = allManagedSheetNames();
  const def = ss.getSheets().find(s =>
    !managed.includes(s.getName()) && s.getLastRow() === 0 && s.getLastColumn() === 0);
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
}
