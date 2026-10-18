/**
 * Builds the whole workbook from an empty spreadsheet: sheets, formulas,
 * formatting, validation, charts and demo data. Run `setup` once from the
 * Apps Script editor; the menu item "🧪 สร้างข้อมูลตัวอย่างใหม่" runs it again.
 */

var GREEN = '#188038';
var GREEN_DARK = '#0d652d';
var GREEN_SOFT = '#e6f4ea';
var INK = '#202124';
var MUTE = '#5f6368';
var FONT = 'Sarabun';

function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Bangkok');
  ss.setSpreadsheetLocale('th_TH');
  buildWorkbook(ss, new Date());
}

function resetDemoData() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert('สร้างข้อมูลตัวอย่างใหม่',
    'ข้อมูลสินค้าและรายการทั้งหมดในไฟล์นี้จะถูกแทนที่ด้วยข้อมูลตัวอย่าง ' + DEMO_DAYS + ' วันล่าสุด ต้องการทำต่อไหม?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  setup();
  ui.alert('เสร็จแล้ว', 'สร้างข้อมูลตัวอย่างใหม่เรียบร้อย ลองเปิดเมนู 📦 ระบบสต็อก › ➕ บันทึกรับเข้า / ขาย', ui.ButtonSet.OK);
}

function buildWorkbook(ss, endDate) {
  var names = [SHEET.DASH, SHEET.PRODUCTS, SHEET.TX, SHEET.DAILY, SHEET.HELP];
  // Start clean: a temporary sheet lets every named sheet be deleted.
  var tmp = ss.insertSheet('_tmp_' + Date.now());
  ss.getSheets().forEach(function (sh) { if (sh.getSheetId() !== tmp.getSheetId()) ss.deleteSheet(sh); });
  ss.getNamedRanges().forEach(function (r) { r.remove(); });
  names.forEach(function (name, i) { ss.insertSheet(name, i); });
  ss.deleteSheet(tmp);

  // Reference date for every "last 30 days" formula: the latest day with data,
  // so the demo never shows an empty "today". Named first so formulas resolve.
  var refCell = ss.getSheetByName(SHEET.DASH).getRange('N1');
  refCell.setFormula("=MAX('" + SHEET.TX + "'!L2:L)").setNumberFormat('d/m/yyyy').setFontColor('#ffffff');
  ss.setNamedRange('REF_DATE', refCell);

  var products = demoProducts();
  var movements = generateDemoMovements(products, endDate, DEMO_DAYS, DEMO_SEED);
  buildProducts(ss.getSheetByName(SHEET.PRODUCTS), products);
  buildTransactions(ss.getSheetByName(SHEET.TX), movements);
  buildDaily(ss.getSheetByName(SHEET.DAILY));
  buildDashboard(ss);
  buildHelp(ss.getSheetByName(SHEET.HELP));
  names.forEach(function (name) { ss.getSheetByName(name).getDataRange().setFontFamily(FONT); });
  ss.setActiveSheet(ss.getSheetByName(SHEET.DASH));
  SpreadsheetApp.flush();
}

function header(sh, labels, color) {
  var r = sh.getRange(1, 1, 1, labels.length);
  r.setValues([labels]).setFontWeight('bold').setFontColor('#ffffff').setBackground(color || GREEN)
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);
}

function trimSheet(sh, rows, cols) {
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
}

// ---------- สินค้า ----------

function buildProducts(sh, products) {
  var tx = "'" + SHEET.TX + "'!";
  header(sh, ['รหัส', 'ชื่อสินค้า', 'หมวด', 'หน่วย', 'ต้นทุน/หน่วย', 'ราคาขาย/หน่วย', 'จุดสั่งซื้อ']);
  trimSheet(sh, 200, 13);
  sh.getRange(2, 1, products.length, 7).setValues(products.map(function (p) {
    return [p.code, p.name, p.category, p.unit, p.cost, p.price, p.reorder];
  }));
  var f = {
    H1: '={"คงเหลือ"; ARRAYFORMULA(IF(A2:A="",, SUMIF(' + tx + 'C2:C, A2:A, ' + tx + 'J2:J)))}',
    I1: '={"มูลค่าคงเหลือ (ทุน)"; ARRAYFORMULA(IF(A2:A="",, H2:H * E2:E))}',
    // SUMIFS does not expand inside ARRAYFORMULA, so per-row sums use MAP.
    J1: '={"ขาย 30 วัน (ชิ้น)"; MAP(A2:A, LAMBDA(c, IF(c = "",, SUMIFS(' + tx + 'E2:E, ' + tx + 'C2:C, c, ' + tx + 'B2:B, "ขาย", ' + tx + 'L2:L, ">" & (REF_DATE - 30)))))}',
    K1: '={"ยอดขาย 30 วัน"; MAP(A2:A, LAMBDA(c, IF(c = "",, SUMIFS(' + tx + 'G2:G, ' + tx + 'C2:C, c, ' + tx + 'B2:B, "ขาย", ' + tx + 'L2:L, ">" & (REF_DATE - 30)))))}',
    L1: '={"พอขายอีก (วัน)"; ARRAYFORMULA(IF(A2:A="",, IF(J2:J = 0, "-", ROUNDDOWN(H2:H / (J2:J / 30)))))}',
    M1: '={"สถานะ"; ARRAYFORMULA(IF(A2:A="",, IF(H2:H <= 0, "หมด", IF(H2:H <= G2:G, "ใกล้หมด", "ปกติ"))))}',
  };
  Object.keys(f).forEach(function (a1) { sh.getRange(a1).setFormula(f[a1]); });
  sh.getRange('H1:M1').setFontWeight('bold').setFontColor('#ffffff').setBackground(GREEN_DARK).setWrap(true);

  sh.getRange('E2:F').setNumberFormat('#,##0');
  sh.getRange('G2:H').setNumberFormat('#,##0');
  sh.getRange('I2:I').setNumberFormat('#,##0');
  sh.getRange('J2:J').setNumberFormat('#,##0');
  sh.getRange('K2:K').setNumberFormat('#,##0');
  sh.getRange('L2:L').setHorizontalAlignment('right');
  sh.getRange('M2:M').setHorizontalAlignment('center');
  [80, 230, 120, 60, 95, 105, 90, 80, 120, 100, 110, 95, 85].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenColumns(2);

  var status = sh.getRange('M2:M');
  var stock = sh.getRange('H2:H');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('หมด').setBackground('#fce8e6').setFontColor('#c5221f').setBold(true).setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ใกล้หมด').setBackground('#fef7e0').setFontColor('#b06000').setBold(true).setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ปกติ').setFontColor(GREEN).setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"", $H2<=$G2)').setFontColor('#c5221f').setBold(true).setRanges([stock]).build(),
  ]);
  sh.getRange('E2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false).setHelpText('ใส่ตัวเลขตั้งแต่ 0 ขึ้นไป').build());
  sh.getRange('H1:M').protect().setDescription('คำนวณอัตโนมัติ (ไม่ต้องแก้)').setWarningOnly(true);
  sh.getRange(1, 1, products.length + 1, 13).createFilter();
}

// ---------- รายการ ----------

function buildTransactions(sh, movements) {
  var pr = "'" + SHEET.PRODUCTS + "'!";
  header(sh, ['วันที่เวลา', 'ประเภท', 'รหัสสินค้า', '', 'จำนวน', 'ราคาต่อหน่วย', '', 'หมายเหตุ', 'ผู้บันทึก']);
  trimSheet(sh, movements.length + 500, 12);
  var rand = seededRandom(DEMO_SEED + 1);
  var rows = movements.map(function (m) {
    var by = m.type === TYPE.SALE ? DEMO_STAFF[Math.floor(rand() * DEMO_STAFF.length)] : DEMO_STAFF[0];
    return [m.date, m.type, m.code, m.qty, m.price, m.note, by];
  });
  sh.getRange(2, 1, rows.length, 3).setValues(rows.map(function (r) { return r.slice(0, 3); }));
  sh.getRange(2, 5, rows.length, 2).setValues(rows.map(function (r) { return r.slice(3, 5); }));
  sh.getRange(2, 8, rows.length, 2).setValues(rows.map(function (r) { return r.slice(5, 7); }));

  var f = {
    D1: '={"ชื่อสินค้า"; ARRAYFORMULA(IF(C2:C="",, IFERROR(VLOOKUP(C2:C, ' + pr + 'A2:B, 2, FALSE), "⚠️ ไม่พบรหัสสินค้า")))}',
    G1: '={"ยอดเงิน"; ARRAYFORMULA(IF(C2:C="",, E2:E * F2:F))}',
    J1: '={"เปลี่ยนแปลงสต็อก"; ARRAYFORMULA(IF(C2:C="",, IF(B2:B = "ขาย", -ABS(E2:E), E2:E)))}',
    K1: '={"ต้นทุนขาย"; ARRAYFORMULA(IF((C2:C = "") + (B2:B <> "ขาย"),, E2:E * IFERROR(VLOOKUP(C2:C, ' + pr + 'A2:E, 5, FALSE), 0)))}',
    L1: '={"วันที่"; ARRAYFORMULA(IF(A2:A="",, INT(A2:A)))}',
  };
  Object.keys(f).forEach(function (a1) { sh.getRange(a1).setFormula(f[a1]); });
  sh.getRange('D1').setFontWeight('bold').setFontColor('#ffffff').setBackground(GREEN_DARK);
  sh.getRange('G1').setFontWeight('bold').setFontColor('#ffffff').setBackground(GREEN_DARK);
  sh.getRange('J1:L1').setFontWeight('bold').setFontColor('#ffffff').setBackground('#80868b');

  sh.getRange('A2:A').setNumberFormat('d/m/yyyy hh:mm');
  sh.getRange('E2:G').setNumberFormat('#,##0');
  sh.getRange('J2:K').setNumberFormat('#,##0');
  sh.getRange('L2:L').setNumberFormat('d/m/yyyy');
  [130, 80, 90, 230, 70, 100, 90, 230, 90, 115, 90, 90].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([TYPE.SALE, TYPE.IN, TYPE.ADJUST], true).setAllowInvalid(false).build());
  sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActive().getSheetByName(SHEET.PRODUCTS).getRange('A2:A'), true)
    .setAllowInvalid(false).setHelpText('เลือกรหัสจากชีตสินค้า').build());
  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberNotEqualTo(0)
    .setAllowInvalid(false).setHelpText('จำนวนต้องไม่เป็น 0 (ปรับยอดลดใช้ค่าติดลบ)').build());
  sh.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false).build());

  var type = sh.getRange('B2:B');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(TYPE.SALE).setFontColor('#1a73e8').setRanges([type]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(TYPE.IN).setFontColor(GREEN).setBold(true).setRanges([type]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(TYPE.ADJUST).setFontColor('#b06000').setBold(true).setRanges([type]).build(),
  ]);
  ['D1:D', 'G1:G', 'J1:L'].forEach(function (a1) {
    sh.getRange(a1).protect().setDescription('คำนวณอัตโนมัติ (ไม่ต้องแก้)').setWarningOnly(true);
  });
  sh.getRange(1, 1, rows.length + 1, 12).createFilter();
}

// ---------- รายงานรายวัน ----------

function buildDaily(sh) {
  var tx = "'" + SHEET.TX + "'!";
  header(sh, ['วันที่', 'ยอดขาย', 'ต้นทุนขาย', 'กำไรขั้นต้น', 'รายการขาย', 'กำไร %']);
  trimSheet(sh, 40, 20);   // room on the right for the chart
  sh.getRange('A2').setFormula('=SEQUENCE(30, 1, REF_DATE - 29, 1)');
  sh.getRange('B2').setFormula('=MAP(A2:A31, LAMBDA(d, SUMIFS(' + tx + 'G2:G, ' + tx + 'B2:B, "ขาย", ' + tx + 'L2:L, d)))');
  sh.getRange('C2').setFormula('=MAP(A2:A31, LAMBDA(d, SUMIFS(' + tx + 'K2:K, ' + tx + 'L2:L, d)))');
  sh.getRange('D2').setFormula('=ARRAYFORMULA(B2:B31 - C2:C31)');
  sh.getRange('E2').setFormula('=MAP(A2:A31, LAMBDA(d, COUNTIFS(' + tx + 'B2:B, "ขาย", ' + tx + 'L2:L, d)))');
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(B2:B31 = 0,, D2:D31 / B2:B31))');
  sh.getRange('A32:F32').setValues([['รวม 30 วัน', '=SUM(B2:B31)', '=SUM(C2:C31)', '=SUM(D2:D31)', '=SUM(E2:E31)', '=IF(B32 = 0,, D32 / B32)']])
    .setFontWeight('bold').setBackground(GREEN_SOFT);
  sh.getRange('A2:A31').setNumberFormat('d/m');
  sh.getRange('B2:E32').setNumberFormat('#,##0');
  sh.getRange('F2:F32').setNumberFormat('0%');
  [110, 110, 110, 110, 90, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A1:F32').protect().setDescription('สร้างจากชีตรายการอัตโนมัติ').setWarningOnly(true);

  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('A1:B31')).addRange(sh.getRange('D1:D31'))
    .setOption('title', 'ยอดขายและกำไรขั้นต้นรายวัน (บาท)')
    .setOption('colors', [GREEN, '#fbbc04'])
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { format: 'd/M', slantedText: true })
    .setOption('vAxis', { format: '#,##0' })
    .setPosition(2, 8, 0, 0).setOption('width', 760).setOption('height', 380).build());
}

// ---------- แดชบอร์ด ----------

function buildDashboard(ss) {
  var sh = ss.getSheetByName(SHEET.DASH);
  var pr = "'" + SHEET.PRODUCTS + "'!";
  var tx = "'" + SHEET.TX + "'!";
  trimSheet(sh, 50, 14);
  sh.setHiddenGridlines(true);
  // A gap | B C card | D gap | E F card | G gap | H I card | J gap | K L card | M N edge
  [24, 190, 80, 16, 150, 110, 16, 190, 80, 16, 150, 110, 24, 24].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  sh.getRange('B2').setValue('📦 ระบบสต็อกสินค้า · ' + DEMO_SHOP).setFontSize(20).setFontWeight('bold').setFontColor(INK);
  sh.getRange('B3').setFormula('="ข้อมูลถึงวันที่ " & TEXT(REF_DATE, "d/m/yyyy") & " · ตัวเลขทั้งหมดคำนวณจากชีตรายการอัตโนมัติ"').setFontColor(MUTE);

  var cards = [
    ['B5:C5', 'B6:C6', 'ยอดขาย 30 วัน (บาท)', '=SUMIFS(' + tx + 'G2:G, ' + tx + 'B2:B, "ขาย", ' + tx + 'L2:L, ">" & (REF_DATE - 30))', '#,##0'],
    ['E5:F5', 'E6:F6', 'กำไรขั้นต้น 30 วัน (บาท)', '=B6 - SUMIFS(' + tx + 'K2:K, ' + tx + 'L2:L, ">" & (REF_DATE - 30))', '#,##0'],
    ['H5:I5', 'H6:I6', 'มูลค่าสต็อก (ราคาทุน)', '=SUM(' + pr + 'I2:I)', '#,##0'],
    ['K5:L5', 'K6:L6', 'สินค้าใกล้หมด / หมด', '=COUNTIF(' + pr + 'M2:M, "ใกล้หมด") + COUNTIF(' + pr + 'M2:M, "หมด")', '0" รายการ"'],
  ];
  cards.forEach(function (c) {
    sh.getRange(c[0]).merge().setValue(c[2]).setFontColor(MUTE).setFontSize(10);
    sh.getRange(c[1]).merge().setFormula(c[3]).setNumberFormat(c[4]).setFontSize(20).setFontWeight('bold').setFontColor(INK);
    var box = sh.getRange(c[0].charAt(0) + '5:' + c[1].split(':')[1]);
    box.setBackground(GREEN_SOFT).setBorder(true, true, true, true, false, false, '#ceead6', SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange('K6').setFontColor('#c5221f');
  sh.setRowHeight(5, 26);
  sh.setRowHeight(6, 44);

  // Daily sales chart across the full width, fed by รายงานรายวัน.
  var daily = ss.getSheetByName(SHEET.DAILY);
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(daily.getRange('A1:B31')).addRange(daily.getRange('D1:D31'))
    .setNumHeaders(1)
    .setOption('title', 'ยอดขายและกำไรขั้นต้นรายวัน 30 วันล่าสุด (บาท)')
    .setOption('colors', [GREEN, '#fbbc04'])
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { slantedText: true })
    .setOption('vAxis', { format: '#,##0' })
    .setPosition(8, 2, 0, 0).setOption('width', 1100).setOption('height', 330).build());

  // Tables below the chart. SORTN / FILTER results are split with INDEX so
  // each column lands in the right (non-adjacent) sheet column.
  var top = 'SORTN({' + pr + 'B2:B, ' + pr + 'J2:J, ' + pr + 'K2:K}, 5, 0, 3, FALSE)';
  sh.getRange('B26').setValue('สินค้าขายดี 30 วัน').setFontWeight('bold').setFontSize(13);
  sh.getRange('B27:F27').setValues([['สินค้า', 'ชิ้น', '', 'ยอดขาย (บาท)', '']]).setFontWeight('bold').setBackground(GREEN).setFontColor('#ffffff');
  sh.getRange('B28').setFormula('=INDEX(' + top + ', 0, 1)');
  sh.getRange('C28').setFormula('=INDEX(' + top + ', 0, 2)');
  sh.getRange('E28').setFormula('=INDEX(' + top + ', 0, 3)');
  for (var r = 28; r <= 32; r++) {
    sh.getRange('F' + r).setFormula('=IF(E' + r + ' = "",, SPARKLINE(E' + r + ', {"charttype", "bar"; "max", MAX($E$28:$E$32); "color1", "' + GREEN + '"}))');
  }
  sh.getRange('C28:E32').setNumberFormat('#,##0');

  var low = 'SORT(FILTER({' + pr + 'B2:B, ' + pr + 'H2:H, ' + pr + 'G2:G, ' + pr + 'M2:M}, ' + pr + 'M2:M <> "ปกติ", ' + pr + 'A2:A <> ""), 2, TRUE)';
  sh.getRange('H26').setValue('ต้องสั่งซื้อเพิ่ม').setFontWeight('bold').setFontSize(13);
  sh.getRange('H27:L27').setValues([['สินค้า', 'คงเหลือ', '', 'จุดสั่งซื้อ', 'สถานะ']]).setFontWeight('bold').setBackground('#c5221f').setFontColor('#ffffff');
  sh.getRange('H28').setFormula('=IFERROR(INDEX(' + low + ', 0, 1), "ไม่มีสินค้าใกล้หมด 👍")');
  sh.getRange('I28').setFormula('=IFERROR(INDEX(' + low + ', 0, 2),)');
  sh.getRange('K28').setFormula('=IFERROR(INDEX(' + low + ', 0, 3),)');
  sh.getRange('L28').setFormula('=IFERROR(INDEX(' + low + ', 0, 4),)');
  sh.getRange('I28:K45').setNumberFormat('#,##0');
  sh.getRange('C27:C32').setHorizontalAlignment('right');
  sh.getRange('I27:L45').setHorizontalAlignment('right');
  sh.getRange('L27:L45').setHorizontalAlignment('center');
  var flag = sh.getRange('L28:L45');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('หมด').setBackground('#fce8e6').setFontColor('#c5221f').setBold(true).setRanges([flag]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ใกล้หมด').setBackground('#fef7e0').setFontColor('#b06000').setBold(true).setRanges([flag]).build(),
  ]);
  sh.getRange('B27:F32').setBorder(null, null, true, null, null, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('H27:L36').setBorder(null, null, true, null, null, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A1:N50').protect().setDescription('แดชบอร์ดคำนวณอัตโนมัติ').setWarningOnly(true);
}

// ---------- วิธีใช้ ----------

function buildHelp(sh) {
  trimSheet(sh, 40, 3);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 760);
  var lines = [
    ['วิธีใช้ระบบสต็อก', 'title'],
    ['ไฟล์นี้เป็นตัวอย่างระบบสต็อกบน Google Sheets + Apps Script ข้อมูลร้านและรายการทั้งหมดเป็นข้อมูลสมมติ', 'mute'],
    ['', ''],
    ['เริ่มใช้งาน', 'h'],
    ['1. ไฟล์ › ทำสำเนา (File › Make a copy) เพื่อได้ไฟล์ของตัวเองพร้อมสคริปต์', ''],
    ['2. เปิดไฟล์สำเนา รอเมนู 📦 ระบบสต็อก ขึ้นที่แถบเมนู (ครั้งแรกระบบจะขอสิทธิ์ใช้งานสคริปต์ในไฟล์ของคุณเอง)', ''],
    ['3. ใส่สินค้าของร้านในชีต "สินค้า" (รหัส ชื่อ หน่วย ต้นทุน ราคาขาย จุดสั่งซื้อ) ส่วนคอลัมน์สีเขียวเข้มคำนวณเอง', ''],
    ['', ''],
    ['ใช้งานประจำวัน', 'h'],
    ['• เมนู ➕ บันทึกรับเข้า / ขาย: ฟอร์มด้านข้าง เลือกสินค้า ใส่จำนวน ระบบเช็คให้ว่าสต็อกพอก่อนบันทึก', ''],
    ['• หรือพิมพ์ต่อท้ายชีต "รายการ" ได้เลย (เลือกประเภทและรหัสจาก dropdown วันที่เวลาจะใส่ให้เอง)', ''],
    ['• แดชบอร์ดอัปเดตทันที: ยอดขาย กำไร มูลค่าสต็อก สินค้าขายดี และรายการที่ต้องสั่งซื้อ', ''],
    ['', ''],
    ['แจ้งเตือนอัตโนมัติ', 'h'],
    ['• เมนู ⚙️ ตั้งค่าการแจ้งเตือน: ใส่ Telegram Bot, LINE Messaging API หรืออีเมล (token เก็บในระบบของไฟล์ ไม่โชว์ในชีต)', ''],
    ['• เมนู ⏰ เปิด / ปิด สรุปอัตโนมัติทุกวัน: ส่งสรุปยอดขายและสินค้าใกล้หมดทุกเย็นตามเวลาที่ตั้ง', ''],
    ['• เมนู 🔔 ตรวจสินค้าใกล้หมด: ดูรายการที่ถึงจุดสั่งซื้อ และส่งแจ้งเตือนได้ทันที', ''],
    ['', ''],
    ['ซอร์สโค้ด: github.com/chasvitf/stocksheet · ทำโดย Chanavee (Fastwork @chasvitf)', 'mute'],
  ];
  lines.forEach(function (l, i) {
    var cell = sh.getRange(i + 2, 2).setValue(l[0]).setWrap(true).setVerticalAlignment('middle');
    if (l[1] === 'title') cell.setFontSize(20).setFontWeight('bold');
    if (l[1] === 'h') cell.setFontSize(13).setFontWeight('bold').setFontColor(GREEN);
    if (l[1] === 'mute') cell.setFontColor(MUTE);
  });
  sh.setRowHeights(2, lines.length, 28);
}
