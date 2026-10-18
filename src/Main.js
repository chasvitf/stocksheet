/** @OnlyCurrentDoc */

var SHEET = { DASH: 'แดชบอร์ด', PRODUCTS: 'สินค้า', TX: 'รายการ', DAILY: 'รายงานรายวัน', HELP: 'วิธีใช้' };

// Column numbers (1-based). Formula columns are filled by ARRAYFORMULA in row 1
// and must never be written by the script.
var P = { CODE: 1, NAME: 2, CATEGORY: 3, UNIT: 4, COST: 5, PRICE: 6, REORDER: 7, STOCK: 8 };
var T = { DATE: 1, TYPE: 2, CODE: 3, NAME: 4, QTY: 5, PRICE: 6, AMOUNT: 7, NOTE: 8, BY: 9 };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 ระบบสต็อก')
    .addItem('➕ บันทึกรับเข้า / ขาย', 'showEntryForm')
    .addItem('🔔 ตรวจสินค้าใกล้หมด', 'checkLowStock')
    .addItem('📊 สรุปยอดขายวันล่าสุด', 'showDaySummary')
    .addSeparator()
    .addItem('⚙️ ตั้งค่าการแจ้งเตือน', 'showSettings')
    .addItem('⏰ เปิด / ปิด สรุปอัตโนมัติทุกวัน', 'toggleDailyTrigger')
    .addSeparator()
    .addItem('🧪 สร้างข้อมูลตัวอย่างใหม่', 'resetDemoData')
    .addToUi();
}

/** Typing a row by hand in รายการ: stamp the date and time once the product code is filled in. */
function onEdit(e) {
  var range = e.range;
  if (range.getSheet().getName() !== SHEET.TX || range.getRow() < 2 || range.getColumn() !== T.CODE) return;
  var dateCell = range.getSheet().getRange(range.getRow(), T.DATE);
  if (range.getValue() && !dateCell.getValue()) dateCell.setValue(new Date());
}

// ---------- reading the sheet ----------

function sheet(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '" กรุณารันเมนู 🧪 สร้างข้อมูลตัวอย่างใหม่ หรือฟังก์ชัน setup ก่อน');
  return sh;
}

/** Rows of a sheet up to the last filled cell in column A (formula columns may run further). */
function filledRows(sh, width) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var rows = sh.getRange(2, 1, n, width).getValues();
  var last = rows.length;
  while (last > 0 && rows[last - 1][0] === '') last--;
  return rows.slice(0, last);
}

function readProducts() {
  return filledRows(sheet(SHEET.PRODUCTS), P.STOCK).map(function (r) {
    return { code: String(r[0]), name: r[1], category: r[2], unit: r[3], cost: Number(r[4]) || 0,
      price: Number(r[5]) || 0, reorder: Number(r[6]) || 0, stock: Number(r[7]) || 0 };
  });
}

function readMovements(products) {
  var cost = {};
  products.forEach(function (p) { cost[p.code] = p.cost; });
  return filledRows(sheet(SHEET.TX), T.PRICE)
    .filter(function (r) { return r[0] instanceof Date; })
    .map(function (r) {
      return { date: r[0], type: r[1], code: String(r[2]), qty: Number(r[4]) || 0, price: Number(r[5]) || 0, cost: cost[r[2]] || 0 };
    });
}

function tz() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function dayKey(date) {
  return Utilities.formatDate(date, tz(), 'yyyy-MM-dd');
}

// ---------- entry form ----------

function showEntryForm() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('บันทึกรับเข้า / ขาย');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Data for the sidebar: products with their current stock, and the last name used. */
function getFormData() {
  return {
    products: readProducts(),
    by: PropertiesService.getUserProperties().getProperty('recordedBy') || '',
  };
}

/** Called from Sidebar.html. Validates against the current stock, then writes one row. */
function addEntry(form) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    var entry = { type: form.type, code: String(form.code || '').trim().toUpperCase(),
      qty: Number(form.qty), price: Number(form.price) };
    var product = readProducts().filter(function (p) { return p.code === entry.code; })[0] || null;
    var errors = validateEntry(entry, product);
    if (errors.length) throw new Error(errors.join('\n'));

    var by = String(form.by || '').trim().slice(0, 40);
    PropertiesService.getUserProperties().setProperty('recordedBy', by);
    var sh = sheet(SHEET.TX);
    var row = filledRows(sh, 1).length + 2;
    sh.getRange(row, T.DATE, 1, 3).setValues([[new Date(), entry.type, entry.code]]);
    sh.getRange(row, T.QTY, 1, 2).setValues([[entry.qty, entry.price]]);
    sh.getRange(row, T.NOTE, 1, 2).setValues([[String(form.note || '').slice(0, 200), by]]);
    SpreadsheetApp.flush();

    var after = product.stock + stockChange(entry.type, entry.qty);
    var low = product.reorder > 0 && after <= product.reorder;
    return { row: row, name: product.name, unit: product.unit, stock: after, low: low };
  } finally {
    lock.releaseLock();
  }
}

// ---------- alerts and summaries ----------

function checkLowStock() {
  var ui = SpreadsheetApp.getUi();
  var s = getSettings();
  var items = lowStockItems(readProducts());
  var text = lowStockMessage(s.shopName, items);
  if (!items.length) {
    ui.alert('ตรวจสินค้าใกล้หมด', text, ui.ButtonSet.OK);
    return;
  }
  var answer = ui.alert('ตรวจสินค้าใกล้หมด', text + '\n\nส่งรายการนี้ไปยังช่องทางแจ้งเตือนที่ตั้งไว้ไหม?', ui.ButtonSet.YES_NO);
  if (answer === ui.Button.YES) ui.alert('ผลการส่ง', describeResults(notifyAll('สินค้าใกล้หมด ' + s.shopName, text)), ui.ButtonSet.OK);
}

/** Summary text for one day; the latest day with data when `date` is not given. */
function buildDailyMessage(date) {
  var s = getSettings();
  var products = readProducts();
  var movements = readMovements(products);
  if (!date) {
    date = movements.reduce(function (latest, m) { return m.date > latest ? m.date : latest; }, new Date(0));
  }
  var names = {};
  products.forEach(function (p) { names[p.code] = p.name; });
  var summary = daySummary(movements, dayKey(date), dayKey);
  var label = Utilities.formatDate(date, tz(), 'd/M/yyyy');
  return dailyMessage(s.shopName, summary, label, names, lowStockItems(products));
}

function showDaySummary() {
  var ui = SpreadsheetApp.getUi();
  var text = buildDailyMessage();
  var answer = ui.alert('สรุปยอดขายวันล่าสุด', text + '\n\nส่งสรุปนี้ไปยังช่องทางแจ้งเตือนที่ตั้งไว้ไหม?', ui.ButtonSet.YES_NO);
  if (answer === ui.Button.YES) ui.alert('ผลการส่ง', describeResults(notifyAll('สรุปยอดขาย ' + getSettings().shopName, text)), ui.ButtonSet.OK);
}

/** Time-driven trigger: today's summary plus the low-stock list, every evening. */
function dailyJob() {
  var results = notifyAll('สรุปยอดขาย ' + getSettings().shopName, buildDailyMessage(new Date()));
  console.log(describeResults(results));
}

function findDailyTrigger() {
  return ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'dailyJob'; })[0] || null;
}

function installDailyTrigger(hour) {
  var old = findDailyTrigger();
  if (old) ScriptApp.deleteTrigger(old);
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(hour).inTimezone(tz()).create();
}

function toggleDailyTrigger() {
  var ui = SpreadsheetApp.getUi();
  var old = findDailyTrigger();
  if (old) {
    ScriptApp.deleteTrigger(old);
    ui.alert('สรุปอัตโนมัติ', 'ปิดการส่งสรุปอัตโนมัติแล้ว', ui.ButtonSet.OK);
    return;
  }
  var hour = Number(getSettings().reportHour);
  installDailyTrigger(hour);
  var note = notifyChannelsCount() ? '' : '\nยังไม่ได้ตั้งค่าช่องทางแจ้งเตือน ตั้งได้ที่เมนู ⚙️ ตั้งค่าการแจ้งเตือน';
  ui.alert('สรุปอัตโนมัติ', 'เปิดแล้ว ระบบจะส่งสรุปยอดขายและสินค้าใกล้หมดทุกวัน เวลาประมาณ ' + hour + ':00 น.' + note, ui.ButtonSet.OK);
}

function notifyChannelsCount() {
  var s = getSettings();
  return [s.telegramToken && s.telegramChatId, s.lineToken && s.lineTo, s.email].filter(Boolean).length;
}

function showSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Settings').setWidth(460).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'ตั้งค่าการแจ้งเตือน');
}
