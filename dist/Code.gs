// ===== Main.js =====
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

/**
 * Called from Sidebar.html. Validates against the current stock, then writes
 * one row. A rejected entry (e.g. not enough stock) is a normal answer, not an
 * error, so it comes back as {ok: false} and the execution log stays clean.
 */
function addEntry(form) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    var entry = { type: form.type, code: String(form.code || '').trim().toUpperCase(),
      qty: Number(form.qty), price: Number(form.price) };
    var product = readProducts().filter(function (p) { return p.code === entry.code; })[0] || null;
    var errors = validateEntry(entry, product);
    if (errors.length) return { ok: false, error: errors.join('\n') };

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
    return { ok: true, row: row, name: product.name, unit: product.unit, stock: after, low: low };
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

// ===== Stock.js =====
/**
 * Pure stock logic: no SpreadsheetApp calls here, so test/stock.test.js can run
 * it in Node. Everything that touches the sheet lives in Main.js / Setup.js.
 */

var TYPE = { IN: 'รับเข้า', SALE: 'ขาย', ADJUST: 'ปรับยอด' };

/** Signed change in stock for one movement: sales go out, the rest come in (adjustments may be negative). */
function stockChange(type, qty) {
  return type === TYPE.SALE ? -Math.abs(qty) : qty;
}

/** Current stock per product code from a list of movements [{code, type, qty}]. */
function stockByCode(movements) {
  var stock = {};
  movements.forEach(function (m) {
    stock[m.code] = (stock[m.code] || 0) + stockChange(m.type, m.qty);
  });
  return stock;
}

/**
 * Products at or below their reorder point, most urgent first (out of stock,
 * then the lowest stock relative to the reorder point).
 * products: [{code, name, unit, stock, reorder}]
 */
function lowStockItems(products) {
  return products
    .filter(function (p) { return p.code && p.reorder > 0 && p.stock <= p.reorder; })
    .sort(function (a, b) { return a.stock / a.reorder - b.stock / b.reorder || a.code.localeCompare(b.code); });
}

/** Text for a low-stock alert (Telegram / LINE / email). */
function lowStockMessage(shopName, items) {
  if (!items.length) return '✅ ' + shopName + ': สต็อกทุกรายการอยู่ในระดับปกติ';
  var lines = items.map(function (p) {
    var state = p.stock <= 0 ? 'หมด' : 'เหลือ ' + p.stock + ' ' + p.unit;
    return '• ' + p.name + ' (' + p.code + ') ' + state + ' / จุดสั่งซื้อ ' + p.reorder;
  });
  return '⚠️ ' + shopName + ': สินค้าใกล้หมด ' + items.length + ' รายการ\n' + lines.join('\n');
}

/**
 * Check an entry from the form before it is written.
 * entry: {type, code, qty, price}; product: {code, stock} or null.
 * Returns a list of error messages (empty when the entry is fine).
 */
function validateEntry(entry, product) {
  var errors = [];
  if (!product) errors.push('ไม่พบรหัสสินค้า ' + (entry.code || ''));
  if ([TYPE.IN, TYPE.SALE, TYPE.ADJUST].indexOf(entry.type) < 0) errors.push('ประเภทรายการไม่ถูกต้อง');
  if (!isFinite(entry.qty) || entry.qty === 0) errors.push('จำนวนต้องเป็นตัวเลขและไม่เป็น 0');
  else if (entry.type !== TYPE.ADJUST && entry.qty < 0) errors.push('จำนวนต้องมากกว่า 0');
  else if (Math.round(entry.qty) !== entry.qty) errors.push('จำนวนต้องเป็นจำนวนเต็ม');
  if (!isFinite(entry.price) || entry.price < 0) errors.push('ราคาต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป');
  if (product && !errors.length) {
    var after = product.stock + stockChange(entry.type, entry.qty);
    if (after < 0) errors.push('สต็อกไม่พอ: ' + (product.name || product.code) + ' เหลือ ' + product.stock);
  }
  return errors;
}

/**
 * Totals for one calendar day. movements: [{date: Date, type, code, qty, price, cost}]
 * where price is the unit price and cost the unit cost. dayKey(date) -> 'yyyy-mm-dd'
 * is passed in so the caller decides the time zone.
 */
function daySummary(movements, day, dayKey) {
  var sales = 0, cost = 0, lines = 0, byCode = {};
  movements.forEach(function (m) {
    if (m.type !== TYPE.SALE || dayKey(m.date) !== day) return;
    sales += m.qty * m.price;
    cost += m.qty * m.cost;
    lines += 1;
    byCode[m.code] = (byCode[m.code] || 0) + m.qty * m.price;
  });
  var top = Object.keys(byCode)
    .map(function (code) { return { code: code, sales: byCode[code] }; })
    .sort(function (a, b) { return b.sales - a.sales; })
    .slice(0, 3);
  return { day: day, sales: sales, cost: cost, profit: sales - cost, lines: lines, top: top };
}

/** 12345.6 -> '12,346' (Apps Script has no Thai locale data for toLocaleString). */
function formatBaht(n) {
  var s = String(Math.round(Math.abs(n))).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-' : '') + s;
}

/** Text for the end-of-day message: totals, best sellers and the low-stock list. */
function dailyMessage(shopName, summary, dayLabel, names, lowItems) {
  var margin = summary.sales ? Math.round(summary.profit / summary.sales * 100) : 0;
  var lines = [
    '📊 สรุปยอดขาย ' + shopName + ' วันที่ ' + dayLabel,
    'ยอดขาย ' + formatBaht(summary.sales) + ' บาท (' + summary.lines + ' รายการ)',
    'กำไรขั้นต้น ' + formatBaht(summary.profit) + ' บาท (' + margin + '%)',
  ];
  if (summary.top.length) {
    lines.push('ขายดี: ' + summary.top.map(function (t) { return (names[t.code] || t.code) + ' ' + formatBaht(t.sales); }).join(', '));
  }
  lines.push('');
  lines.push(lowStockMessage(shopName, lowItems));
  return lines.join('\n');
}

/** Small seeded PRNG (mulberry32) so demo data is the same on every run. */
function seededRandom(seed) {
  var t = seed >>> 0;
  return function () {
    t += 0x6D2B79F5;
    var r = Math.imul(t ^ (t >>> 15), 1 | t);
    r ^= r + Math.imul(r ^ (r >>> 7), 61 | r);
    return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
  };
}

/**
 * Demo movements for `days` days ending on `endDate`: an opening stock count,
 * daily sales weighted by product popularity, restocks when an item runs low,
 * and the odd adjustment for damaged goods. Stock never goes below zero, and a
 * few items are left low at the end so the alerts have something to show.
 * products: [{code, cost, price, reorder, popularity, restockQty, leaveLow}]
 */
function generateDemoMovements(products, endDate, days, seed) {
  var rand = seededRandom(seed);
  var stock = {};
  var out = [];
  var start = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate() - days + 1);
  var at = function (dayOffset, hour, minute) {
    return new Date(start.getFullYear(), start.getMonth(), start.getDate() + dayOffset, hour, minute);
  };
  var weightTotal = products.reduce(function (s, p) { return s + p.popularity; }, 0);
  var pick = function () {
    var r = rand() * weightTotal;
    for (var i = 0; i < products.length; i++) {
      r -= products[i].popularity;
      if (r <= 0) return products[i];
    }
    return products[products.length - 1];
  };

  products.forEach(function (p, i) {
    stock[p.code] = p.restockQty;
    out.push({ date: at(0, 8, i), type: TYPE.IN, code: p.code, qty: p.restockQty, price: p.cost, note: 'ยอดยกมา (นับสต็อกเปิดร้าน)' });
  });

  for (var d = 0; d < days; d++) {
    var weekend = at(d, 12, 0).getDay() % 6 === 0;
    var count = Math.round((weekend ? 16 : 10) + rand() * 8);
    var minutes = [];
    for (var k = 0; k < count; k++) minutes.push(Math.floor(9 * 60 + rand() * 11 * 60));
    minutes.sort(function (a, b) { return a - b; });

    minutes.forEach(function (min) {
      var p = pick();
      var qty = 1 + (rand() < 0.25 ? 1 : 0) + (rand() < 0.08 ? 2 : 0);
      if (stock[p.code] < qty) return;
      stock[p.code] -= qty;
      out.push({ date: at(d, Math.floor(min / 60), min % 60), type: TYPE.SALE, code: p.code, qty: qty, price: p.price, note: '' });
    });

    // Restock in the morning when an item is at or below its reorder point,
    // except in the last weeks for items meant to end low.
    products.forEach(function (p, i) {
      var lastWeeks = d >= days - 18;
      if (stock[p.code] > p.reorder || d === days - 1 || (p.leaveLow && lastWeeks)) return;
      stock[p.code] += p.restockQty;
      out.push({ date: at(d + 1, 8, 30 + i), type: TYPE.IN, code: p.code, qty: p.restockQty, price: p.cost, note: 'รับของจากซัพพลายเออร์' });
    });

    if (rand() < 0.06) {
      var damaged = pick();
      if (stock[damaged.code] > 0) {
        stock[damaged.code] -= 1;
        out.push({ date: at(d, 20, 30), type: TYPE.ADJUST, code: damaged.code, qty: -1, price: damaged.cost, note: 'สินค้าชำรุด ตัดออกจากสต็อก' });
      }
    }
  }

  out.sort(function (a, b) { return a.date - b.date; });
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = { TYPE: TYPE, stockChange: stockChange, stockByCode: stockByCode, lowStockItems: lowStockItems,
    lowStockMessage: lowStockMessage, validateEntry: validateEntry, daySummary: daySummary,
    formatBaht: formatBaht, dailyMessage: dailyMessage, seededRandom: seededRandom,
    generateDemoMovements: generateDemoMovements };
}

// ===== Demo.js =====
/**
 * Demo catalogue for a made-up phone and computer accessories shop. Prices are
 * typical Thai retail prices; popularity drives how often an item sells.
 */
var DEMO_SHOP = 'ร้านไอทีตัวอย่าง (ข้อมูลสมมติ)';
var DEMO_STAFF = ['ต้น', 'แพร', 'เจ'];
var DEMO_DAYS = 60;
var DEMO_SEED = 20261018;

var DEMO_PRODUCTS = [
  // code, name, category, unit, cost, price, reorder, popularity, restockQty, leaveLow
  ['CB-001', 'สายชาร์จ USB-C 1 ม.', 'สายชาร์จ', 'เส้น', 45, 99, 15, 10, 60, false],
  ['CB-002', 'สายชาร์จ Lightning 1 ม.', 'สายชาร์จ', 'เส้น', 55, 129, 12, 7, 48, false],
  ['CB-003', 'สาย USB-C to C 2 ม. 60W', 'สายชาร์จ', 'เส้น', 79, 179, 8, 4, 30, false],
  ['CB-004', 'สาย HDMI 1.5 ม.', 'สายชาร์จ', 'เส้น', 65, 159, 5, 2, 20, false],
  ['AD-001', 'หัวชาร์จ 20W USB-C', 'หัวชาร์จ', 'ชิ้น', 120, 290, 10, 6, 40, false],
  ['AD-002', 'หัวชาร์จ GaN 65W 3 พอร์ต', 'หัวชาร์จ', 'ชิ้น', 390, 790, 5, 3, 20, true],
  ['AD-003', 'หัวชาร์จในรถ 2 พอร์ต', 'หัวชาร์จ', 'ชิ้น', 95, 229, 6, 2, 20, false],
  ['EP-001', 'หูฟังมีสาย 3.5 มม.', 'หูฟัง', 'ชิ้น', 60, 159, 8, 3, 30, false],
  ['EP-002', 'หูฟังบลูทูธ TWS', 'หูฟัง', 'ชิ้น', 350, 790, 6, 4, 20, false],
  ['EP-003', 'หูฟังบลูทูธตัดเสียงรบกวน', 'หูฟัง', 'ชิ้น', 890, 1690, 3, 1.5, 10, true],
  ['PB-001', 'พาวเวอร์แบงก์ 10000 mAh', 'พาวเวอร์แบงก์', 'ชิ้น', 290, 590, 6, 4, 24, false],
  ['PB-002', 'พาวเวอร์แบงก์ 20000 mAh 22.5W', 'พาวเวอร์แบงก์', 'ชิ้น', 450, 890, 4, 2.5, 16, false],
  ['CS-001', 'เคสใส iPhone 15', 'เคสและฟิล์ม', 'ชิ้น', 35, 159, 10, 6, 50, false],
  ['CS-002', 'เคสใส iPhone 16', 'เคสและฟิล์ม', 'ชิ้น', 35, 159, 10, 8, 50, true],
  ['CS-003', 'เคส Samsung A55', 'เคสและฟิล์ม', 'ชิ้น', 40, 159, 6, 3, 30, false],
  ['FM-001', 'ฟิล์มกระจก iPhone 15/16', 'เคสและฟิล์ม', 'แผ่น', 25, 129, 15, 9, 80, false],
  ['FM-002', 'ฟิล์มกระจก Samsung A55', 'เคสและฟิล์ม', 'แผ่น', 25, 129, 8, 4, 40, false],
  ['FM-003', 'ฟิล์มกันมอง iPhone', 'เคสและฟิล์ม', 'แผ่น', 45, 199, 6, 3, 30, false],
  ['PC-001', 'เมาส์ไร้สาย', 'อุปกรณ์คอม', 'ชิ้น', 120, 290, 6, 3, 24, false],
  ['PC-002', 'แฟลชไดรฟ์ 64GB', 'อุปกรณ์คอม', 'ชิ้น', 110, 249, 8, 3, 30, false],
  ['PC-003', 'การ์ด microSD 128GB', 'อุปกรณ์คอม', 'ชิ้น', 230, 459, 6, 2.5, 24, true],
  ['PC-004', 'ฮับ USB-C 6 in 1', 'อุปกรณ์คอม', 'ชิ้น', 420, 890, 3, 1.2, 10, false],
  ['HL-001', 'ขาตั้งมือถือตั้งโต๊ะ', 'อื่น ๆ', 'ชิ้น', 60, 169, 6, 2, 24, false],
  ['HL-002', 'ไม้เซลฟี่พร้อมขาตั้ง', 'อื่น ๆ', 'ชิ้น', 140, 349, 4, 1.5, 12, false],
];

/** DEMO_PRODUCTS as objects for generateDemoMovements(). */
function demoProducts() {
  return DEMO_PRODUCTS.map(function (r) {
    return { code: r[0], name: r[1], category: r[2], unit: r[3], cost: r[4], price: r[5], reorder: r[6],
      popularity: r[7], restockQty: r[8], leaveLow: r[9] };
  });
}

if (typeof module !== 'undefined') {
  module.exports = { DEMO_SHOP: DEMO_SHOP, DEMO_STAFF: DEMO_STAFF, DEMO_DAYS: DEMO_DAYS, DEMO_SEED: DEMO_SEED,
    DEMO_PRODUCTS: DEMO_PRODUCTS, demoProducts: demoProducts };
}

// ===== Notify.js =====
/**
 * Alerts to Telegram, LINE (Messaging API push) and email. Tokens live in the
 * document's properties, not in a cell, so they never show up in the sheet
 * and are not copied when someone makes a copy of the file.
 */

var SETTING_KEYS = ['shopName', 'reportHour', 'email', 'telegramToken', 'telegramChatId', 'lineToken', 'lineTo'];

function getSettings() {
  var props = PropertiesService.getDocumentProperties().getProperties();
  var s = {};
  SETTING_KEYS.forEach(function (k) { s[k] = props[k] || ''; });
  if (!s.shopName) s.shopName = DEMO_SHOP;
  if (!s.reportHour) s.reportHour = '20';
  return s;
}

/** Called from Settings.html. Empty secret fields keep the saved value. Invalid input comes back as {error}. */
function saveSettings(form) {
  var props = PropertiesService.getDocumentProperties();
  var hour = Number(form.reportHour);
  if (!(Math.round(hour) === hour && hour >= 0 && hour <= 23)) return { error: 'เวลาส่งสรุปต้องเป็นชั่วโมง 0-23' };
  if (form.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(form.email).trim())) return { error: 'รูปแบบอีเมลไม่ถูกต้อง' };
  SETTING_KEYS.forEach(function (k) {
    var v = String(form[k] || '').trim();
    var secret = k === 'telegramToken' || k === 'lineToken';
    if (secret && !v) return;
    props.setProperty(k, v);
  });
  if (findDailyTrigger()) installDailyTrigger(hour);   // keep the trigger in sync with the hour
  return settingsForForm();
}

/** Settings for the form: secrets are masked, only whether they are set is sent. */
function settingsForForm() {
  var s = getSettings();
  return {
    shopName: s.shopName, reportHour: s.reportHour, email: s.email, telegramChatId: s.telegramChatId, lineTo: s.lineTo,
    telegramSet: !!s.telegramToken, lineSet: !!s.lineToken, dailyOn: !!findDailyTrigger(),
  };
}

function clearSecret(key) {
  if (key !== 'telegramToken' && key !== 'lineToken') throw new Error('unknown key');
  PropertiesService.getDocumentProperties().deleteProperty(key);
  return settingsForForm();
}

function sendTelegram(token, chatId, text) {
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', payload: { chat_id: chatId, text: text }, muteHttpExceptions: true,
  });
  var body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) throw new Error(body.description || 'HTTP ' + res.getResponseCode());
}

function sendLine(token, to, text) {
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text.slice(0, 5000) }] }),
  });
  if (res.getResponseCode() !== 200) {
    var body = JSON.parse(res.getContentText() || '{}');
    throw new Error(body.message || 'HTTP ' + res.getResponseCode());
  }
}

/** Send to every configured channel. Returns [{channel, ok, error}] (empty when none is set up). */
function notifyAll(subject, text) {
  var s = getSettings();
  var jobs = [];
  if (s.telegramToken && s.telegramChatId) jobs.push(['Telegram', function () { sendTelegram(s.telegramToken, s.telegramChatId, text); }]);
  if (s.lineToken && s.lineTo) jobs.push(['LINE', function () { sendLine(s.lineToken, s.lineTo, text); }]);
  if (s.email) jobs.push(['อีเมล', function () { MailApp.sendEmail(s.email, subject, text); }]);
  return jobs.map(function (job) {
    try {
      job[1]();
      return { channel: job[0], ok: true };
    } catch (err) {
      console.error(job[0] + ' failed: ' + err.message);
      return { channel: job[0], ok: false, error: err.message };
    }
  });
}

function describeResults(results) {
  if (!results.length) return 'ยังไม่ได้ตั้งค่าช่องทางแจ้งเตือน (เมนู 📦 ระบบสต็อก › ⚙️ ตั้งค่าการแจ้งเตือน)';
  return results.map(function (r) { return (r.ok ? '✅ ส่งทาง ' : '❌ ส่งทาง ') + r.channel + (r.ok ? ' แล้ว' : ' ไม่สำเร็จ: ' + r.error); }).join('\n');
}

/** "ทดสอบส่ง" button in Settings.html. */
function sendTestMessage() {
  var s = getSettings();
  return describeResults(notifyAll('ทดสอบการแจ้งเตือน', '🔔 ทดสอบการแจ้งเตือนจากระบบสต็อก ' + s.shopName));
}

// ===== Setup.js =====
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
