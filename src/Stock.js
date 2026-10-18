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
