// Tests for the pure logic in src/Stock.js and the demo data in src/Demo.js.
// Run: npm test
const test = require('node:test');
const assert = require('node:assert/strict');
const S = require('../src/Stock.js');
const D = require('../src/Demo.js');

const key = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

test('sales take stock out, receipts and adjustments keep their sign', () => {
  assert.equal(S.stockChange(S.TYPE.SALE, 3), -3);
  assert.equal(S.stockChange(S.TYPE.IN, 5), 5);
  assert.equal(S.stockChange(S.TYPE.ADJUST, -1), -1);
  const stock = S.stockByCode([
    { code: 'A', type: S.TYPE.IN, qty: 10 },
    { code: 'A', type: S.TYPE.SALE, qty: 4 },
    { code: 'A', type: S.TYPE.ADJUST, qty: -1 },
    { code: 'B', type: S.TYPE.IN, qty: 2 },
  ]);
  assert.deepEqual(stock, { A: 5, B: 2 });
});

test('low stock: out of stock first, then lowest relative to reorder point', () => {
  const items = S.lowStockItems([
    { code: 'A', stock: 5, reorder: 10 },
    { code: 'B', stock: 0, reorder: 3 },
    { code: 'C', stock: 9, reorder: 5 },
    { code: 'D', stock: 2, reorder: 10 },
    { code: 'E', stock: 1, reorder: 0 },
  ]);
  assert.deepEqual(items.map((p) => p.code), ['B', 'D', 'A']);
});

test('low stock message lists each item, or says all is fine', () => {
  const msg = S.lowStockMessage('ร้าน', [{ code: 'B', name: 'หูฟัง', unit: 'ชิ้น', stock: 0, reorder: 3 }]);
  assert.match(msg, /ใกล้หมด 1 รายการ/);
  assert.match(msg, /หูฟัง \(B\) หมด/);
  assert.match(S.lowStockMessage('ร้าน', []), /ปกติ/);
});

test('entry validation', () => {
  const p = { code: 'A', name: 'สาย', stock: 2 };
  assert.deepEqual(S.validateEntry({ type: S.TYPE.SALE, code: 'A', qty: 2, price: 99 }, p), []);
  assert.match(S.validateEntry({ type: S.TYPE.SALE, code: 'A', qty: 3, price: 99 }, p).join(), /สต็อกไม่พอ/);
  assert.match(S.validateEntry({ type: S.TYPE.SALE, code: 'X', qty: 1, price: 99 }, null).join(), /ไม่พบรหัสสินค้า/);
  assert.match(S.validateEntry({ type: S.TYPE.IN, code: 'A', qty: -1, price: 1 }, p).join(), /มากกว่า 0/);
  assert.match(S.validateEntry({ type: S.TYPE.IN, code: 'A', qty: 1.5, price: 1 }, p).join(), /จำนวนเต็ม/);
  assert.match(S.validateEntry({ type: S.TYPE.IN, code: 'A', qty: 1, price: -5 }, p).join(), /ราคา/);
  assert.match(S.validateEntry({ type: 'ยืม', code: 'A', qty: 1, price: 1 }, p).join(), /ประเภท/);
  assert.deepEqual(S.validateEntry({ type: S.TYPE.ADJUST, code: 'A', qty: -2, price: 0 }, p), []);
  assert.match(S.validateEntry({ type: S.TYPE.ADJUST, code: 'A', qty: -3, price: 0 }, p).join(), /สต็อกไม่พอ/);
});

test('day summary adds up sales, cost and best sellers for that day only', () => {
  const day = new Date(2026, 9, 18, 10);
  const other = new Date(2026, 9, 17, 10);
  const mv = [
    { date: day, type: S.TYPE.SALE, code: 'A', qty: 2, price: 100, cost: 40 },
    { date: day, type: S.TYPE.SALE, code: 'B', qty: 1, price: 500, cost: 300 },
    { date: day, type: S.TYPE.IN, code: 'A', qty: 50, price: 40, cost: 40 },
    { date: other, type: S.TYPE.SALE, code: 'A', qty: 9, price: 100, cost: 40 },
  ];
  const s = S.daySummary(mv, '2026-10-18', key);
  assert.deepEqual([s.sales, s.cost, s.profit, s.lines], [700, 380, 320, 2]);
  assert.deepEqual(s.top.map((t) => t.code), ['B', 'A']);
});

test('formatBaht and the daily message', () => {
  assert.equal(S.formatBaht(1234567.4), '1,234,567');
  assert.equal(S.formatBaht(-1500), '-1,500');
  assert.equal(S.formatBaht(0), '0');
  const msg = S.dailyMessage('ร้าน', { sales: 7000, profit: 3500, lines: 12, top: [{ code: 'A', sales: 4000 }] }, '18/10/2026', { A: 'สาย' }, []);
  assert.match(msg, /ยอดขาย 7,000 บาท \(12 รายการ\)/);
  assert.match(msg, /50%/);
  assert.match(msg, /ขายดี: สาย 4,000/);
});

test('demo data is deterministic, never goes below zero, and leaves a few items low', () => {
  const products = D.demoProducts();
  const end = new Date(2026, 9, 18);
  const a = S.generateDemoMovements(products, end, D.DEMO_DAYS, D.DEMO_SEED);
  const b = S.generateDemoMovements(products, end, D.DEMO_DAYS, D.DEMO_SEED);
  assert.deepEqual(a, b);

  // Replay in order: stock must never be negative at any point.
  const stock = {};
  for (const m of a) {
    stock[m.code] = (stock[m.code] || 0) + S.stockChange(m.type, m.qty);
    assert.ok(stock[m.code] >= 0, `${m.code} went negative on ${m.date}`);
  }
  const days = new Set(a.map((m) => key(m.date)));
  assert.equal(days.size, D.DEMO_DAYS);
  assert.equal(key(a[0].date), '2026-08-20');
  assert.equal(key(a[a.length - 1].date), '2026-10-18');

  const low = S.lowStockItems(products.map((p) => ({ ...p, stock: stock[p.code] })));
  assert.ok(low.length >= 2 && low.length <= 6, `expected a few low items, got ${low.length}`);
  assert.ok(low.some((p) => p.stock === 0), 'expected one item out of stock');
});

test('every demo product has sane prices', () => {
  const codes = new Set();
  for (const p of D.demoProducts()) {
    assert.ok(!codes.has(p.code), `duplicate ${p.code}`);
    codes.add(p.code);
    assert.match(p.code, /^[A-Z]{2}-\d{3}$/);
    assert.ok(p.price > p.cost, `${p.code} sells below cost`);
    assert.ok(p.restockQty > p.reorder, `${p.code} restock does not clear the reorder point`);
  }
});
