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
