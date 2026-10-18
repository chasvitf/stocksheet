// Concatenate src/*.js into dist/Code.gs for pasting into the Apps Script
// editor (the editor runs every file in one global scope, so this is equivalent
// to pushing the files one by one with clasp). HTML files are copied as is.
const fs = require('fs');
const path = require('path');

const src = path.join(__dirname, '..', 'src');
const dist = path.join(__dirname, '..', 'dist');
const order = ['Main.js', 'Stock.js', 'Demo.js', 'Notify.js', 'Setup.js'];

fs.mkdirSync(dist, { recursive: true });
const code = order.map((f) => `// ===== ${f} =====\n` + fs.readFileSync(path.join(src, f), 'utf8')).join('\n');
fs.writeFileSync(path.join(dist, 'Code.gs'), code);
for (const f of fs.readdirSync(src).filter((f) => f.endsWith('.html'))) {
  fs.copyFileSync(path.join(src, f), path.join(dist, f));
}
console.log(`dist/Code.gs ${code.length} chars, ${code.split('\n').length} lines`);
