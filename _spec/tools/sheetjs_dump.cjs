// Dump công thức SheetJS (xlsx@0.18.5) cho bộ kiểm verify_formulas.py.
// Dùng: node sheetjs_dump.cjs <thư mục nguồn> <tệp JSON ra>   (XLFORM_DIR mặc định /tmp/xlform)
const path = require('path'), fs = require('fs');
const XLSX = require(path.join(process.env.XLFORM_DIR || '/tmp/xlform', 'node_modules', 'xlsx'));
const src = process.argv[2], out = process.argv[3];
const res = {};
for (const name of fs.readdirSync(src).sort()) {
  const p = path.join(src, name);
  if (!fs.statSync(p).isFile()) continue;
  const b = fs.readFileSync(p);
  const isOle = b.length > 4 && b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0;
  const isZip = b.length > 2 && b[0] === 0x50 && b[1] === 0x4b;
  if (!isOle && !isZip) continue;
  let wb;
  try { wb = XLSX.read(b, { type: 'buffer', cellFormula: true }); } catch (e) { res[name] = { __error: String(e.message || e) }; continue; }
  const o = {};
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn], m = {};
    for (const ref of Object.keys(ws)) {
      if (ref[0] === '!') continue;
      const c = ws[ref];
      const f = c && (c.f || c.F);
      if (f) m[ref] = String(f);
    }
    o[sn] = m;
  }
  res[name] = o;
}
fs.writeFileSync(out, JSON.stringify(res));
