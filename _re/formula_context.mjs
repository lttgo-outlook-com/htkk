#!/usr/bin/env node
/**
 * formula_context.mjs — compact, labeled, de-duplicated formula context for one form folder.
 *
 * For every spreadsheet in the folder (xls/xls1/xlsx, plus magic-byte detection),
 * groups cells by (file, sheet, exact formula string) and prints each distinct
 * formula once with: occurrence count, up to 10 refs, and the nearest text labels
 * (left / above) of the first ref — so a reader can name each rule without
 * re-opening the workbook.
 *
 * Usage: node _re/formula_context.mjs <form-folder>
 */
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';

const XLSX = createRequire(
  path.join(process.env.XLFORM_DIR || '/tmp/xlform', 'package.json'),
)('xlsx');

const dir = path.resolve(process.argv[2]);
if (!fs.statSync(dir).isDirectory()) {
  console.error(`not a directory: ${dir}`);
  process.exit(2);
}
const form = path.basename(dir);
const group = path.basename(path.dirname(dir));
const MAX_PATTERN_PER_SHEET = 400;
const MAX_REFS = 10;

function colNum(s) {
  return [...s].reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0);
}
function colStr(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - r) / 26;
  }
  return s;
}
function refKey(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return [0, 0];
  return [Number(m[2]), colNum(m[1])];
}
function magicKind(buf) {
  if (buf.length < 4) return null;
  if (buf[0] === 0x50 && buf[1] === 0x4b) return 'xlsx';
  if (buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0) return 'xls';
  return null;
}

const files = fs
  .readdirSync(dir)
  .filter((f) => !f.startsWith('cong_thuc') && fs.statSync(path.join(dir, f)).isFile());

const out = [];
out.push(`FORM: ${group}/${form}`);

let totalCells = 0;
let totalDistinct = 0;

for (const f of files.sort()) {
  const buf = fs.readFileSync(path.join(dir, f));
  const ext = path.extname(f).toLowerCase();
  let kind = ext === '.xlsx' ? 'xlsx' : ext === '.xls' || ext === '.xls1' ? 'xls' : magicKind(buf);
  if (!kind) continue;
  let wb;
  try {
    wb = XLSX.read(buf, { type: 'buffer', cellFormula: true });
  } catch (e) {
    out.push(`\n=== FILE: ${f} === (không đọc được: ${e.message})`);
    continue;
  }
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn] || {};
    const refList = Object.keys(ws).filter((r) => !r.startsWith('!'));
    if (refList.length === 0) continue;
    // text cells (for labels)
    const text = new Map(); // ref -> string
    for (const ref of refList) {
      const c = ws[ref];
      if (c && c.t === 's' && c.v != null && String(c.v).trim() !== '') {
        text.set(ref, String(c.v).replace(/\s+/g, ' ').slice(0, 90));
      }
    }
    // text lookup by row / col
    const byRow = new Map();
    const byCol = new Map();
    for (const ref of text.keys()) {
      const [r, c] = refKey(ref);
      if (!byRow.has(r)) byRow.set(r, new Map());
      byRow.get(r).set(c, ref);
      if (!byCol.has(c)) byCol.set(c, new Map());
      byCol.get(c).set(r, ref);
    }
    function nearestLeft(row, col, maxGap) {
      const m = byRow.get(row);
      if (!m) return '';
      let best = null;
      for (const c of m.keys()) if (c < col && (best === null || c > best)) best = c;
      if (best === null || col - best > maxGap) return '';
      return text.get(m.get(best));
    }
    function nearestAbove(row, col, maxGap) {
      const m = byCol.get(col);
      if (!m) return '';
      let best = null;
      for (const r of m.keys()) if (r < row && (best === null || r > best)) best = r;
      if (best === null || row - best > maxGap) return '';
      return text.get(m.get(best));
    }
    // group formulas
    const groups = new Map(); // formula -> {refs:[]}
    for (const ref of refList) {
      const c = ws[ref];
      if (!c) continue;
      const fml = c.f || c.F;
      if (!fml) continue;
      const key = String(fml);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(ref);
    }
    out.push(`\n=== FILE: ${f} | SHEET: ${sn} === (${refList.length} cells, ${groups.size} patterns khác nhau)`);
    const arr = [...groups.entries()];
    arr.sort((a, b) => {
      const [ra, ca] = refKey(a[1][0]);
      const [rb, cb] = refKey(b[1][0]);
      return ra - rb || ca - cb;
    });
    const shown = arr.slice(0, MAX_PATTERN_PER_SHEET);
    totalCells += refList.length;
    totalDistinct += groups.size;
    for (const [fml, refs] of shown) {
      refs.sort((x, y) => {
        const [rx, cx] = refKey(x);
        const [ry, cy] = refKey(y);
        return rx - ry || cx - cy;
      });
      const [r0, c0] = refKey(refs[0]);
      const left = nearestLeft(r0, c0, 8);
      const above = nearestAbove(r0, c0, 3);
      const refsShown = refs.slice(0, MAX_REFS).join(',');
      const more = refs.length > MAX_REFS ? ` (+${refs.length - MAX_REFS} nữa)` : '';
      const label = [left, above].filter(Boolean).join(' | ') || '(không có nhãn lân cận)';
      out.push(`--- [${refs[0]}] x${refs.length} :: ${refsShown}${more}`);
      out.push(`    LABEL: ${label}`);
      out.push(`    FORMULA: ${fml.length > 400 ? fml.slice(0, 400) + ' …[cắt]' : fml}`);
    }
    if (arr.length > shown.length) {
      out.push(`… còn ${arr.length - shown.length} pattern khác (xem cong_thuc.txt phần thô)`);
    }
  }
}
out.push(`\nTỔNG: ${totalCells} cell, ${totalDistinct} pattern khác nhau`);
console.log(out.join('\n'));
