#!/usr/bin/env node
/**
 * extract_formulas.mjs — extract spreadsheet formulas from one HTKK form folder.
 *
 * Reads every Excel template in the given form folder (._forms/<group>/<form>/):
 *   .xls / .xls1 (BIFF8, OLE) and .xlsx (OOXML) — plus extensionless files
 *   detected by magic bytes — and collects every cell that carries a formula.
 *
 * Output: cong_thuc.txt inside that folder (one section per source file, so
 * formulas from different files never mix).
 *
 * Usage:
 *   node _re/extract_formulas.mjs <form-folder>
 *
 * Requires the SheetJS module; resolved from $XLFORM_DIR (default /tmp/xlform,
 * set up with: mkdir -p /tmp/xlform && cd /tmp/xlform && npm init -y && npm i xlsx@0.18.5)
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
const outFile = path.join(dir, 'cong_thuc.txt');

const EXT_XLS = new Set(['.xls', '.xls1', '.xlsx']);

function magicKind(buf) {
  if (buf.length < 4) return null;
  if (buf[0] === 0x50 && buf[1] === 0x4b) return 'xlsx'; // ZIP
  if (buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0) return 'xls'; // OLE
  return null;
}

function refKey(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return [0, 0, ref];
  const col = [...m[1]].reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0);
  return [Number(m[2]), col, ref];
}

function readFile(p) {
  const ext = path.extname(p).toLowerCase();
  const buf = fs.readFileSync(p);
  let kind = EXT_XLS.has(ext) ? (ext === '.xlsx' ? 'xlsx' : 'xls') : null;
  if (!kind) kind = magicKind(buf); // extensionless
  if (!kind) return null; // not a spreadsheet
  try {
    return XLSX.read(buf, { type: 'buffer', cellFormula: true });
  } catch (e) {
    return { __error: String(e.message || e) };
  }
}

function extract(wb) {
  // Map: sheetName -> [{ref, formula}]
  const perSheet = new Map(wb.SheetNames.map((s) => [s, []]));
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    for (const ref of Object.keys(ws)) {
      if (ref.startsWith('!')) continue;
      const c = ws[ref];
      const f = c && (c.f || c.F);
      if (f) perSheet.get(sn).push({ ref, formula: String(f) });
    }
  }
  for (const list of perSheet.values()) {
    list.sort((a, b) => {
      const ka = refKey(a.ref), kb = refKey(b.ref);
      return ka[0] - kb[0] || ka[1] - kb[1] || ka[2].localeCompare(kb[2]);
    });
  }
  return perSheet;
}

const files = fs
  .readdirSync(dir)
  .filter((f) => f !== 'cong_thuc.txt' && fs.statSync(path.join(dir, f)).isFile());

const sections = [];
let total = 0;
const errors = [];

for (const f of files.sort()) {
  const wb = readFile(path.join(dir, f));
  if (!wb) continue; // not a spreadsheet
  if (wb.__error) {
    errors.push(`${f}: ${wb.__error}`);
    sections.push(`=== FILE: ${f}\nLỖI: không đọc được — ${wb.__error}\n`);
    continue;
  }
  const perSheet = extract(wb);
  let n = 0;
  const lines = [`=== FILE: ${f} ===`];
  for (const sn of wb.SheetNames) {
    const cells = perSheet.get(sn);
    lines.push(`--- Sheet "${sn}" ---`);
    if (cells.length === 0) {
      lines.push('(không có cell công thức)');
    } else {
      for (const { ref, formula } of cells) {
        lines.push(`${ref}: ${formula}`);
        n++;
      }
    }
  }
  total += n;
  lines.push(`=> ${n} cell công thức trong file này\n`);
  sections.push(lines.join('\n'));
}

if (sections.length === 0) {
  fs.writeFileSync(
    outFile,
    `CÔNG THỨC TÍNH — ${form} (${group})\nTrích xuất: 2026-10-08 — thư mục không có file Excel template (chỉ có XML/XSD hoặc khác) → không có công thức spreadsheet.\n`,
    'utf8',
  );
  console.log(`${form}: no spreadsheet files`);
  process.exit(0);
}

const header = [
  `CÔNG THỨC TÍNH — ${form}`,
  `Thư mục: ${group}/${form} (HTKK 5.7.9, _forms/)`,
  `Trích xuất: 2026-10-08 — mọi cell chứa công thức trong các file Excel (mẫu CMC) của form này, theo từng file để không lẫn.`,
  `Ghi chú: sheet "Header" = khối thông tin chung người nộp thuế; cell dạng "REF: công thức" (REF = địa chỉ ô, A1 notation, tham chiếu sheet khác dạng "Sheet!REF").`,
  '',
];
const footer = [`\nTỔNG CỘNG: ${total} cell công thức trong ${sections.length} file\n`];
if (errors.length) footer.push(`LỖI ĐỌC: ${errors.join(' | ')}\n`);
fs.writeFileSync(outFile, [...header, ...sections, ...footer].join('\n'), 'utf8');
console.log(`${form}: ${total} formula cells in ${sections.length} file(s)`);
