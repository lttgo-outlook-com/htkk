#!/usr/bin/env node
/**
 * merge_cong_thuc.mjs — fold the human-readable section into the head of cong_thuc.txt.
 *
 * For every form folder under _forms/<group>/<form>/ that has BOTH cong_thuc.txt
 * (raw Excel formulas) and cong_thuc_hr.txt (human-readable labeled summary),
 * rewrite cong_thuc.txt as:
 *     [hr section]  +  separator  +  [original raw content]
 * then delete the now-redundant cong_thuc_hr.txt.
 *
 * Orphan handling: if a folder has cong_thuc_hr.txt but NO cong_thuc.txt (the raw
 * file was moved to _to_delete because it held no formulas), move the hr file
 * into _to_delete/<group>/<form>/ so _forms/ stays clean.
 *
 * Dry-run by default (reports what it would do). Pass --write to apply.
 *   node _re/merge_cong_thuc.mjs          # report
 *   node _re/merge_cong_thuc.mjs --write  # apply
 */
import fs from 'node:fs';
import path from 'node:path';

const ROOT = '/Users/tuanle/codes/HTKK/_forms';
const TO_DELETE = '/Users/tuanle/codes/HTKK/_to_delete';
const WRITE = process.argv.includes('--write');
const SEP = '\n' + '='.repeat(78) + '\nDƯỚI ĐÂY LÀ DANH SÁCH CELL CÔNG THỨC NGUYÊN THÔ (RAW)\n' + '='.repeat(78) + '\n\n';

// Enumerate all depth-2 form dirs: _forms/<group>/<form>/
const formDirs = [];
for (const g of fs.readdirSync(ROOT, { withFileTypes: true })) {
  if (!g.isDirectory()) continue;
  for (const f of fs.readdirSync(path.join(ROOT, g.name), { withFileTypes: true })) {
    if (f.isDirectory()) formDirs.push(`${g.name}/${f.name}`);
  }
}
formDirs.sort();
let merged = 0;      // raw + hr  -> merge, drop hr
const orphan = [];   // hr only (raw was moved to _to_delete) -> relocate hr
const missing = [];  // raw but no hr  (agent not done yet)

for (const rel of formDirs) {
  const d = path.join(ROOT, rel);
  const raw = path.join(d, 'cong_thuc.txt');
  const hr = path.join(d, 'cong_thuc_hr.txt');
  const hasRaw = fs.existsSync(raw);
  const hasHr = fs.existsSync(hr);
  if (hasRaw && hasHr) {
    const rawTxt = fs.readFileSync(raw, 'utf8');
    const hrTxt = fs.readFileSync(hr, 'utf8').trimEnd() + '\n';
    if (WRITE) {
      fs.writeFileSync(raw, hrTxt + SEP + rawTxt);
      fs.rmSync(hr);
    }
    merged++;
  } else if (hasHr && !hasRaw) {
    orphan.push(rel);
    if (WRITE) {
      const destDir = path.join(TO_DELETE, rel);
      fs.mkdirSync(destDir, { recursive: true });
      fs.renameSync(hr, path.join(destDir, 'cong_thuc_hr.txt'));
    }
  } else if (hasRaw && !hasHr) {
    missing.push(rel);
  }
}

console.log(`form folders: ${formDirs.length}`);
console.log(`  raw+hr (merge target): ${merged}`);
console.log(`  hr-only orphan (raw moved to _to_delete) -> relocate: ${orphan.length}`);
for (const o of orphan) console.log('    - ' + o);
console.log(`  raw-without-hr (agent pending or not produced): ${missing.length}`);
for (const m of missing) console.log('    - ' + m);
if (WRITE) console.log(`\nWROTE: merged ${merged}, relocated ${orphan.length} orphan hr file(s) to _to_delete`);
else console.log('\nDRY RUN — pass --write to apply.');
