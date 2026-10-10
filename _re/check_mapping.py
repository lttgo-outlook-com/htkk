#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_mapping.py - independent check of _forms/ against _extracted/ (run: python3 -I).

(a) every template file (.xls/.xls1/.xml family, incl. cab-truncated names) in _extracted,
    excluding catalogue_*, dm_*, menu/config, has exactly one place in _forms (manifest row,
    destination exists, sha256 equal) or one row in _re/pending_mapping.csv; and every data
    file in _forms has a manifest row (no orphan).
(b) sheet signature: in each form folder and layer (.xls / .xls1), every form workbook (.xls/.xls1, OLE; bang_ke_/khbs_/
    pl_/bao_cao_ excluded as in _audit/qc2/sheetsig.cjs) agrees with the others:
    b1 = main-sheet prefix equals the folder majority (sheetsig.cjs rule);
    b2 = sheet-name overlap |A&B|/min(|A|,|B|) >= 0.5 with at least one other workbook.
(c) the 10 OLE templates of audit item (2) + header.xls + in_tmp.xls have a cong_thuc.txt
    section "=== FILE: <name> ===" in their folder.
(d) number of form folders, folders with cong_thuc.txt at git HEAD and now.
(e) one own template stem per folder; rule R5x <=> path under _ngoai_menu/; every other folder
    (outside _shared/_ngoai_menu) is an upper-cased menu.xml FolderName or MaFile.
(f) audit reconciliation; every file with formulas (scan_extracted.json) placed in _forms
    (incl. _ngoai_menu) has a cong_thuc.txt section; pending files with formulas listed with reason.
(g) every "=== FILE: X ===" section in a cong_thuc.txt names a file X in the same folder.
(h) every file named in ShowBangKe("...") (business classes + frmHTKKSystems.cs) that exists in
    _extracted (28-char cab name, +1 digit on collision) is in the manifest with ShowBangKe
    evidence, or in pending with non-empty candidates.
Reads only; never executes anything from _extracted.
Env: HTKK_SRC (default /Users/tuanle/codes/HTKK/_extracted), FORMS_DST (default <repo>/_forms),
     HTKK_DECOMPILED (default /Users/tuanle/codes/HTKK/_decompiled; read as text only).
"""
import os, re, sys, csv, struct, hashlib, subprocess, collections

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.environ.get('HTKK_SRC', '/Users/tuanle/codes/HTKK/_extracted')
F = os.environ.get('FORMS_DST', os.path.join(ROOT, '_forms'))
MANIFEST = os.path.join(F, '_manifest.csv')
PENDING = os.path.join(ROOT, '_re', 'pending_mapping.csv')
NOT_DATA = {'cong_thuc.txt', '.DS_Store', '_manifest.csv'}
CONFIG = {'menu.xml', 'versiontokhai.xml', 'message.xml', 'config.xml', 'config_header.xml',
          'convertdatacommon.xml', 'convertversion.xml', 'createdatafilekhbsmiss.xml',
          'replacedatafile.xml', 'validcombobox.xml'}
TEMPLATE_ROOT = re.compile(r'^(Sections|HSoThueDTu|TKhai(_.*)?|PL.*|BKe.*|ToKhai.*|Workbook)$', re.I)
REQUIRED = ['_1_ac_blp.xls', '_2_pthu_vsp_tt80_251_.xls', '_3_tbac_blp.xls', '_4_b01_htx_bctc.xls',
            '_5_ds_tncn.xls', '_8_bctc.xls', '_9_bctc_dd.xls', 'bao_cao_01bc_sdhd_cnkd.xls',
            'bao_cao_bc26_ac.xls', 'bao_cao_bc26_ac_blp.xls', 'header.xls', 'in_tmp.xls']

def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()

def classify(path):
    """-> 'xls' | 'xml' | 'xsd' | None (not a template)."""
    name = os.path.basename(path).lower()
    if name.startswith(('catalogue', 'dm_', 'dmngayle', 'menu_backup')) or name in CONFIG:
        return None
    with open(path, 'rb') as f:
        b = f.read(4096)
    ext = os.path.splitext(name.rstrip('.'))[1]
    if b[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' or b[:2] == b'PK':
        return None if ext == '.doc' else 'xls'
    txt = None
    if b[:2] in (b'\xff\xfe', b'\xfe\xff'):
        txt = b.decode('utf-16', 'replace')
    elif b.lstrip(b'\xef\xbb\xbf \r\n\t')[:1] == b'<':
        txt = b.lstrip(b'\xef\xbb\xbf \r\n\t').decode('utf-8', 'replace')
    if txt is not None:
        m = re.search(r'<(?![?!])([A-Za-z_][\w.\-:]*)', txt)
        root = m.group(1).split(':')[-1] if m else ''
        if root == 'schema':
            return 'xsd'
        return 'xml' if TEMPLATE_ROOT.match(root or '-') else None
    if ext == '.xls' and b[:2] in (b'\x09\x08', b'\x09\x04', b'\x09\x02', b'\x03\x04'):
        return 'xls'   # raw BIFF stream without OLE container
    return None

def sheet_names(path):
    d = open(path, 'rb').read()
    if d[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        return None
    try:
        ssz = 1 << struct.unpack_from('<H', d, 30)[0]
        mssz = 1 << struct.unpack_from('<H', d, 32)[0]
        nfat, dstart = struct.unpack_from('<II', d, 44)
        mcut, mfs, _, dfs, ndf = struct.unpack_from('<IIIII', d, 56)
        difat = list(struct.unpack_from('<109I', d, 76))
        s = dfs
        for _ in range(ndf):
            v = struct.unpack_from('<%dI' % (ssz // 4), d, (s + 1) * ssz)
            difat += v[:-1]; s = v[-1]
        fat = []
        for x in difat[:nfat]:
            fat += struct.unpack_from('<%dI' % (ssz // 4), d, (x + 1) * ssz)
        def chain(st):
            out, seen = [], set()
            while st < 0xFFFFFFFA and st not in seen and st < len(fat):
                seen.add(st); out.append(st); st = fat[st]
            return out
        rd = lambda st: b''.join(d[(x + 1) * ssz:(x + 2) * ssz] for x in chain(st))
        db = rd(dstart)
        ents = [(db[i:i + 128][:max(struct.unpack_from('<H', db, i + 64)[0] - 2, 0)].decode('utf-16le', 'replace'),
                 db[i + 66], struct.unpack_from('<I', db, i + 116)[0], struct.unpack_from('<I', db, i + 120)[0])
                for i in range(0, len(db) - 127, 128)]
        wb = next((e for e in ents if e[0] in ('Workbook', 'Book') and e[1] == 2), None)
        if wb is None:
            return None
        if wb[3] < mcut:
            ms = rd(ents[0][2]); mf = []
            for x in chain(mfs):
                mf += struct.unpack_from('<%dI' % (ssz // 4), d, (x + 1) * ssz)
            parts, st, seen = [], wb[2], set()
            while st < 0xFFFFFFFA and st not in seen and st < len(mf):
                seen.add(st); parts.append(ms[st * mssz:(st + 1) * mssz]); st = mf[st]
            data = b''.join(parts)[:wb[3]]
        else:
            data = rd(wb[2])[:wb[3]]
        names, pos = [], 0
        while pos + 4 <= len(data):
            t, l = struct.unpack_from('<HH', data, pos)
            b = data[pos + 4:pos + 4 + l]
            if t == 0x85 and len(b) >= 8:
                n, hi = b[6], b[7] & 1
                names.append(b[8:8 + n * (2 if hi else 1)].decode('utf-16le' if hi else 'latin-1'))
            elif t == 0x0A:
                break
            pos += 4 + l
        return names
    except (struct.error, IndexError):
        return None

# ShowBangKe("<file>") targets, scanned independently of reorg_forms.py
DECOMP = os.environ.get('HTKK_DECOMPILED', '/Users/tuanle/codes/HTKK/_decompiled')
SBK_TRUNC = 28
sbk_calls = collections.defaultdict(set)    # 28-char cab name -> full called names
for dp_, dn_, fn_ in os.walk(DECOMP):
    for f_ in fn_:
        if f_.endswith('.cs'):
            for m_ in re.finditer(r'ShowBangKe\("([^"]+)"\)', open(os.path.join(dp_, f_), encoding='utf-8', errors='replace').read()):
                sbk_calls[m_.group(1).lower()[:SBK_TRUNC]].add(m_.group(1).lower())
def sbk_target(n):
    low = n.lower()
    if low in sbk_calls:
        return low
    k = low[:SBK_TRUNC]
    if len(low) == SBK_TRUNC + 1 and low[-1].isdigit() and k in sbk_calls and all(len(x) > SBK_TRUNC for x in sbk_calls[k]):
        return k
    return None

ok_all = True
# ------------------------------------------------------------------ (a)
universe = {}
sbk_files = {}
for n in sorted(os.listdir(SRC)):
    p = os.path.join(SRC, n)
    if os.path.isfile(p):
        k = classify(p)
        t = sbk_target(n)
        if t:
            sbk_files[n] = t
            k = k or 'sbk'
        if k:
            universe[n] = k
man = list(csv.DictReader(open(MANIFEST, encoding='utf-8')))
pend = list(csv.DictReader(open(PENDING, encoding='utf-8')))
by_src = collections.Counter(r['src'] for r in man)
in_pend = collections.Counter(r['src'] for r in pend)
problems = []
for n, k in universe.items():
    c = by_src[n] + in_pend[n]
    if c != 1:
        problems.append('%s: %d vi tri (manifest %d, pending %d)' % (n, c, by_src[n], in_pend[n]))
bad_sha, missing_dst = [], []
for r in man:
    p = os.path.join(F, r['dst'])
    if not os.path.exists(p):
        missing_dst.append(r['dst'])
    elif sha(p) != sha(os.path.join(SRC, r['src'])) or r['sha256'] != sha(p):
        bad_sha.append(r['dst'])
dst_set = {r['dst'] for r in man}
orphans = []
for dp, dn, fn in os.walk(F):
    for f in fn:
        rel = os.path.relpath(os.path.join(dp, f), F)
        if f not in NOT_DATA and rel not in dst_set:
            orphans.append(rel)
extra_src = [r['src'] for r in man if r['src'] not in universe]
kinds = collections.Counter(universe.values())
pk = collections.Counter(universe[r['src']] for r in pend if r['src'] in universe)
print('(a) template files in _extracted: %d (xls-family %d, xml %d, xsd %d, other ShowBangKe target %d)' % (
    len(universe), kinds['xls'], kinds['xml'], kinds['xsd'], kinds['sbk']))
print('    placed in _forms: %d | pending: %d (xls %d, xml %d, xsd %d)' % (
    sum(1 for n in universe if by_src[n] == 1), sum(1 for n in universe if in_pend[n] == 1), pk['xls'], pk['xml'], pk['xsd']))
print('    files with !=1 location: %d | manifest dst missing: %d | sha mismatch: %d | orphan files in _forms: %d | manifest src outside universe: %d'
      % (len(problems), len(missing_dst), len(bad_sha), len(orphans), len(extra_src)))
for x in (problems + missing_dst + bad_sha + orphans + extra_src)[:20]:
    print('     !', x)
ok_a = not (problems or missing_dst or bad_sha or orphans or extra_src)
ok_all &= ok_a
# ------------------------------------------------------------------ (b)
SKIP = re.compile(r'^(bang_?ke|khbs|pl_|bao_cao)', re.I)
def norm(s): return re.sub(r'[\s\-_/]+', '_', s.strip().upper())
def main_prefix(names):
    c = sorted(s for s in names if not re.search('header|khbs', s, re.I))
    return re.sub(r'[-_ ].*', '', c[0]) if c else ''
def sset(names): return {norm(s) for s in names if not re.search('header|khbs', s, re.I)}
rule_of = {r['dst']: r['rule'] for r in man}
b1, b2, exc, checked = [], [], [], 0
for g in sorted(os.listdir(F)):
    gp = os.path.join(F, g)
    if not os.path.isdir(gp):
        continue
    for d in sorted(os.listdir(gp)):
        dp = os.path.join(gp, d)
        if not os.path.isdir(dp):
            continue
        layers = collections.defaultdict(dict)   # .xls (formula layer) vs .xls1 (prefill layer)
        for f in sorted(os.listdir(dp)):
            m = re.search(r'\.(xls1?)$', f, re.I)
            if m and not SKIP.match(f):
                s = sheet_names(os.path.join(dp, f))
                if s is not None:
                    layers[m.group(1).lower()][f] = s
        for layer, sig in layers.items():
            if len(sig) < 2:
                continue
            checked += 1
            maj = collections.Counter(main_prefix(v) for v in sig.values()).most_common(1)[0][0]
            for f, v in sig.items():
                a = sset(v)
                best = max((len(a & sset(w)) / max(1, min(len(a), len(sset(w)))) for h, w in sig.items() if h != f), default=1)
                msgs = []
                if main_prefix(v) != maj:
                    msgs.append('main=%s majority=%s' % (main_prefix(v), maj))
                if a and best < 0.5:
                    msgs.append('sheets=%s best_overlap=%.2f' % (v, best))
                if not msgs:
                    continue
                line = '%s/%s: %s %s' % (g, d, f, '; '.join(msgs))
                rule = rule_of.get('%s/%s/%s' % (g, d, f), '')
                if rule in ('R1', 'R1m'):
                    exc.append(line + ' [ngoai le: %s, ten tep == FolderName/MaFile + version cua menu; khac biet nam trong noi dung mau giua cac version]' % rule)
                elif main_prefix(v) != maj:
                    b1.append(line)
                else:
                    b2.append(line)
print('(b) folder-layers with >=2 form workbooks: %d | b1 main-sheet mismatches: %d | b2 low-overlap: %d | exceptions with reason: %d'
      % (checked, len(b1), len(b2), len(exc)))
for x in b1 + b2:
    print('     !', x)
for x in exc:
    print('     ~', x)
ok_all &= not (b1 or b2)
# ------------------------------------------------------------------ (c)
dst_of = {r['src']: r['dst'] for r in man}
okc = 0
for n in REQUIRED:
    d = dst_of.get(n)
    if not d:
        print('     ! %s: khong co trong manifest' % n)
        continue
    ct = os.path.join(F, os.path.dirname(d), 'cong_thuc.txt')
    t = open(ct, encoding='utf-8').read() if os.path.exists(ct) else ''
    sec = re.search(r'=== FILE: %s ===\n(.*?)=> (\d+) cell' % re.escape(os.path.basename(d)), t, re.S)
    if sec:
        okc += 1
        print('    %-28s -> %-58s %5s cell' % (n, d, sec.group(2)))
    else:
        print('     ! %s -> %s: thieu muc trong cong_thuc.txt' % (n, d))
print('(c) required workbooks with cong_thuc.txt section: %d/%d' % (okc, len(REQUIRED)))
ok_all &= okc == len(REQUIRED)
# ------------------------------------------------------------------ (d)
def form_dirs_now():
    out = []
    for g in os.listdir(F):
        gp = os.path.join(F, g)
        if os.path.isdir(gp):
            out += [g + '/' + d for d in os.listdir(gp) if os.path.isdir(os.path.join(gp, d))]
    return out
now = form_dirs_now()
now_ct = sum(1 for x in now if os.path.exists(os.path.join(F, x, 'cong_thuc.txt')))
BASE = sys.argv[sys.argv.index('--base') + 1] if '--base' in sys.argv else None
try:
    if BASE is None:
        BASE = subprocess.run(['git', '-C', ROOT, 'merge-base', 'HEAD', 'main'], capture_output=True,
                              text=True, check=True).stdout.strip()
    ls = subprocess.run(['git', '-C', ROOT, 'ls-tree', '-r', '--name-only', BASE, '--', '_forms'],
                        capture_output=True, text=True, check=True).stdout.split('\n')
    head_dirs = {'/'.join(p.split('/')[1:3]) for p in ls if p.count('/') >= 3}
    head_ct = sum(1 for p in ls if p.count('/') == 3 and p.endswith('/cong_thuc.txt'))
    print('(d) form folders: base %s %d -> now %d | with cong_thuc.txt: base %d -> now %d'
          % (BASE[:7], len(head_dirs), len(now), head_ct, now_ct))
except (subprocess.CalledProcessError, FileNotFoundError):
    print('(d) form folders now %d, with cong_thuc.txt %d (git HEAD not readable)' % (len(now), now_ct))
# ------------------------------------------------------------------ (e) mixed stems
FULL = re.compile(r'^(?P<stem>.+?)(?:_xml\d?)?_(?P<ver>\d{3,5})(?:_|_bu)?$')
stems = collections.defaultdict(set)
for r in man:
    if r['rule'].split('/')[0] in ('R1', 'R1m', 'R5v', 'R5n', 'R6') and not re.match(r'^(bang_?ke_|khbs_|pl_|bao_cao_)', r['src']):
        n = r['src'].rstrip('.').rsplit('.', 1)[0].lower()
        m = FULL.match(n)
        stems['/'.join(r['dst'].split('/')[:2])].add(m.group('stem') if m else n)
mixed = sorted(k for k, v in stems.items() if len(v) > 1)
import xml.etree.ElementTree as ET
menu_names = set()
for m_ in ET.parse(os.path.join(SRC, 'menu.xml')).getroot().iter('Menu'):
    if m_.get('MaFile'):
        menu_names.add(m_.get('MaFile').upper())
    for v_ in m_.findall('Validity'):
        if v_.get('FolderName'):
            menu_names.add(v_.get('FolderName').upper())
x_bad = [r['dst'] for r in man if (r['rule'] == 'R5x') != r['dst'].startswith('_ngoai_menu/')]
not_menu = sorted({'/'.join(r['dst'].split('/')[:2]) for r in man
                   if r['dst'].split('/')[0] not in ('_shared', '_ngoai_menu') and r['dst'].split('/')[1] not in menu_names})
xdirs = {'/'.join(r['dst'].split('/')[:2]) for r in man if r['dst'].startswith('_ngoai_menu/')}
print('(e) folders with >1 own template stem (R1/R1m/R5v/R5n, no prefix): %d | R5x outside/non-R5x inside _ngoai_menu: %d'
      ' | folders outside menu in regular groups: %d | _ngoai_menu: %d folders, %d files'
      % (len(mixed), len(x_bad), len(not_menu), len(xdirs), sum(1 for r in man if r['dst'].startswith('_ngoai_menu/'))))
for x in mixed + x_bad + not_menu:
    print('     !', x)
ok_all &= not mixed and not x_bad and not not_menu
# ------------------------------------------------------------------ (f) audit reconciliation
AUD = os.environ.get('AUDIT_DIR', '/Users/tuanle/codes/HTKK/_audit')
rep_csv = os.path.join(AUD, 'out', 'completeness', 'report.csv')
scan = os.path.join(AUD, 'qc1', 'scan_extracted.json')
if os.path.exists(rep_csv) and os.path.exists(scan):
    import json
    placed, pset = set(by_src), set(in_pend)
    miss = []
    rows = [r for r in csv.DictReader(open(rep_csv, encoding='utf-8-sig'))
            if r['loai_thieu'] in ('tep_mau_ten_cat_chua_map', 'tep_mau_chua_map_vao_forms')]
    cnt = collections.Counter()
    for r in rows:
        n = r['duong_dan_nguon'].split('/')[-1]
        st_ = 'manifest' if n in placed else 'pending' if n in pset else 'excluded:catalogue' if n.startswith('catalogue') else 'MISSING'
        cnt[st_] += 1
        if st_ == 'MISSING':
            miss.append('report.csv row not placed: ' + n)
    print('(f) audit report.csv unmapped rows: %d -> %s' % (len(rows), dict(cnt)))
    sc = json.load(open(scan, encoding='utf-8'))
    c2 = collections.Counter(); miss2 = []; pend_f = []
    for k, v in sc.items():
        if not v.get('n'):
            continue
        n = k.split('/')[-1]
        if n.lower().startswith('catalogue'):
            c2['excluded:catalogue'] += 1; continue
        if n in in_pend:
            pr = next(r for r in pend if r['src'] == n)
            c2['pending'] += 1; pend_f.append('pending: %s (%d cell) | %s | %s' % (n, v['n'], pr['reason'], pr['candidates'][:160])); continue
        d = dst_of.get(n)
        if not d:
            c2['MISSING'] += 1; miss2.append('missing: ' + n); continue
        ct = os.path.join(F, os.path.dirname(d), 'cong_thuc.txt')
        ok = os.path.exists(ct) and ('=== FILE: %s ===' % os.path.basename(d)) in open(ct, encoding='utf-8').read()
        c2['in_cong_thuc' if ok else 'NO_SECTION'] += 1
        if not ok:
            miss2.append('no section: %s -> %s' % (n, d))
    print('    scan_extracted.json files with formulas: %d -> %s' % (sum(c2.values()), dict(c2)))
    for x in miss + miss2:
        print('     !', x)
    for x in pend_f:
        print('     ~', x)
    ok_all &= not miss and not c2['MISSING'] and not c2['NO_SECTION']
else:
    print('(f) SKIP: thieu %s hoac %s' % (rep_csv, scan))
    ok_all = False
# ------------------------------------------------------------------ (g) cong_thuc sections point to files in the folder
stale_sec, n_sec = [], 0
for dp_, dn_, fn_ in os.walk(F):
    if 'cong_thuc.txt' in fn_:
        for x in re.findall(r'^=== FILE: (.+?) ===$', open(os.path.join(dp_, 'cong_thuc.txt'), encoding='utf-8').read(), re.M):
            n_sec += 1
            if x not in fn_:
                stale_sec.append('%s: %s' % (os.path.relpath(dp_, F), x))
print('(g) cong_thuc.txt sections: %d | pointing to a file not in the folder: %d' % (n_sec, len(stale_sec)))
for x in stale_sec:
    print('     !', x)
ok_all &= not stale_sec
# ------------------------------------------------------------------ (h) ShowBangKe targets
pend_by = {r['src']: r for r in pend}
man_by = {r['src']: r for r in man}
bad_h, cnt_h = [], collections.Counter()
for n, t in sorted(sbk_files.items()):
    if n in man_by and 'ShowBangKe' in man_by[n]['evidence']:
        cnt_h['manifest'] += 1
    elif n in pend_by and pend_by[n]['candidates'].strip():
        cnt_h['pending'] += 1
    else:
        bad_h.append('%s (goi: %s): %s' % (n, ','.join(sorted(sbk_calls[t])),
                     'manifest khong co bang chung ShowBangKe' if n in man_by else 'pending khong co ung vien' if n in pend_by else 'khong o dau'))
print('(h) ShowBangKe names: %d | files in _extracted: %d -> %s | not covered: %d'
      % (sum(len(v) for v in sbk_calls.values()), len(sbk_files), dict(cnt_h), len(bad_h)))
for x in bad_h:
    print('     !', x)
ok_all &= not bad_h
excl = collections.Counter()
for n in os.listdir(SRC):
    if os.path.isfile(os.path.join(SRC, n)) and n not in universe:
        ln = n.lower()
        excl['catalogue_/dm_' if ln.startswith(('catalogue', 'dm_', 'dmngayle')) else
             'menu/config' if ln in CONFIG or ln.startswith('menu_backup') else 'khac (anh, dll, xml cau hinh)'] += 1
print('    excluded from universe: ' + ', '.join('%s=%d' % kv for kv in sorted(excl.items())))
rules = collections.Counter(r['rule'].split('/')[0] for r in man)
print('    manifest rules: ' + ', '.join('%s=%d' % kv for kv in sorted(rules.items())))
print('RESULT: %s' % ('PASS' if ok_all else 'FAIL'))
sys.exit(0 if ok_all else 1)
