#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reorg_forms.py — reorganize HTKK 5.7.9 form/report templates.

Source:  _extracted/  (flat, 4526 files, extracted from Data1.cab)
Target:  _forms/      (grouped by business -> form -> renamed files)

Catalogue of record: _extracted/menu.xml (314 Menu, 242 form-level MaFile).
NOTES.md §6.2 = revised business classification (sanity check only).

Usage:
  python3 _re/reorg_forms.py            # DRY RUN: print mapping + unmatched
  python3 _re/reorg_forms.py --write    # copy + rename into _forms/
"""
import os, re, sys, shutil, collections, xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, '_extracted')
DST = os.environ.get('FORMS_DST', os.path.join(ROOT, '_forms'))
WRITE = '--write' in sys.argv

# ---------------------------------------------------------------- catalogue
GROUP_SLUGS = {
    '101_1':  '01_GTGT',
    '101_2':  '02_TNDN',
    '101_3':  '03_TAIN',
    '101_4':  '04_TTDB',
    '101_6':  '06_BCTC',
    '101_7':  '07_HOA_DON',
    '101_8':  '08_THUE_NHA_THAU',
    '101_9':  '09_PHI_LE_PHI',
    '101_11': '11_TNCN',
    '101_12': '12_THUE_BVMT',
    '101_15': '15_THUE_KHOAN_THUE_TAI_SAN',
    '101_16': '16_BIEN_LAI',
    '101_17': '17_THUY_DIEN',
    '101_18': '18_DAU_KHI',
    '101_20': '20_TEM_RUOU',
    '101_24': '24_CO_TUC_LOI_NHUAN',
    '101_25': '25_HOAN_THUE',
    '101_26': '26_GIA_HAN_NOP_THUE',
    '101_28': '28_SHTK',
    '101_29': '29_NCC_NUOC_NGOAI',
    '101_30': '30_MIEN_TIEN_CHAM_NOP',
    '101_31': '31_DAT_DAI',
    '101_32': '32_TAI_SAN_BAO_DAM',
    '101_33': '33_DOI_KY_TINH_THUE',
    '101_34': '34_THUONG_MAI_DIEN_TU',
    '101_35': '35_CHUYEN_DOI_CO_SO_HUU',
    '101_36': '36_THUE_TOI_THIEU_TOAN_CAU',
    '101_37': '37_HO_CN_KINH_DOANH',
    '101_38': '38_TAI_SAN_MA_HOA',
    'KHBS': 'KHBS',
    'BANG_KE': 'BANG_KE',
    'PHU_LUC': 'PHU_LUC',
    '99_KHAC': '99_KHAC',
}

menus = list(ET.parse(os.path.join(SRC, 'menu.xml')).getroot().findall('Menu'))
byid = {m.get('ID'): m for m in menus}

def group_of(m):
    cur = m
    while cur is not None:
        pid = cur.get('ParentID', '')
        parent = byid.get(pid)
        if parent is None:
            return None
        if re.match(r'^(10[0-4])_\d+$', pid):
            return pid
        cur = parent
    return None

mafile = {}   # MAFILE.upper() -> info
for m in menus:
    mf = m.get('MaFile')
    if not mf:
        continue
    gid = group_of(m)
    gcap = byid[gid].get('Caption', '') if gid and gid in byid else ''
    info = mafile.setdefault(mf.upper(), {
        'group': gid, 'gcap': gcap, 'name': m.get('TenToKhai', ''),
        'tongtu': m.get('MoTaBieuMau', ''), 'versions': {}, 'class': m.get('Class', ''),
        'menus': 0,
    })
    info['menus'] += 1
    info['versions'][m.get('VersionXML', '')] = m.get('VersionKHBS', '')

# stem matchers (case-insensitive)
stem_keys = {}
for mf, info in mafile.items():
    k = mf.lower()
    stem_keys.setdefault(k, info)
    if len(k) > 1 and k[0] == '0' and k[1].isdigit():
        stem_keys.setdefault(k[1:], info)   # "04_gtgt_tt89" -> "4_gtgt_tt89"
    if mf.startswith('200_'):            # BCTC 200_* -> file stem 00_*
        stem_keys.setdefault('00_' + k[4:], info)

stem_ver = {}  # key -> set of version digits (from menu VersionXML)
for k, v in mafile.items():
    vds = set(verstr.replace('.', '') for verstr in v['versions'])
    stem_ver.setdefault(k.lower(), set()).update(vds)
    kl = k.lower()
    k2 = kl[1:] if kl[0] == '0' else kl
    stem_ver.setdefault(k2, set()).update(vds)

def resolve_stem(stem, ver=None):
    s = stem.lower()
    if s in stem_keys:
        return stem_keys[s]
    # cabinet strips the FIRST digit of multi-digit form numbers:
    # 16_TH_DKNPT -> 6_th_dknpt, 107_B01_... -> 07_b01_..., 20_TH_DK_TCT -> 0_th_dk_tct
    cands = [k for k in stem_keys if k and k[0].isdigit() and k[1:] == s]
    if len(cands) == 1:
        return stem_keys[cands[0]]
    if len(cands) > 1 and ver:
        cands2 = [k for k in cands if ver in stem_ver.get(k, set())]
        if len(cands2) == 1:
            return stem_keys[cands2[0]]
    return None

# ---------------------------------------------------------------- file scan
EXT_KIND = {
    'xls':  'mau',
    'xls1': 'mau-b',
    'xlsx': 'mau',
    'xml':  'mau-xml',
    'xm':   'mau-xml',
    'x':    'mau-xml?',   # truncated: could be xls or xml
    'xs':   'xsd',
    'xl':   'mau',
    'xl1':  'mau-b',
    'xml1': 'mau-xml',
    'xml9': 'mau-xml',
    'xsd':  'xsd',
    '':     '?',
}
# per-file truth from `file` output (XLS = composite doc, else text)
ftruth = {}
with open('/tmp/htkk_filetypes.tsv', encoding='utf-8', errors='replace') as f:
    for line in f:
        if '\t' not in line:
            continue
        name, kind = line.rsplit('\t', 1)
        ftruth[name.strip()] = kind.strip()

def true_kind(name, ext):
    t = ftruth.get(name, '')
    if t in ('XLS', 'XLSX'):
        return 'mau'
    if ext in ('x', 'xm', 'xl', 'xs', ''):
        return 'mau-xml' if t in ('XML', 'TEXT') else 'mau?'
    return EXT_KIND.get(ext, '?')

def ver_str(d):
    if len(d) == 3:
        return '.'.join(d)
    return d   # keep raw digits when ambiguous

FORM_RE = re.compile(r'^_?(?P<stem>.+?)(?:_xml)?_(?P<ver>\d{3,5})$')

plan = collections.defaultdict(lambda: collections.defaultdict(dict))  # (group, form) -> ver -> {src: meta}
pending = []   # (name, kind_prefix, rest_stem, ver) — unresolved by exact match
excluded = collections.Counter()

mafile_by_info = {}
for k, v in mafile.items():
    mafile_by_info.setdefault(id(v), k)

def assign(gid, form, ver, name, kind, keepname):
    plan[(gid, form)][ver or 'nv'][name] = \
        {'kind': kind, 'keepname': keepname} if keepname else kind

for name in sorted(os.listdir(SRC)):
    if os.path.isdir(os.path.join(SRC, name)):
        continue
    disp = name.rstrip('.')                     # some cab entries end with '.'
    stem0, ext = os.path.splitext(disp)[0], os.path.splitext(disp)[1].lstrip('.').lower()
    if name.startswith('catalogue_'):
        excluded['catalogue_* (danh muc lieu)'] += 1
        continue
    if name in ('menu.xml', 'versiontokhai.xml', 'message.xml', 'Config.xml'):
        excluded['catalog/config (giu lai _extracted)'] += 1
        continue

    if name.startswith(('bang_ke_', 'bangke_', 'khbs_', 'pl_')):
        rest0 = re.sub(r'^(bang_ke|bangke|khbs|pl)_', '', stem0)
        m = FORM_RE.match(rest0)
        ver, rest = (m.group('ver'), m.group('stem')) if m else (None, rest0)
        info = resolve_stem(rest, ver) if rest else None
        if info is not None:
            assign(info['group'], mafile_by_info[id(info)], ver, name, true_kind(name, ext), True)
        else:
            pending.append((name, name.split('_')[0] + '_', rest, ver))
        continue

    m = FORM_RE.match(stem0)
    if not m:
        excluded['khac (hinh anh/dll/config...)'] += 1
        continue
    stem, ver = m.group('stem'), m.group('ver')
    info = resolve_stem(stem, ver)
    if info is None:
        pending.append((name, '', stem, ver))
        continue
    assign(info['group'], mafile_by_info[id(info)], ver, name, true_kind(name, ext), False)

# ------------------------------------------------- fallback pass (phụ lục)
def split_name(name):
    disp = name.rstrip('.')
    return os.path.splitext(disp)[0], os.path.splitext(disp)[1].lstrip('.').lower()

def tokens(s):
    return [t for t in s.lower().split('_') if t]

def nn_match(rest):
    """'3_1_nd132_tndn_tt80' -> best MaFile whose leading number == first token."""
    ts = tokens(rest)
    if not ts or not ts[0].isdigit():
        return None
    nn = int(ts[0])
    restset = set(ts[1:])
    best, bestscore, bestntok = None, 0, 0
    for mf, info in mafile.items():
        mts = tokens(mf)
        if not mts or not mts[0].lstrip('0').isdigit() or int(mts[0].lstrip('0') or 0) != nn:
            continue
        score = len(restset & set(mts[1:]))
        if score > bestscore or (score == bestscore and score > 0 and bestntok and len(mts) < bestntok):
            best, bestscore, bestntok = (mf, info), score, len(mts)
    return best if bestscore >= 1 else None

BCTC_SHEETS = ('thtc', 'kqhdkd', 'lctttt', 'lcttgt', 'lctt_tl', 'sxkd', 'cdkt')

def bctc_folder(rest):
    ts = tokens(rest)
    return '_'.join(t.upper() for t in ts[:3])

BCTC_KW = ('cdkt', 'sxkd', 'lctttt', 'lcttgt', 'lctt', 'thtc', 'kqhdkd', 'bctc', 'cdtk')
TAX_KW = [
    ('tain', '101_3'), ('tndn', '101_2'), ('tncn', '101_11'),
    ('ntnn', '101_8'), ('gtgt', '101_1'), ('ttdb', '101_4'),
    ('dnht', '101_25'), ('tmdn', '101_31'), ('sddnn', '101_31'),
    ('tbvmt', '101_12'), ('blp', '101_16'),
]
BANGKE_OVERRIDES = {
    '01ntnn':        ('101_8', '01_NTNN'),
    '01ntnn_tt80':   ('101_8', '01_NTNN'),
    '01ntnn_tt89':   ('101_8', '01_NTNN'),
    '01_2_ht':       ('101_25', '01_DNHT'),
    '01_3a_ht':      ('101_25', '01_DNHT'),
    '01_3b_ht':      ('101_25', '01_DNHT'),
    '01_4_ht':       ('101_25', '01_DNHT'),
    '05dk_th':       ('101_11', '05_DK_TH'),
    '16th_dangkycapma': ('101_11', '16_TH_DKNPT'),
    '20th_dangkycapma': ('101_11', '20_TH_DK_TCT'),
    'khautruthuetncn': ('101_11', 'BK_KHAU_TRU_TNCN'),
    'nq142_cnkd_hkd': ('101_37', 'NQ142_CNKD_HKD'),
    'nq142_cnkd_hkd_tt18': ('101_37', 'NQ142_CNKD_HKD'),
    '01_7gtgt_tt154_masothue': ('101_1', 'PL_01_7_GTGT_TT154'),
    # bảng kê hóa đơn không khớp MaFile (thiếu gạch dưới trong mã / phụ lục của form)
    '01tbac':        ('101_7', '01_TBAC'),
    '01bcac':        ('101_7', 'BC26_AC'),
    '03_7':          ('101_7', 'BK_03_7'),
    'bc26_mau_3_10': ('101_7', 'BC26_AC'),
    'bc26_mau_3_12': ('101_7', 'BC26_AC'),
    # PL01_KHBS — phụ lục khai bổ sung điều chỉnh (root element trong file)
    'xml': ('KHBS', 'PL01_KHBS'),
}

for name, prefix, rest, ver in pending:
    _, ext = split_name(name)
    kind = true_kind(name, ext)
    low = rest.lower()
    # 0) explicit bang-ke / pl overrides
    if low in BANGKE_OVERRIDES:
        gid, folder = BANGKE_OVERRIDES[low]
        assign(gid, folder, ver, name, kind, True)
        continue
    # 1) BCTC: sheet/sets not in MaFile list
    if low.startswith('00_') or 'b01' in low or 'bctc' in low \
            or any(k in low for k in BCTC_KW):
        cand = resolve_stem('1' + low, ver) if re.match(r'^\d', low) else None
        if cand is not None:
            assign('101_6', mafile_by_info[id(cand)], ver, name, kind, True)
        else:
            assign('101_6', bctc_folder(rest), ver, name, kind, True)
        continue
    # 2) phụ lục / bang ke / khbs: first token = form number of parent
    hit = nn_match(rest)
    if hit:
        mf, info = hit
        assign(info['group'], mf, ver, name, kind, True)
        continue
    # 3) tax-type keyword token (old PL: 1A_TNDN, 3A_1_TD_TAIN, pl_*masothue...)
    for kw, gid in TAX_KW:
        if kw in low.split('_'):
            assign(gid, rest.upper(), ver, name, kind, True)
            break
    else:
        # 4) dedicated buckets
        if name.startswith('khbs_'):
            assign('KHBS', rest.upper(), ver, name, kind, True)
        elif name.startswith('bang_ke_'):
            assign('BANG_KE', rest.upper(), ver, name, kind, True)
        elif name.startswith('pl_'):
            assign('PHU_LUC', rest.upper(), ver, name, kind, True)
        else:
            assign('99_KHAC', rest.upper(), ver, name, kind, True)

fallback_assigned = sum(sum(len(f) for f in v.values()) for v in plan.values())

# ---------------------------------------------------------------- report
print('=' * 78)
print(f'catalog menu.xml: {len(mafile)} MaFile, {sum(len(i["versions"]) for i in mafile.values())} (MaFile,version)')
total = 0
for (gid, form), vers in sorted(plan.items(), key=lambda kv: (GROUP_SLUGS.get(kv[0][0], 'zz'), kv[0][1])):
    gslug = GROUP_SLUGS.get(gid, '?')
    for ver, files in sorted(vers.items()):
        for name, kind in sorted(files.items()):
            total += 1
print(f'total matched files: {total}')
print()
print('per group / form:')
per_group = collections.Counter()
for (gid, form), vers in sorted(plan.items()):
    n = sum(len(files) for files in vers.values())
    per_group[(GROUP_SLUGS.get(gid, f'?{gid}'), gid)] += n
for (gslug, gid), n in sorted(per_group.items()):
    gcap = next((i['gcap'] for i in mafile.values() if i['group'] == gid), '')
    nforms = sum(1 for (g, f) in plan if g == gid)
    print(f'  {gslug:28} ({gid:7}) {n:5} files, {nforms:3} forms  - {gcap[:50]}')
print()
print(f'fallback-assigned files (pending pass): {len(pending)}')
for gid in ('KHBS', 'BANG_KE', 'PHU_LUC', '99_KHAC'):
    forms = {f: v for (g, f), v in plan.items() if g == gid}
    if not forms:
        continue
    print(f'  -- {gid}: {sum(sum(len(f) for f in v.values()) for v in forms.values())} files in {len(forms)} folders')
    for f, v in sorted(forms.items()):
        files = [n for vv in v.values() for n in vv]
        print(f'     {f:40} {len(files):4}  e.g. {sorted(files)[:2]}')
print()
print('excluded:')
for k, c in excluded.items():
    print(f'  {k}: {c}')

# ---------------------------------------------------------------- write
if WRITE:
    if os.path.exists(DST):
        sys.exit(f'{DST} already exists — remove it or pick another name first')
    for (gid, form), vers in sorted(plan.items()):
        gslug = GROUP_SLUGS.get(gid, f'99_{form[:24]}')
        for ver, files in sorted(vers.items()):
            vlabel = 'v' + ver_str(ver) if ver not in ('nv',) else 'no-version'
            ddir = os.path.join(DST, gslug, form)
            os.makedirs(ddir, exist_ok=True)
            for name, meta in sorted(files.items()):
                if isinstance(meta, dict):
                    kind, keepname = meta['kind'], meta['keepname']
                else:
                    kind, keepname = meta, False
                ext = os.path.splitext(name)[1]
                if keepname:
                    bname = name[1:] if name.startswith('_') else name
                    dst = os.path.join(ddir, bname)
                else:
                    dst = os.path.join(ddir, f'{form}_{vlabel}_{kind}{ext}')
                # de-dup: two source files with same (form, ver, kind) -> suffix
                base, n = dst, 2
                while os.path.exists(dst):
                    stem2, e2 = os.path.splitext(dst)
                    dst = f'{stem2}-{n}{e2}'
                    n += 1
                shutil.copy2(os.path.join(SRC, name), dst)

    print(f'\nWROTE {sum(sum(len(f) for f in v.values()) for v in plan.values())} files -> {DST}')
