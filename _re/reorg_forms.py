#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reorg_forms.py - map HTKK 5.7.9 template files (Data1.cab -> _extracted/) into _forms/.

Source:  _extracted/  (flat, 4526 files; UNTRUSTED data: only read bytes, never executed)
Target:  _forms/<group>/<form>/        (one folder per template stem = menu FolderName)
         _forms/_shared/<NAME>/        (shared templates: HEADER, IN_TMP)
Outputs: _forms/_manifest.csv          (src -> dst, rule, evidence; one row per placed file)
         _re/pending_mapping.csv       (template files that no rule can place with evidence)

Mapping rules (evidence ladder, first hit wins) - see NOTES.md section 3.9:
  R1 foldername  stem == cabkey(menu.xml Validity/@FolderName)  (runtime opens
                 InterfaceTemplates\\Excel\\<FolderName>_<VersionXML>.xls, WinFnProcess.cs:866)
  R1m mafile     stem == cabkey(menu.xml @MaFile) and no FolderName matches
  R2 datafile    stem == cabkey(menu.xml Sheet/@DataFile) of exactly one menu
  R3 prefixed    bang_ke_/bangke_/khbs_/pl_/bao_cao_ + rest, rest resolved by R1/R1m/R2
  R4 truncated   cab-truncated name: longest key K with name == K + prefix-of("_xml_NNN"|"_NNN");
                 XML root must agree with the tail of the menu Sheet/@XMLnode when it is PL_*
  R5 backup      stem == FolderName in menu_backup2009.xml / menu_backup161118.xml (older catalogues)
  R6 stem        no catalogue entry: folder = the file's own stem (filename evidence only);
                 group by filename keyword (TAX_KW/BCTC_KW) else 99_KHAC
  OVR            BANGKE_OVERRIDES (manual, evidence noted per entry)
  SHARED         header.*, in_tmp.xls -> _forms/_shared/
  PENDING        ambiguous / conflicting evidence -> _re/pending_mapping.csv
OLE workbooks placed by R2/R4/R5 must share sheet names with the menu Sheet IDs of the
target form (BOUNDSHEET parse); a mismatch sends the file to PENDING.

cabkey(): Data1.cab replaces the FIRST character of names that start with a digit by "_"
(01_GTGT -> _1_gtgt, 107_B01 -> _07_b01, 20_TH_DK_TCT -> _0_th_dk_tct); letter-initial
names are kept (bc26_ac). Lower-cased.

Usage (python3 -I):
  python3 -I _re/reorg_forms.py            # DRY RUN: summary + diff against current _forms
  python3 -I _re/reorg_forms.py --write    # incremental apply into _forms + manifest + pending
Env: HTKK_SRC (default /Users/tuanle/codes/HTKK/_extracted), FORMS_DST (default <repo>/_forms)
"""
import os, re, sys, csv, struct, shutil, hashlib, collections, xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.environ.get('HTKK_SRC', '/Users/tuanle/codes/HTKK/_extracted')
DST = os.environ.get('FORMS_DST', os.path.join(ROOT, '_forms'))
PENDING_CSV = os.path.join(ROOT, '_re', 'pending_mapping.csv')
MANIFEST = os.path.join(DST, '_manifest.csv')
WRITE = '--write' in sys.argv
KEEP_IN_DST = {'cong_thuc.txt', '.DS_Store', '_manifest.csv'}   # never touched by apply

# ---------------------------------------------------------------- groups
GROUP_SLUGS = {
    '101_1': '01_GTGT', '101_2': '02_TNDN', '101_3': '03_TAIN', '101_4': '04_TTDB',
    '101_6': '06_BCTC', '101_7': '07_HOA_DON', '101_8': '08_THUE_NHA_THAU',
    '101_9': '09_PHI_LE_PHI', '101_11': '11_TNCN', '101_12': '12_THUE_BVMT',
    '101_15': '15_THUE_KHOAN_THUE_TAI_SAN', '101_16': '16_BIEN_LAI', '101_17': '17_THUY_DIEN',
    '101_18': '18_DAU_KHI', '101_20': '20_TEM_RUOU', '101_24': '24_CO_TUC_LOI_NHUAN',
    '101_25': '25_HOAN_THUE', '101_26': '26_GIA_HAN_NOP_THUE', '101_28': '28_SHTK',
    '101_29': '29_NCC_NUOC_NGOAI', '101_30': '30_MIEN_TIEN_CHAM_NOP', '101_31': '31_DAT_DAI',
    '101_32': '32_TAI_SAN_BAO_DAM', '101_33': '33_DOI_KY_TINH_THUE',
    '101_34': '34_THUONG_MAI_DIEN_TU', '101_35': '35_CHUYEN_DOI_CO_SO_HUU',
    '101_36': '36_THUE_TOI_THIEU_TOAN_CAU', '101_37': '37_HO_CN_KINH_DOANH',
    '101_38': '38_TAI_SAN_MA_HOA',
    'KHBS': 'KHBS', 'BANG_KE': 'BANG_KE', 'PHU_LUC': 'PHU_LUC', '99_KHAC': '99_KHAC',
    '_shared': '_shared', '_ngoai_menu': '_ngoai_menu',
}

def cabkey(name):
    n = name.lower()
    return '_' + n[1:] if n[:1].isdigit() else n

# ---------------------------------------------------------------- catalogue (menu.xml)
def load_menu(path):
    root = ET.parse(path).getroot()
    menus = list(root.iter('Menu'))
    byid = {m.get('ID'): m for m in menus}
    def group_of(m):
        cur = m
        while cur is not None:
            pid = cur.get('ParentID', '')
            if re.match(r'^10[0-4]_\d+$', pid):
                return pid
            cur = byid.get(pid)
        return None
    out = []
    for m in menus:
        if not m.get('MaFile'):
            continue
        fns = [v.get('FolderName') for v in m.findall('Validity') if v.get('FolderName')]
        if not fns:
            continue
        sheets = list(m.iter('Sheet'))
        out.append({
            'id': m.get('ID'), 'mafile': m.get('MaFile'), 'folder': fns[0],
            'ver': (m.get('VersionXML') or '').replace('.', ''), 'group': group_of(m),
            'sheet_ids': [s.get('ID', '') for s in sheets],
            'cls': m.get('Class', ''), 'caption': m.get('Caption', ''),
            'sheet_names': [s.get('SheetName', '') for s in sheets],
            'datafiles': [(s.get('DataFile', ''), s.get('XMLnode', '') or '') for s in sheets],
        })
    return out

CAT = load_menu(os.path.join(SRC, 'menu.xml'))

# ---------------------------------------------------------------- R1b: ShowBangKe("<file>") in code
# Decompiled code (read as text only) names the import/print workbook each form opens:
#  - business class <Class>Nckt<ver>.cs  -> menu with that Class (+ VersionXML)
#  - frmHTKKSystems.cs  case "102_x": ShowBangKe(...)  -> menu 102_x Caption code (e.g. 03-7/TNDN)
#    matched against Sheet/@SheetName and Caption of form menus.
DECOMP = os.environ.get('HTKK_DECOMPILED', '/Users/tuanle/codes/HTKK/_decompiled')
MENU_ALL = {m.get('ID'): m for m in ET.parse(os.path.join(SRC, 'menu.xml')).getroot().iter('Menu')}

def _ncode(x):
    return re.sub(r'\s+', '', x.upper()).replace('Đ', 'D')

def _caption_codes(cap):
    return [_ncode(c) for c in re.findall(r'[\w.\-]+/[\w\-]+(?:/[\w\-]+)?', cap)]

def _tt(x):
    return set(re.findall(r'tt\d+|nd\d+', x.lower()))

SBK_TRUNC = 28   # Data1.cab keeps the first 28 chars of a lower-cased name (+1 digit on collision)

def sbk_scan():
    calls = []    # (file_stem_lower, src_rel, line, kind, key)
    for dp, dn, fn in os.walk(DECOMP):
        for f in fn:
            if not f.endswith('.cs'):
                continue
            path = os.path.join(dp, f)
            try:
                lines = open(path, encoding='utf-8', errors='replace').read().split('\n')
            except OSError:
                continue
            case = None
            for i, ln in enumerate(lines, 1):
                mc = re.search(r'case "([^"]+)":', ln)
                if mc:
                    case = mc.group(1)
                m = re.search(r'ShowBangKe\("([^"]+)"\)', ln)
                if not m or 'void ShowBangKe' in ln:
                    continue
                stem = m.group(1).lower()   # full file name as called (lower)
                rel = os.path.relpath(path, DECOMP)
                if f == 'frmHTKKSystems.cs':
                    calls.append((stem, rel, i, 'menu', case))
                else:
                    calls.append((stem, rel, i, 'class', f[:-3]))
    return calls

def sbk_resolve(stem, kind, key):
    """-> (entries list, evidence text); entries empty = standalone menu or unresolved."""
    if kind == 'class':
        mc = re.match(r'^(.*?)Nckt(\d+)$', key)
        cls, ver = (mc.group(1), mc.group(2)) if mc else (key, '')
        es = [e for e in CAT if e['cls'] == cls]
        if len({e['folder'] for e in es}) > 1:
            ev = [e for e in es if e['ver'] == ver]
            es = ev or es
        return es, 'Class=%s' % cls
    m = MENU_ALL.get(key)
    if m is None:
        return [], 'menu %s khong co' % key
    cap = m.get('Caption', '')
    codes = _caption_codes(cap)
    blp = 'BLP' in cap.upper()
    es = []
    for e in CAT:
        names = [_ncode(x) for x in e['sheet_names'] + [e['caption']]]
        if any(c and any(re.search(r'(^|[^\w])' + re.escape(c) + r'($|[^\w])', n) for n in names) for c in codes):
            if ('BLP' in e['folder'].upper()) == blp:
                es.append(e)
    if len({e['folder'] for e in es}) > 1:
        t = _tt(stem)
        es2 = [e for e in es if _tt(e['folder']) == t]
        es = es2 or es
    return es, 'menu %s "%s"' % (key, cap)

SBK = {}   # cab-truncated lower file name -> dict(cands, ev, menus, full)
for stem, rel, line, kind, key in sbk_scan():
    es, ev = sbk_resolve(os.path.splitext(stem)[0], kind, key)
    d = SBK.setdefault(stem[:SBK_TRUNC], {'cands': {}, 'ev': [], 'menus': [], 'full': set()})
    d['full'].add(stem)
    d['ev'].append('%s:%d ShowBangKe; %s' % (rel, line, ev))
    for e in es:
        d['cands'][e['folder']] = e
        d.setdefault('byf', {}).setdefault(stem, {})[e['folder']] = e
    d.setdefault('byf', {}).setdefault(stem, {})
    if kind == 'menu':
        d['menus'].append(key)
BACKUP = []
for bf in ('menu_backup161118.xml', 'menu_backup2009.xml'):
    try:
        BACKUP += [dict(e, source=bf) for e in load_menu(os.path.join(SRC, bf))]
    except ET.ParseError:
        pass

FK = collections.defaultdict(list)   # cabkey(FolderName) -> entries
MK = collections.defaultdict(list)   # cabkey(MaFile)     -> entries
DK = collections.defaultdict(list)   # cabkey(DataFile)   -> (entry, xmlnode)
for e in CAT:
    FK[cabkey(e['folder'])].append(e)
    MK[cabkey(e['mafile'])].append(e)
for e in CAT:
    for df, xn in e['datafiles']:
        k = cabkey(df)
        if not df or df.upper() == 'KHBS' or k in FK or k in MK:
            continue
        if all(x[0] is not e for x in DK[k]):
            DK[k].append((e, xn))
BK = collections.defaultdict(list)
for e in BACKUP:
    k = cabkey(e['folder'])
    if k not in FK and k not in MK and all(x['folder'] != e['folder'] for x in BK[k]):
        BK[k].append(e)

# ---------------------------------------------------------------- content sniffing
def sniff(path):
    with open(path, 'rb') as f:
        b = f.read(4096)
    if b[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        return 'OLE', None
    if b[:2] == b'PK':
        return 'ZIP', None
    if b[:2] == b'MZ':
        return 'PE', None
    txt = None
    if b[:2] in (b'\xff\xfe', b'\xfe\xff'):
        txt = b.decode('utf-16', 'replace')
    else:
        s = b.lstrip(b'\xef\xbb\xbf \r\n\t')
        if s[:1] == b'<':
            txt = s.decode('utf-8', 'replace')
    if txt is None:
        return 'OTHER', None
    m = re.search(r'<(?![?!])([A-Za-z_][\w.\-:]*)', txt)
    return 'XML', (m.group(1).split(':')[-1] if m else '')

TEMPLATE_ROOT = re.compile(r'^(Sections|HSoThueDTu|TKhai(_.*)?|PL.*|BKe.*|ToKhai.*|Workbook)$', re.I)

def ole_sheet_names(path):
    """Sheet names from BIFF BOUNDSHEET (0x85) records; pure parse, no execution."""
    d = open(path, 'rb').read()
    if d[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        return None
    try:
        ssz = 1 << struct.unpack_from('<H', d, 30)[0]
        mssz = 1 << struct.unpack_from('<H', d, 32)[0]
        nfat, dir_start = struct.unpack_from('<II', d, 44)
        mini_cut, mfat_start, _, difat_start, ndifat = struct.unpack_from('<IIIII', d, 56)
        difat = list(struct.unpack_from('<109I', d, 76))
        s = difat_start
        for _ in range(ndifat):
            vals = struct.unpack_from('<%dI' % (ssz // 4), d, (s + 1) * ssz)
            difat += vals[:-1]
            s = vals[-1]
        fat = []
        for sec in difat[:nfat]:
            fat += struct.unpack_from('<%dI' % (ssz // 4), d, (sec + 1) * ssz)
        def chain(st):
            out, seen = [], set()
            while st < 0xFFFFFFFA and st not in seen and st < len(fat):
                seen.add(st); out.append(st); st = fat[st]
            return out
        def read(st):
            return b''.join(d[(x + 1) * ssz:(x + 2) * ssz] for x in chain(st))
        dirb = read(dir_start)
        ents = []
        for i in range(len(dirb) // 128):
            e = dirb[i * 128:(i + 1) * 128]
            nl = struct.unpack_from('<H', e, 64)[0]
            ents.append((e[:max(nl - 2, 0)].decode('utf-16le', 'replace'), e[66],
                         struct.unpack_from('<I', e, 116)[0], struct.unpack_from('<I', e, 120)[0]))
        wb = next((e for e in ents if e[0] in ('Workbook', 'Book') and e[1] == 2), None)
        if wb is None:
            return None
        if wb[3] < mini_cut:
            ms = read(ents[0][2])
            mfat = []
            for x in chain(mfat_start):
                mfat += struct.unpack_from('<%dI' % (ssz // 4), d, (x + 1) * ssz)
            parts, st, seen = [], wb[2], set()
            while st < 0xFFFFFFFA and st not in seen and st < len(mfat):
                seen.add(st); parts.append(ms[st * mssz:(st + 1) * mssz]); st = mfat[st]
            data = b''.join(parts)[:wb[3]]
        else:
            data = read(wb[2])[:wb[3]]
        names, pos = [], 0
        while pos + 4 <= len(data):
            t, l = struct.unpack_from('<HH', data, pos)
            b = data[pos + 4:pos + 4 + l]
            if t == 0x0085 and len(b) >= 8:
                cch, hi = b[6], b[7] & 1
                names.append(b[8:8 + cch * (2 if hi else 1)].decode('utf-16le' if hi else 'latin-1'))
            elif t == 0x000A:
                break
            pos += 4 + l
        return names
    except (struct.error, IndexError):
        return None

def sheet_jaccard(names, menu_ids):
    a = {norm_sheet(x) for x in names if 'HEADER' not in x.upper() and 'KHBS' not in x.upper()}
    b = {norm_sheet(x) for x in menu_ids if 'KHBS' not in x.upper()}
    if not a or not b:
        return None
    return len(a & b) / len(a | b)

def lcs_len(a, b):
    best = 0
    for i in range(len(a)):
        for j in range(i + best + 1, len(a) + 1):
            if a[i:j] in b:
                best = j - i
            else:
                break
    return best

def norm_root(r):
    return re.sub(r'^(PLUC_?|PL_?)', '', r.upper())

def norm_sheet(s):
    return re.sub(r'[\s\-_/]+', '_', s.strip().upper())

def sheet_overlap(names, menu_ids):
    a = {norm_sheet(x) for x in names if 'HEADER' not in x.upper() and 'KHBS' not in x.upper()}
    b = {norm_sheet(x) for x in menu_ids if 'KHBS' not in x.upper()}
    if not a or not b:
        return None
    return len(a & b) / min(len(a), len(b))

# ---------------------------------------------------------------- kind labels (unchanged naming)
EXT_KIND = {'xls': 'mau', 'xls1': 'mau-b', 'xlsx': 'mau', 'xml': 'mau-xml', 'xm': 'mau-xml',
            'x': 'mau-xml?', 'xs': 'xsd', 'xl': 'mau', 'xl1': 'mau-b', 'xml1': 'mau-xml',
            'xml9': 'mau-xml', 'xsd': 'xsd', '': '?'}

def true_kind(magic, ext):
    if magic in ('OLE', 'ZIP'):
        return 'mau'
    if ext in ('x', 'xm', 'xl', 'xs', ''):
        return 'mau-xml' if magic == 'XML' else 'mau?'
    return EXT_KIND.get(ext, '?')

def ver_str(d):
    return '.'.join(d) if len(d) == 3 else d

# ---------------------------------------------------------------- legacy tables (R6 / OVR)
BCTC_KW = ('cdkt', 'sxkd', 'lctttt', 'lcttgt', 'lctt', 'thtc', 'kqhdkd', 'bctc', 'cdtk', 'b01')
TAX_KW = [('tain', '101_3'), ('tndn', '101_2'), ('tncn', '101_11'), ('ntnn', '101_8'),
          ('gtgt', '101_1'), ('ttdb', '101_4'), ('dnht', '101_25'), ('tmdn', '101_31'),
          ('sddnn', '101_31'), ('tbvmt', '101_12'), ('blp', '101_16')]
# manual decisions kept from the 2026-10-08 run; applied only when R1-R5 find nothing.
BANGKE_OVERRIDES = {
    '01ntnn': ('101_8', '01_NTNN'), '01ntnn_tt80': ('101_8', '01_NTNN'),
    '01ntnn_tt89': ('101_8', '01_NTNN'),
    '01_2_ht': ('101_25', '01_DNHT'), '01_3a_ht': ('101_25', '01_DNHT'),
    '01_3b_ht': ('101_25', '01_DNHT'), '01_4_ht': ('101_25', '01_DNHT'),
    '05dk_th': ('101_11', '05_DK_TH'), '16th_dangkycapma': ('101_11', '16_TH_DKNPT'),
    '20th_dangkycapma': ('101_11', '20_TH_DK_TCT'),
    'khautruthuetncn': ('101_11', 'BK_KHAU_TRU_TNCN'),
    'nq142_cnkd_hkd': ('101_37', 'NQ142_CNKD_HKD'), 'nq142_cnkd_hkd_tt18': ('101_37', 'NQ142_CNKD_HKD'),
    '01_7gtgt_tt154_masothue': ('101_1', 'PL_01_7_GTGT_TT154'),
    '01tbac': ('101_7', '01_TBAC'), '01bcac': ('101_7', 'BC26_AC'), '03_7': ('101_7', 'BK_03_7'),
    'bc26_mau_3_10': ('101_7', 'BC26_AC'), 'bc26_mau_3_12': ('101_7', 'BC26_AC'),
    'xml': ('KHBS', 'PL01_KHBS'),
}
SHARED = {'header': 'HEADER', 'in_tmp': 'IN_TMP'}
# candidate hints for pending rows (not used for placement)
PENDING_HINTS = {
    '_5kk_tncn.xml': '05_KK_TNCN (ten 5kk_tncn; root Workbook, khong co khoa trong menu)',
    '_2_lncl_vsp_tt80_251.xls': 'khong xac dinh; sheet 01/01-1/01-2/KHBS khong khop 02_PTHU_VSP_TT80',
}
# per-file notes for _re/pending_mapping.csv (what was checked, why no rule applies)
PENDING_NOTES = {
    '_3_tain.xls': 'BIFF tho khong co vo OLE; khong co FolderName/MaFile/DataFile nao co dang ?3_TAIN',
    '_5kk_tncn.xml': 'SpreadsheetML (root Workbook); khong co khoa ?5KK_TNCN trong menu.xml',
    '_9_01_cdkt_tl__hn_xml_26.xml': 'ten loi "__" + version 2 chu so; gan voi DataFile 49_01_CDKT_TL_HN (menu H9 49_BCTC_TL_HN) nhung khong khop dung khoa',
    'bangkekhautru_01ttdb.xls': 'tien to bangkekhautru_ ngoai quy uoc; "01ttdb" khong khop MaFile 01_TTDB (thieu "_"), 2 ung vien FolderName 01_TTDB_*',
    'bangkekhautru_01ttdb_tt80.xl': 'tien to bangkekhautru_ ngoai quy uoc; "01ttdb_tt80" khong khop khoa nao; ten bi cat (.xl)',
    'template.xls': 'BIFF tho, ten chung khong mang ma form',
}
SHARED_XSD = {'tkhaithue', 'tkhaithuenccnn', 'tkhaithue_hddt', 'tkhaithue_dbhc', 'tkhaidkythue',
              'xmldsig_core_schema'}   # base schemas imported by every form XSD
# root=Sections but app configuration (convert/validate tables), not a form template
CONFIG_SECTIONS = {'config_header.xml', 'convertdatacommon.xml', 'convertversion.xml',
                   'createdatafilekhbsmiss.xml', 'replacedatafile.xml', 'validcombobox.xml'}
CONFIG_NAMES = {'menu.xml', 'versiontokhai.xml', 'message.xml', 'config.xml'}

# ---------------------------------------------------------------- resolution helpers
FULL_RE = re.compile(r'^(?P<stem>.+?)(?P<xml>_xml\d?)?_(?P<ver>\d{3,5})(?:_|_bu)?$')
PREFIXES = ('bang_ke_', 'bangke_', 'khbs_', 'pl_', 'bao_cao_')
TRUNC_TAIL = re.compile(r'^(?:_(?:x(?:m(?:l(?:\d|_(?P<v1>\d{0,5}))?)?)?)?)?$|^_(?P<v2>\d{1,5})_?$')

def lookup(key):
    """key (already cab-form or plain lower) -> list of (rule, entry, xmlnode)."""
    if key in FK:
        return [('R1', e, None) for e in FK[key]]
    if key in MK:
        return [('R1m', e, None) for e in MK[key]]
    if key in DK:
        return [('R2', e, xn) for e, xn in DK[key]]
    return []

def lookup_any(rest):
    """rest of a prefixed name: original spelling (not cab-mangled) or cab form."""
    for k in (rest.lower(), cabkey(rest)):
        hit = lookup(k)
        if hit:
            return hit
    return []

ALL_KEYS = sorted(set(FK) | set(MK) | set(DK), key=len, reverse=True)

def pick(hits, ver, sheets):
    """disambiguate several catalogue hits by version, then by sheet names."""
    if len({id(h[1]) for h in hits}) == 1:
        return hits[0], ''
    if ver:
        hv = [h for h in hits if h[1]['ver'] == ver]
        if len(hv) == 1:
            return hv[0], 'version'
    if all(h[0] == 'R1m' for h in hits) and len({h[1]['mafile'] for h in hits}) == 1:
        return hits[0], 'cung MaFile ' + hits[0][1]['mafile']
    if sheets:
        sc = [(sheet_jaccard(sheets, h[1]['sheet_ids']) or 0, h) for h in hits]
        sc.sort(key=lambda x: -x[0])
        if sc[0][0] >= 0.5 and (len(sc) == 1 or sc[1][0] < sc[0][0]):
            return sc[0][1], 'sheet'
    return None, 'ambiguous:' + '|'.join(sorted({h[1]['folder'] for h in hits}))

# ---------------------------------------------------------------- scan
records = []       # dicts: src, magic, root, kind, ext, gid, folder, ver, rule, evidence, keepname
pending = []       # dicts: src, magic, root, reason, candidates
excluded = collections.Counter()
known_vers = collections.defaultdict(set)   # folder -> versions seen in full names

files = sorted(n for n in os.listdir(SRC) if os.path.isfile(os.path.join(SRC, n)))
meta = {}
for name in files:
    magic, root = sniff(os.path.join(SRC, name))
    disp = name.rstrip('.')
    stem0, ext = os.path.splitext(disp)
    ext = ext.lstrip('.').lower()
    meta[name] = (magic, root, stem0, ext)

def in_scope(name):
    magic, root, stem0, ext = meta[name]
    low = name.lower()
    if low.startswith('catalogue'):
        return False, 'catalogue_* (danh muc)'
    if low.startswith('dm_') or low.startswith('dmngayle'):
        return False, 'dm_* (danh muc)'
    if low in CONFIG_NAMES or low in CONFIG_SECTIONS or low.startswith('menu_backup'):
        return False, 'catalogue menu/config'
    if magic in ('OLE', 'ZIP'):
        if ext in ('doc',):
            return False, 'OLE khac (.doc)'
        return True, ''
    if magic == 'XML':
        if root == 'schema':
            return True, ''          # xsd: mapped, reported separately
        if root and TEMPLATE_ROOT.match(root):
            return True, ''
        return False, 'XML cau hinh (root=%s)' % (root or '?')
    if ext == 'xls':
        return True, ''              # raw BIFF stream without OLE wrapper
    return False, 'khac (%s)' % magic

def add(name, gid, folder, ver, rule, evidence, keepname):
    magic, root, stem0, ext = meta[name]
    note = globals().get('sbk_note', {}).get(name)
    if note and 'ShowBangKe' not in evidence:     # ShowBangKe target placed by a filename rule
        evidence += '; ' + note
    mnote = globals().get('MANUAL_NOTES', {}).get(name)
    if mnote:
        evidence += '; ' + mnote
    records.append(dict(src=name, magic=magic, root=root or '', ext=ext, gid=gid, folder=folder,
                        ver=ver or 'nv', rule=rule, evidence=evidence, keepname=keepname,
                        kind=true_kind(magic, ext)))

def to_pending(name, reason, cands=''):
    magic, root, stem0, ext = meta[name]
    pending.append(dict(src=name, magic=magic, root=root or '', reason=reason, candidates=cands))

sheets_cache = {}
def sheets_of(name):
    if name not in sheets_cache:
        sheets_cache[name] = ole_sheet_names(os.path.join(SRC, name)) if meta[name][0] == 'OLE' else None
    return sheets_cache[name]

def place_hit(name, hit, how, ver, rule_override=None):
    rule, e, xn = hit
    rule = rule_override or rule
    ev = 'menu %s MaFile=%s FolderName=%s VersionXML=%s' % (e['id'], e['mafile'], e['folder'], e['ver'])
    if rule in ('R2',) or (rule_override and rule_override.startswith('R3') and hit[0] == 'R2'):
        ev += '; Sheet/@DataFile'
    if how:
        ev += '; chon theo ' + how
    # sheet gate for OLE placed by indirect rules
    sh = sheets_of(name)
    if sh is not None and rule not in ('R1', 'R1m') and not rule.startswith('R3'):
        ov = sheet_overlap(sh, e['sheet_ids'])
        if ov is not None and ov < 0.5:
            to_pending(name, 'sheet_lech_menu (overlap=%.2f)' % ov, e['folder'])
            return
        if ov is not None:
            ev += '; sheet overlap=%.2f' % ov
    keep = rule != 'R1' or name.startswith(PREFIXES)
    folder = (e['mafile'] if rule.split('/')[-1] == 'R1m' else e['folder']).upper()
    add(name, e['group'], folder, ver, rule, ev, keep)

deferred = []   # truncated / stem-only names, resolved after full names (need known_vers)

SBK_PREFIX = re.compile(r'^(bang_?ke_|bao_cao_|bangkekhautru_|pl_|khbs_)')
sbk_note = {}   # extracted name -> ShowBangKe evidence when no menu form resolves (falls to R5x)

def sbk_match(name):
    """extracted name -> (key, entry) when it is a ShowBangKe target: equal to the 28-char cab
    name, or (only when the called name is longer than 28) that name + 1 collision digit."""
    low = name.lower()
    if low in SBK:
        return low, SBK[low]
    k = low[:SBK_TRUNC]
    if len(low) == SBK_TRUNC + 1 and low[-1].isdigit() and k in SBK \
            and all(len(f) > SBK_TRUNC for f in SBK[k]['full']):
        return k, SBK[k]
    return None, None

def sbk_place(name, sb):
    """True when the file was placed (R1b) or sent to pending; False = fall through."""
    ev = ' | '.join(sorted(set(sb['ev'])))
    fulls = sorted(sb['full'])
    if len(fulls) > 1:   # cab collision: several called names share one 28-char cab name
        cl = ['%s -> %s' % (f, ','.join(sorted(sb['byf'][f])) or 'khong co form trong menu') for f in fulls]
        to_pending(name, 'ShowBangKe: ten cab bi cat trung %d ten goi, khong xac dinh tep nao' % len(fulls),
                   '; '.join(cl) + ' | ' + ev)
        return True
    fstem = os.path.splitext(fulls[0])[0]
    cands, how = dict(sb['cands']), ''
    if len(cands) > 1:   # same Class shared by variants: drop _VIEW, then keep equal tt/nd tokens
        nv = {k: v for k, v in cands.items() if not k.upper().endswith('_VIEW')} or cands
        tt = {k: v for k, v in nv.items() if _tt(k) == _tt(fstem)} or nv
        if len(tt) < len(cands):
            how = ' (loai bien the _VIEW/khac tt)'
        cands = tt
    if len(cands) > 1:
        rest = SBK_PREFIX.sub('', fstem)
        ex = {k: v for k, v in cands.items() if cabkey(k) in (rest, cabkey(rest))}
        tok = set(fstem.split('_'))
        sub = {k: v for k, v in cands.items()
               if set(k.lower().split('_')) <= tok and _tt(k) == _tt(fstem)}
        if len(ex) == 1:
            cands, how = ex, how + '; ten goi bo tien to == FolderName %s' % list(ex)[0]
        elif len(sub) == 1:
            cands, how = sub, how + '; chi FolderName %s co moi token (ke ca tt/nd) nam trong ten goi' % list(sub)[0]
    if len(cands) == 1:
        e = list(cands.values())[0]
        add(name, e['group'], e['folder'].upper(), None, 'R1b',
            ev + how + ' -> FolderName=%s (menu %s)' % (e['folder'], e['id']), True)
        return True
    if len(cands) > 1:
        to_pending(name, 'ShowBangKe: nhieu ung vien form', ','.join(sorted(cands)) + ' | ' + ev)
        return True
    sbk_note[name] = ev + ('; goi tu menu he thong %s, khong co FolderName' % ','.join(sorted(set(sb['menus'])))
                           if sb['menus'] else '; Class khong co menu MaFile/FolderName trong menu.xml')
    return False

# Ngoại lệ đặt tay có bằng chứng tệp:dòng (xét trước mọi quy tắc). Giá trị: (gid, thư mục, phiên bản, rule, bằng chứng).
MANUAL_PLACEMENTS = {
    '_9_01_cdkt_tl__hn_xml_26.xml': (
        '101_6', '49_BCTC_TL_HN', '226', 'OVR',
        'OVR: ten cab bi cat 28 ky tu (_9_01_cdkt_tl__hn_xml_26.xml); menu.xml:871 Sheet ID=01 '
        'DataFile=49_01_CDKT_TL_HN cua MaFile=49_BCTC_TL_HN; sha256 trung byte voi '
        '06_BCTC/49_BCTC_TL_HN/9_01_cdkt_tl_hn_xml_226.xml (7230b0d3...) nen phien ban 226'),
    'pl_01_tndn_tt199_masothue.xl': (
        '_ngoai_menu', 'PL_01_TNDN_TT199_MASOTHUE', None, 'R5x',
        'R5x: stem=01_tndn_tt199_masothue, phien_ban=khong; khong co FolderName/MaFile/DataFile trong menu.xml; '
        'chi menu he thong 102_19 goi: htkk/HTKK/frmHTKKSystems.cs:872 ShowBangKe("PL_01_TNDN_TT199_MaSoThue.xls"); '
        'cung quy uoc voi 6 tep PL_*_MASOTHUE/BKGIAHAN_TT170_MASOTHUE da o _ngoai_menu (frmHTKKSystems.cs:854-875); '
        'ung vien 01_TKTT_IIR,01_TKTT_QDMTT,01_TNDN_DK,01_TNDN_TT80_DK chi trung ten, khong co loi goi'),
}
# Ghi chú thêm vào bằng chứng của tệp đã đặt theo quy tắc (không đổi vị trí).
MANUAL_NOTES = {
    'copy_of_01_tbvmt_201.xls':
        'giu o _ngoai_menu: bo tien to copy_of_ thi stem la FolderName 01_TBVMT (menu.xml) nhung sha256 khac '
        'ban 01_TBVMT v2.0.1 (_forms/12_THUE_BVMT/01_TBVMT/01_TBVMT_v2.0.1_mau.xls sha aa9c0e1d..., ban sao 6987581e...), nen la ban sao khac noi dung, khong phai mau cua form',
}

for name in files:
    if name in MANUAL_PLACEMENTS:
        g_, f_, v_, r_, e_ = MANUAL_PLACEMENTS[name]
        add(name, g_, f_, v_, r_, e_, True)
        continue
    key_, sb = sbk_match(name)          # (C) ShowBangKe target: looked up before in_scope
    if sb is not None and sbk_place(name, sb):
        continue
    ok, why = in_scope(name)
    if not ok and name not in sbk_note:
        excluded[why] += 1
        continue
    magic, root, stem0, ext = meta[name]
    low = stem0.lower()
    if low in SHARED:
        add(name, '_shared', SHARED[low], None, 'SHARED',
            'tep dung chung, khong gan ma form' + ('; HEADER.xls WinFnProcess.cs:864' if low == 'header' else ''),
            True)
        continue
    if low in SHARED_XSD:
        add(name, '_shared', 'XSD_CHUNG', None, 'SHARED', 'schema goc import boi XSD tung mau (NOTES.md muc 2)', True)
        continue
    prefix = next((p for p in PREFIXES if low.startswith(p)), None)
    if prefix:
        rest0 = low[len(prefix):]
        m = FULL_RE.match(rest0)
        ver, rest = (m.group('ver'), m.group('stem')) if m else (None, re.sub(r'_xml\d?$', '', rest0))
        hits = lookup_any(rest)
        if hits:
            h, how = pick(hits, ver, sheets_of(name))
            if h:
                place_hit(name, h, how, ver, rule_override='R3/' + h[0])
                if m:
                    known_vers[h[1]['folder']].add(ver)
                continue
            to_pending(name, how, rest)
            continue
        deferred.append((name, prefix, rest, ver))
        continue
    m = FULL_RE.match(low)
    if m:
        stem, ver = m.group('stem'), m.group('ver')
        hits = lookup(stem)
        if hits:
            h, how = pick(hits, ver, sheets_of(name))
            if h:
                place_hit(name, h, how, ver)
                known_vers[h[1]['folder']].add(ver)
                continue
            to_pending(name, how, stem)
            continue
        deferred.append((name, '', stem, ver))
        continue
    hits = lookup(low)                    # versionless exact key (e.g. _9_bctc_dd.xls)
    if hits:
        h, how = pick(hits, None, sheets_of(name))
        if h:
            place_hit(name, h, how, None, rule_override='R5v/' + h[0])
            continue
        to_pending(name, how, low)
        continue
    deferred.append((name, '', None, None))

# R4 truncated + R5 backup, then R5n / R5x for names that carry a stem but no menu form
VER2_RE = re.compile(r'^(?P<stem>.+?)(?:_xml\d?)?_(?P<ver>\d{2})_?$')
r6_cands = []
for name, prefix, stem, ver in deferred:
    magic, root, stem0, ext = meta[name]
    low = stem0.lower()
    if stem is None:
        # R4: truncated name
        done = False
        for k in ALL_KEYS:
            if not low.startswith(k):
                continue
            tm = TRUNC_TAIL.match(low[len(k):])
            if not tm:
                continue
            hits = lookup(k)
            h, how = pick(hits, None, sheets_of(name))
            if not h:
                to_pending(name, 'truncated ' + how, k)
                done = True
                break
            e, xn = h[1], h[2]
            if root and root.upper().startswith('PL') and xn:
                tail = xn.split('/')[-1].upper()
                if max(lcs_len(norm_root(root), norm_root(tail)), lcs_len(norm_root(root), k.upper())) < 4:
                    to_pending(name, 'truncated: root %s khac XMLnode %s' % (root, tail), e['folder'])
                    done = True
                    break
            part = tm.group('v1') or tm.group('v2') or ''
            kv = sorted(v for v in known_vers.get(e['folder'], ()) if part and v.startswith(part))
            v = kv[0] if len(kv) == 1 else None
            place_hit(name, h, (how + '; ' if how else '') + 'ten bi cat, khop tien to "%s"%s' % (
                k, ', version %s tu %s' % (v, part) if v else ''), v, rule_override='R4/' + h[0])
            done = True
            break
        if done:
            continue
        m2 = VER2_RE.match(low)
        if m2:                               # 2-digit version suffix
            r6_cands.append((name, '', m2.group('stem'), m2.group('ver')))
        else:                                # versionless, no key: R5x only next to a versioned sibling
            r6_cands.append((name, '', low, None))
        continue
    key = cabkey(stem) if not prefix else None
    # R5: older catalogue
    bhits = BK.get(key or '', []) or (BK.get(stem.lower(), []) + BK.get(cabkey(stem), []) if prefix else [])
    if bhits and len({b['folder'] for b in bhits}) == 1 and bhits[0]['group'] in GROUP_SLUGS:
        b = bhits[0]
        add(name, b['group'], b['folder'], ver, 'R5', '%s menu %s FolderName=%s' % (b['source'], b['id'], b['folder']), True)
        continue
    r6_cands.append((name, prefix, stem, ver))

# R5n: stem == exactly one FolderName/MaFile (cab key; trailing "_" dropped), only the version is not in menu.
# R5x: stem + version recognised but no menu form -> _forms/_ngoai_menu/<STEM>/ (one folder per stem).
NEAR_KEYS = {}
for e in CAT:
    NEAR_KEYS.setdefault(cabkey(e['folder']), e['folder'])
    NEAR_KEYS.setdefault(cabkey(e['mafile']), e['folder'])

def nearest(stem):
    import difflib
    k = cabkey(stem)
    base = re.sub(r'_(tt|nd)\d+', '', k)
    out = [NEAR_KEYS[x] for x in NEAR_KEYS if re.sub(r'_(tt|nd)\d+', '', x) == base]
    out += [NEAR_KEYS[x] for x in difflib.get_close_matches(k, list(NEAR_KEYS), n=3, cutoff=0.8)]
    return ','.join(sorted(dict.fromkeys(out)))

def x_folder(stem):
    return stem.strip('_').upper()

r5n_cnt = 0
base_x = {}       # cab stem -> folder name under _ngoai_menu
later = []
for name, prefix, stem, ver in r6_cands:
    if prefix:
        later.append((name, prefix, stem, ver))
        continue
    s2 = stem.rstrip('_')
    fol = sorted({e['folder'] for k in {stem, s2} for e in FK.get(k, []) + MK.get(k, [])})
    if len(fol) == 1 and ver:
        e = (FK.get(s2) or FK.get(stem) or MK.get(s2) or MK.get(stem))[0]
        if ver != e['ver']:
            add(name, e['group'], e['folder'].upper(), ver, 'R5n',
                'R5n: stem=%s, phien_ban=%s ngoai menu; menu %s FolderName=%s VersionXML=%s'
                % (s2, ver_str(ver), e['id'], e['folder'], e['ver']), True)
            r5n_cnt += 1
            continue
    if len(fol) > 1:
        to_pending(name, 'stem khop nhieu form', ','.join(fol))
        continue
    if ver or name in sbk_note:
        base_x.setdefault(cabkey(s2), x_folder(s2))
    later.append((name, prefix, stem, ver))

for name, prefix, stem, ver in later:
    s2 = stem.rstrip('_')
    if prefix:
        rk = cabkey(s2)
        if rk in base_x:
            folder, how = base_x[rk], 'tien to %s + stem cua thu muc %s' % (prefix, base_x[rk])
        elif s2 in ('', 'xml'):
            to_pending(name, 'ten chung (%s + %s), khong mang stem form' % (prefix, s2 or '""'), '')
            continue
        else:
            folder, how = x_folder(prefix + s2), 'stem gom tien to'
    elif cabkey(s2) in base_x:
        folder = base_x[cabkey(s2)]
        how = '' if ver or name in sbk_note else 'khong phien ban, cung stem voi tep co phien ban'
    else:
        to_pending(name, 'khong nhan dang duoc stem/phien ban (khong khop khoa menu, khong co tep cung stem co phien ban)',
                   PENDING_HINTS.get(name.lower(), '') or nearest(s2))
        continue
    ov = BANGKE_OVERRIDES.get(s2)
    near = nearest(s2) or (ov[1] if ov else '')
    ev = 'R5x: stem=%s, phien_ban=%s; khong co FolderName/MaFile/DataFile trong menu.xml%s%s%s' % (
        s2, ver_str(ver) if ver else 'khong', '; ' + how if how else '',
        '; ung vien gan nhat: ' + near if near else '', '; ' + sbk_note[name] if name in sbk_note else '')
    add(name, '_ngoai_menu', folder, ver, 'R5x', ev, True)

# ---------------------------------------------------------------- destination names
plan = collections.defaultdict(lambda: collections.defaultdict(dict))
for r in records:
    plan[(r['gid'], r['folder'])][r['ver']][r['src']] = r
dest = {}
used = set()
for (gid, form), vers in sorted(plan.items()):
    gslug = GROUP_SLUGS.get(gid, '99_KHAC')
    for ver, fs in sorted(vers.items()):
        vlabel = 'v' + ver_str(ver) if ver != 'nv' else 'no-version'
        for name, r in sorted(fs.items()):
            ddir = os.path.join(gslug, form)
            ext = os.path.splitext(name)[1]
            if r['keepname']:
                bname = name[1:] if name.startswith('_') else name
                d = os.path.join(ddir, bname)
            else:
                d = os.path.join(ddir, '%s_%s_%s%s' % (form, vlabel, r['kind'], ext))
            base, n = d, 2
            while d in used:
                st, e2 = os.path.splitext(base)
                d = '%s-%d%s' % (st, n, e2)
                n += 1
            used.add(d)
            dest[name] = d
            r['dst'] = d

# ---------------------------------------------------------------- report
def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for blk in iter(lambda: f.read(1 << 20), b''):
            h.update(blk)
    return h.hexdigest()

rule_cnt = collections.Counter(r['rule'].split('/')[0] for r in records)
print('=' * 78)
print('catalogue menu.xml: %d menu co MaFile+FolderName; backup: %d' % (len(CAT), len(BACKUP)))
print('in-scope files: %d placed, %d pending; R5n %d' % (len(records), len(pending), r5n_cnt))
for k, v in sorted(rule_cnt.items()):
    print('  rule %-7s %5d' % (k, v))
print('excluded:')
for k, v in sorted(excluded.items()):
    print('  %-40s %5d' % (k, v))
print('pending reasons:')
for k, v in collections.Counter(p['reason'].split(':')[0].split(' (')[0] for p in pending).most_common():
    print('  %-40s %5d' % (k, v))

current = {}
if os.path.isdir(DST):
    for dp, dn, fn in os.walk(DST):
        for f in fn:
            if f in KEEP_IN_DST:
                continue
            current[os.path.relpath(os.path.join(dp, f), DST)] = os.path.join(dp, f)
planned = {r['dst']: r for r in records}
add_ = [d for d in planned if d not in current]
rem_ = [d for d in current if d not in planned]
same = [d for d in planned if d in current]
diffc = [d for d in same if sha(current[d]) != sha(os.path.join(SRC, planned[d]['src']))]
print('diff vs %s: keep %d (content differs %d), add %d, remove %d' % (DST, len(same), len(diffc), len(add_), len(rem_)))
if '--verbose' in sys.argv:
    for d in sorted(add_):
        print('  + %s  <- %s [%s]' % (d, planned[d]['src'], planned[d]['rule']))
    for d in sorted(rem_):
        print('  - %s' % d)
    for p in pending:
        print('  ? %s  %s  %s' % (p['src'], p['reason'], p['candidates']))

# ---------------------------------------------------------------- write
if WRITE:
    for d in sorted(rem_):
        os.remove(current[d])
    for d, r in sorted(planned.items()):
        p = os.path.join(DST, d)
        s = os.path.join(SRC, r['src'])
        if d in current and d not in diffc:
            continue
        os.makedirs(os.path.dirname(p), exist_ok=True)
        shutil.copy2(s, p)
    # drop folders left with no source file (cong_thuc.txt handled by caller, reported here)
    for dp, dn, fn in sorted(os.walk(DST, topdown=False)):
        if dp == DST:
            continue
        real = [f for f in fn if f not in ('.DS_Store',)]
        if not real and not dn:
            for f in fn:
                os.remove(os.path.join(dp, f))
            os.rmdir(dp)
        elif real == ['cong_thuc.txt'] and not dn:
            print('  ORPHAN cong_thuc.txt (no source file left): %s' % os.path.relpath(dp, DST))
    with open(MANIFEST, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['src', 'dst', 'rule', 'version', 'magic', 'root', 'sha256', 'evidence'])
        for r in sorted(records, key=lambda r: r['dst']):
            w.writerow([r['src'], r['dst'], r['rule'], r['ver'], r['magic'], r['root'],
                        sha(os.path.join(SRC, r['src'])), r['evidence']])
    with open(PENDING_CSV, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['src', 'magic', 'root', 'reason', 'candidates', 'ghi_chu'])
        for p in sorted(pending, key=lambda p: p['src']):
            w.writerow([p['src'], p['magic'], p['root'], p['reason'], p['candidates'],
                        PENDING_NOTES.get(p['src'], '')])
    print('WROTE %d planned files (%d copied/updated, %d removed) -> %s; manifest + pending written'
          % (len(planned), len(add_) + len(diffc), len(rem_), DST))
