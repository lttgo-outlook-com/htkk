#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_formulas.py: bộ kiểm độc lập cho _spec/formulas (không dùng mã của biff_formulas.py).

Chạy: /tmp/audit_venv/bin/python -I _spec/tools/verify_formulas.py \
        --formulas _spec/formulas --src /Users/tuanle/codes/HTKK/_extracted \
        --forms /Users/tuanle/codes/HTKK/_forms [--sheetjs /tmp/sj/dump.json] [--determinism]
Đọc lại byte BIFF bằng bộ quét riêng (đếm bản ghi FORMULA/SHRFMLA/ARRAY/NAME/SUPBOOK,
bỏ qua mã giải ptg, chỉ lướt token để lấy chỉ số ixti) rồi đối chiếu với JSON.
"""
import sys, os, re, struct, json, hashlib, zipfile, argparse, collections, subprocess, tempfile, io
try:
    import olefile
except ImportError:
    sys.exit('Thiếu olefile: chạy bằng /tmp/audit_venv/bin/python -I')

# Số kỳ vọng của tập _forms không viết cứng: tính từ _forms/_manifest.csv nối với index.json của T1
# (khóa nối: src = source_name và sha256 = source_sha256), xem expected_from_manifest().
# Lịch sử: _forms cũ 418.494 ô / SHRFMLA 3.138 / defined name 852 trong 429 tệp; _forms mới (tệp vào/ra pending)
# 417.348 ô / SHRFMLA 3.139 / defined name nhóm .xls+.xls1 855 trong 432 tệp, không mất dữ liệu.
EXPECT = {'ext_min': 183, 'catalogue_files': 42, 'catalogue_formulas': 8872}
FAIL = []


def expected_from_manifest(forms_dir, files):
    """Số kỳ vọng của tập _forms: manifest nối index (src + sha256), cộng theo từng dòng manifest có công thức."""
    import csv
    ent = {(e['source_name'], e['source_sha256']): e for e in files if e['status'] == 'ok'}
    ex = collections.Counter()
    with open(os.path.join(forms_dir, '_manifest.csv'), encoding='utf-8') as fh:
        for r in csv.DictReader(fh):
            e = ent.get((r['src'], r['sha256']))
            if e is None:
                continue
            ex['rows'] += 1
            ex['total'] += e['formulas']
            ex['xlsx' if e['format'] == 'xlsx' else 'xls'] += e['formulas']
            if e['format'] == 'biff8':
                ex['shrfmla'] += e['shrfmla_records']
                ex['array'] += e['array_records']
            if r['dst'].endswith(('.xls', '.xls1')):
                ex['names_user'] += e['defined_names_user']
                ex['names_files'] += 1 if e['defined_names_user'] else 0
    return ex


def check(ok, label):
    print(('PASS ' if ok else 'FAIL ') + label)
    if not ok:
        FAIL.append(label)


# ------------------------------------------------------------ bộ quét BIFF độc lập
SZ = {0x01: 4, 0x02: 4, 0x1C: 1, 0x1D: 1, 0x1E: 2, 0x1F: 8, 0x20: 7, 0x21: 2, 0x22: 3, 0x23: 4,
      0x24: 4, 0x25: 8, 0x26: 6, 0x27: 6, 0x28: 6, 0x29: 2, 0x2A: 4, 0x2B: 8, 0x2C: 4, 0x2D: 8,
      0x39: 6, 0x3A: 6, 0x3B: 10, 0x3C: 6, 0x3D: 10}


def ixtis(rg):
    """Các ixti của ptg 3D/NameX trong rgce, hoặc None nếu gặp ptg lạ."""
    out = []
    i = 0
    n = len(rg)
    while i < n:
        t = rg[i]
        i += 1
        if 0x03 <= t <= 0x16:
            continue
        if t == 0x17:
            cch = rg[i]
            hi = rg[i + 1] & 1
            i += 2 + cch * (2 if hi else 1)
            continue
        if t == 0x19:
            g = rg[i]
            w = struct.unpack_from('<H', rg, i + 1)[0]
            i += 3
            if g & 0x04:
                i += 2 * (w + 1)
            continue
        b = t if t < 0x20 else ((t & 0x1F) | 0x20)
        if b not in SZ:
            return None
        if b in (0x39, 0x3A, 0x3B, 0x3C, 0x3D):
            out.append(struct.unpack_from('<H', rg, i)[0])
        i += SZ[b]
    return out


def lost3d(rg, xtif, sup_ext):
    """ptgRef3d/ptgArea3d trỏ sheet đã xóa (itab 0xFFFE/0xFFFF, SUPBOOK nội bộ): [(ptg, row, col)] lấy từ byte."""
    out = []
    i = 0
    n = len(rg)
    while i < n:
        t = rg[i]
        i += 1
        if 0x03 <= t <= 0x16:
            continue
        if t == 0x17:
            cch = rg[i]
            hi = rg[i + 1] & 1
            i += 2 + cch * (2 if hi else 1)
            continue
        if t == 0x19:
            g = rg[i]
            w = struct.unpack_from('<H', rg, i + 1)[0]
            i += 3
            if g & 0x04:
                i += 2 * (w + 1)
            continue
        b = t if t < 0x20 else ((t & 0x1F) | 0x20)
        if b not in SZ:
            return out
        if b in (0x3A, 0x3B):
            ix = struct.unpack_from('<H', rg, i)[0]
            if ix < len(xtif) and xtif[ix][1] in (0xFFFE, 0xFFFF) and not sup_ext[xtif[ix][0]]:
                if b == 0x3A:
                    row, colf = struct.unpack_from('<HH', rg, i + 2)
                    out.append((0x3A, row, colf & 0xFF))
                else:
                    r1, r2, c1, c2 = struct.unpack_from('<HHHH', rg, i + 2)
                    out.append((0x3B, r1, c1 & 0xFF))
        i += SZ[b]
    return out


def raw_records(d):
    pos = 0
    out = []
    while pos + 4 <= len(d):
        t, l = struct.unpack_from('<HH', d, pos)
        b = d[pos + 4:pos + 4 + l]
        if t == 0x3C and out and out[-1][0] == 0x17:
            out[-1] = (0x17, out[-1][1] + b)
        else:
            out.append((t, b))
        pos += 4 + l
    return out


def scan_biff(data):
    ole = olefile.OleFileIO(io.BytesIO(data))
    s = 'Workbook' if ole.exists('Workbook') else ('Book' if ole.exists('Book') else None)
    if s is None:
        return None
    d = ole.openstream(s).read()
    ole.close()
    sheets = []
    sup_ext = []
    xti = []
    xti_full = []
    shr_lost = {}
    names_user = names_builtin = 0
    cells = []     # (sheetidx, r, c, ixti_list or None, is_exp)
    n_shr = n_arr = 0
    depth = 0
    sub = -1
    prev_cell = {}
    shr = {}
    pending = []
    for t, b in raw_records(d):
        if t == 0x809:
            depth += 1
            if depth == 1:
                sub += 1
            continue
        if t == 0x0A:
            depth -= 1
            continue
        if sub == 0:
            if t == 0x85:
                cch = b[6]
                fl = b[7]
                sheets.append(b[8:8 + cch * 2].decode('utf-16le') if fl & 1 else b[8:8 + cch].decode('latin-1'))
            elif t == 0x1AE:
                sup_ext.append(not (len(b) == 4 and b[2:4] in (b'\x01\x04', b'\x01\x3a')))
            elif t == 0x17:
                cx = struct.unpack_from('<H', b, 0)[0]
                xti = [struct.unpack_from('<H', b, 2 + 6 * k)[0] for k in range(cx)]
                xti_full = [struct.unpack_from('<HHH', b, 2 + 6 * k) for k in range(cx)]
            elif t == 0x18:
                if struct.unpack_from('<H', b, 0)[0] & 0x20:
                    names_builtin += 1
                else:
                    names_user += 1
        elif sub >= 1 and depth == 1:
            si = sub - 1
            if t == 0x06:
                r, c = struct.unpack_from('<HH', b, 0)
                cce = struct.unpack_from('<H', b, 20)[0]
                rg = b[22:22 + cce]
                exp = len(rg) == 5 and rg[0] == 0x01
                tgt = struct.unpack_from('<HH', rg, 1) if exp else (r, c)
                cells.append([si, r, c, None if exp else ixtis(rg), exp, tuple(tgt), None if exp else lost3d(rg, xti_full, sup_ext)])
                prev_cell[si] = len(cells) - 1
            elif t == 0x4BC:
                n_shr += 1
                cce = struct.unpack_from('<H', b, 8)[0]
                idx = prev_cell.get(si)
                if idx is not None:
                    shr[(si, cells[idx][1], cells[idx][2])] = ixtis(b[10:10 + cce])
                    shr_lost[(si, cells[idx][1], cells[idx][2])] = lost3d(b[10:10 + cce], xti_full, sup_ext)
            elif t == 0x221:
                n_arr += 1
    # gán ixti cho ô shared qua ptgExp
    for cell in cells:
        if cell[4]:
            pass
    return {'sheets': sheets, 'sup_ext': sup_ext, 'xti': xti, 'names_user': names_user,
            'names_builtin': names_builtin, 'cells': cells, 'n_shr': n_shr, 'n_arr': n_arr,
            'shr': shr, 'shr_lost': shr_lost, 'raw': d}


def cell_ixtis(sc, cell):
    """ixti của ô (kể cả ô shared) từ bộ quét độc lập."""
    si, r, c, ix, exp, _ = cell
    if not exp:
        return ix
    return None


def a1(r, c):
    s = ''
    c += 1
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return '%s%d' % (s, r + 1)


def scan_xlsx(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    wb = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rm = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        t = m.group(0)
        rm[re.search(r'Id="([^"]*)"', t).group(1)] = re.search(r'Target="([^"]*)"', t).group(1)
    import html
    sheets = []
    for m in re.finditer(r'<sheet\b[^>]*>', wb):
        t = m.group(0)
        nm = html.unescape(re.search(r'name="([^"]*)"', t).group(1))
        rid = re.search(r'r:id="([^"]*)"', t).group(1)
        tg = rm[rid].lstrip('/')
        sheets.append((nm, tg if tg.startswith('xl/') else 'xl/' + tg))
    cells = set()
    shr = arr = 0
    for nm, part in sheets:
        x = z.read(part).decode('utf8')
        for m in re.finditer(r'<c\b([^>]*?)(?<!/)>(.*?)</c>', x, re.S):
            if '<f' in m.group(2):
                cells.add((nm, re.search(r'\br="([A-Z]+\d+)"', m.group(1)).group(1)))
        shr += len(re.findall(r'<f\b[^>]*t="shared"[^>]*>[^<]', x))
        arr += len(re.findall(r'<f\b[^>]*t="array"', x))
    names_user = len([1 for m in re.finditer(r'<definedName\b[^>]*name="([^"]*)"', wb) if not m.group(1).startswith('_xlnm.')])
    return {'cells': cells, 'n_shr': shr, 'n_arr': arr, 'names_user': names_user}


# ------------------------------------------------------------ chuẩn hóa so với SheetJS
def make_norm(enabled):
    def nz(x):
        p = re.split(r'("(?:[^"]|"")*")', x)
        for i in range(0, len(p), 2):
            s = p[i]
            if 'ref_sheet' in enabled:
                s = re.sub(r"(?:'[^']*'|[\w.]+)!#REF!", '#REF!', s)
            if 'xlfn' in enabled:
                s = s.replace('_xlfn.', '')
            if 'ref_junk' in enabled:
                s = re.sub(r'#REF!\$?[A-Z]{1,3}\$?\d+', '#REF!', s)
            if 'quote' in enabled:
                s = re.sub(r"'([^']*)'!", r'\1!', s)
            if 'ws' in enabled:
                s = re.sub(r'\s+', '', s)
            p[i] = s
        return ''.join(p)
    return nz


ALL_NORMS = ['ref_sheet', 'xlfn', 'ref_junk', 'quote', 'ws']
NORM_DESC = {
    'ws': 'khoảng trắng/xuống dòng quanh toán tử (attrSpace) mà SheetJS bỏ hoặc đổi chỗ',
    'quote': "SheetJS bỏ nháy đơn quanh tên sheet ('01'!A1 thành 01!A1)",
    'ref_sheet': 'tham chiếu 3D lỗi: bản này giữ tên sheet (Sheet!#REF!), SheetJS chỉ in #REF!',
    'xlfn': 'hàm tương lai: bản này giữ _xlfn.IFNA (như lưu trong tệp), SheetJS bỏ tiền tố',
    'ref_junk_lost3d': 'ptgRef3d/ptgArea3d trỏ sheet đã xóa (itab 0xFFFE/0xFFFF): SheetJS in #REF!V30, bản này in #REF! và giữ địa chỉ trong lost_refs',
    'ref_junk_referr': 'ptgRefErr (0x2A): SheetJS in #REF! kèm địa chỉ rác, bản này chỉ #REF!',
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--formulas', required=True)
    ap.add_argument('--src', required=True)
    ap.add_argument('--forms', required=True)
    ap.add_argument('--sheetjs', default='/tmp/sj/dump.json')
    ap.add_argument('--determinism', action='store_true')
    a = ap.parse_args()

    idx = json.load(open(os.path.join(a.formulas, 'index.json'), encoding='utf-8'))
    files = idx['files']
    ok_files = [e for e in files if e['status'] == 'ok']
    print('index: %d tệp, ok=%d, bỏ qua=%d' % (len(files), len(ok_files), len(files) - len(ok_files)))
    for e in files:
        if e['status'] != 'ok':
            print('  bỏ qua: %s (%s)' % (e['source_name'], e['reason']))
    EX = expected_from_manifest(a.forms, files)
    print('kỳ vọng từ _manifest.csv + index.json: %d dòng manifest có công thức, %d ô (.xls/.xls1 %d, .xlsx %d), SHRFMLA %d, ARRAY %d, defined name .xls/.xls1 %d trong %d tệp' % (
        EX['rows'], EX['total'], EX['xls'], EX['xlsx'], EX['shrfmla'], EX['array'], EX['names_user'], EX['names_files']))
    by_sha = collections.defaultdict(list)
    for e in ok_files:
        by_sha[e['source_sha256']].append(e)

    # ---- tập tệp thuộc _forms (theo sha256)
    forms_files = []
    for dp, dn, fn in os.walk(a.forms):
        for f in sorted(fn):
            fp = os.path.join(dp, f)
            with open(fp, 'rb') as fh:
                mg = fh.read(4)
            if mg == b'\xd0\xcf\x11\xe0' or mg[:2] == b'PK':   # nhận dạng theo magic bytes, không theo đuôi
                forms_files.append(fp)
    forms_files.sort()
    forms_hash = {}
    for p in forms_files:
        forms_hash[p] = hashlib.sha256(open(p, 'rb').read()).hexdigest()
    miss = [p for p, h in forms_hash.items() if h not in by_sha]
    print('_forms: %d tệp bảng tính, %d sha duy nhất, không khớp nguồn: %d' % (
        len(forms_files), len(set(forms_hash.values())), len(miss)))

    # ---- đọc JSON, đối chiếu quét độc lập
    data_cache = {}
    tot_json = {}
    mism = []
    n_exp_cells = 0
    all_stats = collections.Counter()
    lost_cells_n = [0, set()]
    for e in ok_files:
        j = json.load(open(os.path.join(a.formulas, e['json']), encoding='utf-8'))
        data_cache[e['key']] = j
        raw = open(e['source_path'] if os.path.exists(e['source_path']) else os.path.join(a.src, e['source_name']), 'rb').read()
        assert hashlib.sha256(raw).hexdigest() == e['source_sha256'], 'sha lệch ' + e['source_name']
        assert j['header']['source_sha256'] == e['source_sha256']
        jset = collections.Counter((f['sheet'], f['cell']) for f in j['formulas'])
        if e['format'] == 'biff8':
            sc = scan_biff(raw)
            names = sc['sheets']
            iset = collections.Counter((names[c[0]], a1(c[1], c[2])) for c in sc['cells'])
            n_shr, n_arr = sc['n_shr'], sc['n_arr']
            nuser = sc['names_user']
            # kind khớp ptgExp: số ô ptgExp = shared + array
            exp_cells = sum(1 for c in sc['cells'] if c[4])
            kinds = sum(1 for f in j['formulas'] if f['kind'] != 'normal')
            if exp_cells != kinds:
                mism.append((e['source_name'], 'ptgExp %d != shared+array %d' % (exp_cells, kinds)))
            # SUPBOOK ngoài
            nsup = sum(1 for x in sc['sup_ext'] if x)
            if nsup != e['supbooks_external']:
                mism.append((e['source_name'], 'supbook ngoài %d != %d' % (nsup, e['supbooks_external'])))
            # ô có tham chiếu ngoài (ixti → supbook ngoài), chỉ ô không shared
            ext_i = set()
            for c in sc['cells']:
                if c[4]:
                    ix = sc['shr'].get((c[0], c[5][0], c[5][1]))
                else:
                    ix = c[3]
                if ix is None:
                    continue
                if any(i < len(sc['xti']) and sc['xti'][i] < len(sc['sup_ext']) and sc['sup_ext'][sc['xti'][i]] for i in ix):
                    ext_i.add((names[c[0]], a1(c[1], c[2])))
            ext_j = set((f['sheet'], f['cell']) for f in j['formulas'] if f['ext_refs'])
            # ô shared dùng chung SHRFMLA: chỉ so cho ô chuẩn (anchor); ô shared khác kiểm bằng tập con
            anchors = set()
            if ext_i - ext_j or ext_j - ext_i:
                # ô shared khác anchor không có ixti trong bản quét độc lập; kiểm tập con
                if not (ext_i <= ext_j):
                    mism.append((e['source_name'], 'ext độc lập có ô JSON không có: %s' % sorted(ext_i - ext_j)[:3]))
                extra = ext_j - ext_i
                bad = [x for x in extra if not any(f['kind'] == 'shared' and (f['sheet'], f['cell']) == x for f in j['formulas'])]
                if bad:
                    mism.append((e['source_name'], 'ext JSON thừa ô không shared: %s' % bad[:3]))
        else:
            sx = scan_xlsx(raw)
            iset = collections.Counter(sx['cells'])
            n_shr, n_arr = sx['n_shr'], sx['n_arr']
            nuser = sx['names_user']
        if iset != jset:
            mism.append((e['source_name'], 'tập ô lệch: chỉ độc lập %d, chỉ JSON %d' % (
                sum((iset - jset).values()), sum((jset - iset).values()))))
        if n_shr != e['shrfmla_records'] or n_arr != e['array_records']:
            mism.append((e['source_name'], 'SHRFMLA/ARRAY %d/%d != %d/%d' % (n_shr, n_arr, e['shrfmla_records'], e['array_records'])))
        if nuser != e['defined_names_user']:
            mism.append((e['source_name'], 'defined name user %d != %d' % (nuser, e['defined_names_user'])))
        if e['format'] == 'biff8':
            lost_ind = {}
            for c in sc['cells']:
                key = (names[c[0]], a1(c[1], c[2]))
                lst = c[6] if not c[4] else sc['shr_lost'].get((c[0], c[5][0], c[5][1]), [])
                if lst:
                    lost_ind[key] = lst
            lost_j = {(f['sheet'], f['cell']): f['lost_refs'] for f in j['formulas'] if 'lost_refs' in f}
            if set(lost_ind) != set(lost_j):
                mism.append((e['source_name'], 'lost_refs: tập ô lệch (độc lập %d, JSON %d)' % (len(lost_ind), len(lost_j))))
            for key in set(lost_ind) & set(lost_j):
                if len(lost_ind[key]) != len(lost_j[key]):
                    mism.append((e['source_name'], 'lost_refs %s: số ptg lệch' % (key,)))
                elif not any(c[4] for c in sc['cells'] if (names[c[0]], a1(c[1], c[2])) == key):
                    if [(x[0], x[1], x[2]) for x in lost_ind[key]] != [(int(x['ptg'], 16), x['row'], x['col']) for x in lost_j[key]]:
                        mism.append((e['source_name'], 'lost_refs %s: địa chỉ lệch so với byte' % (key,)))
            if lost_j:
                lost_cells_n[0] += len(lost_j)
                lost_cells_n[1].add(e['source_name'])
        e['_ind'] = {'n_shr': n_shr, 'n_arr': n_arr, 'names_user': nuser, 'cells': sum(iset.values())}
        for f in j['formulas']:
            if f['formula'] is None:
                all_stats['formula_none'] += 1
            if '\r' in f['formula'] or '\n' in f['formula']:
                all_stats['formula_has_crlf'] += 1
        for er in j['errors']:
            all_stats['errors'] += 1

    # ---- (1)
    print('\n== (1) Tổng ô công thức trên tập tệp _forms')
    xls_t = xlsx_t = 0
    per_file = {}
    for p, h in forms_hash.items():
        e = by_sha.get(h, [None])[0]
        if e is None:
            continue
        n = e['formulas']
        per_file[p] = n
        if e['format'] == 'xlsx':
            xlsx_t += n
        else:
            xls_t += n
    tot = xls_t + xlsx_t
    print('  .xls/.xls1: %d | .xlsx: %d | tổng: %d (theo từng tệp _forms, %d tệp)' % (xls_t, xlsx_t, tot, len(per_file)))
    uniq = {}
    for p, h in forms_hash.items():
        if h in by_sha:
            uniq[h] = by_sha[h][0]['formulas']
    print('  theo sha duy nhất: %d' % sum(uniq.values()))
    check(tot == EX['total'], '(1) tổng = %d (kỳ vọng %d); .xls %d/%d, .xlsx %d/%d' % (
        tot, EX['total'], xls_t, EX['xls'], xlsx_t, EX['xlsx']))
    if tot != EX['total']:
        print('  lệch: %d' % (tot - EX['total']))
    print('  tổng toàn bộ nguồn (cả tệp ngoài _forms): %d trong %d tệp' % (
        sum(e['formulas'] for e in ok_files), len(ok_files)))

    # ---- (2)
    print('\n== (2) Tập (tệp, sheet, ô) so với đếm bản ghi độc lập')
    n_ind = sum(e['_ind']['cells'] for e in ok_files)
    n_js = sum(e['formulas'] for e in ok_files)
    print('  ô độc lập (FORMULA): %d | ô JSON: %d | tệp lệch: %d' % (n_ind, n_js, len(mism)))
    for m in mism[:20]:
        print('  lệch:', m)
    f_shr = sum(by_sha[h][0]['_ind']['n_shr'] for h in forms_hash.values() if h in by_sha and by_sha[h][0]['format'] == 'biff8')
    f_arr = sum(by_sha[h][0]['_ind']['n_arr'] for h in forms_hash.values() if h in by_sha and by_sha[h][0]['format'] == 'biff8')
    x_shr = sum(by_sha[h][0]['_ind']['n_shr'] for h in forms_hash.values() if h in by_sha and by_sha[h][0]['format'] == 'xlsx')
    print('  SHRFMLA/ARRAY độc lập trên tệp BIFF của _forms: %d / %d (kỳ vọng %d / %d); ngoài ra %d shared master trong .xlsx (không tính vào SHRFMLA)' % (f_shr, f_arr, EX['shrfmla'], EX['array'], x_shr))
    print('  ô lỗi giải mã: %d; ô formula=None: %d' % (all_stats['errors'], all_stats['formula_none']))
    check(not mism and n_ind == n_js and all_stats['errors'] == 0, '(2) tập ô khớp 100%, SHRFMLA/ARRAY/ptgExp/SUPBOOK/defined name khớp, 0 lỗi giải mã')
    check((f_shr, f_arr) == (EX['shrfmla'], EX['array']), '(2b) SHRFMLA=%d ARRAY=%d trên _forms' % (f_shr, f_arr))

    # ---- (3)
    print('\n== (3) So với SheetJS (xlsx@0.18.5), mọi ô không có tham chiếu ngoài')
    if not os.path.exists(a.sheetjs):
        os.makedirs(os.path.dirname(a.sheetjs), exist_ok=True)
        subprocess.run(['node', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sheetjs_dump.cjs'), a.src, a.sheetjs], check=True)
    sj = json.load(open(a.sheetjs, encoding='utf-8'))
    full = make_norm(set(ALL_NORMS))
    cnt = collections.Counter()
    need = collections.Counter()
    unexplained = []
    n_cmp = n_eq = 0
    for e in ok_files:
        j = data_cache[e['key']]
        s = sj.get(e['source_name'], {})
        if '__error' in s:
            cnt['sheetjs_error_files'] += 1
            continue
        theirs = {(sh, c): v for sh, m in s.items() for c, v in m.items()}
        mine = {(f['sheet'], f['cell']): f for f in j['formulas']}
        cnt['cells_only_json'] += len(set(mine) - set(theirs))
        cnt['cells_only_sheetjs'] += len(set(theirs) - set(mine))
        for k, f in mine.items():
            if f['ext_refs']:
                cnt['skipped_ext'] += 1
                continue
            b = theirs.get(k)
            if b is None:
                continue
            n_cmp += 1
            x = f['formula']
            if x == b:
                n_eq += 1
                continue
            if full(x) != full(b):
                unexplained.append((e['source_name'], k, x, b))
                continue
            used = tuple(nm for nm in ALL_NORMS if make_norm(set(ALL_NORMS) - {nm})(x) != make_norm(set(ALL_NORMS) - {nm})(b))
            if 'ref_junk' in used:
                lr = f.get('lost_refs')
                used = tuple(('ref_junk_lost3d' if lr else 'ref_junk_referr') if u == 'ref_junk' else u for u in used)
                if lr:
                    got = [m.replace('$', '') for m in re.findall(r'#REF!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)', b)]
                    want = [r_['ref_a1'].replace('$', '') for r_ in lr]
                    cnt['lost3d_addr_match' if got == want else 'lost3d_addr_nomatch'] += 1
            need[used] += 1
    print('  so sánh %d ô: giống y hệt %d; khác nhưng khớp sau chuẩn hóa %d; không giải thích được %d' % (
        n_cmp, n_eq, sum(need.values()), len(unexplained)))
    print('  ô chỉ có ở JSON: %d; chỉ có ở SheetJS: %d; bỏ qua (có tham chiếu ngoài): %d' % (
        cnt['cells_only_json'], cnt['cells_only_sheetjs'], cnt['skipped_ext']))
    print('  các loại khác biệt (ô cần chuẩn hóa nào, đếm theo tổ hợp tối thiểu):')
    for used, v in need.most_common():
        print('    %6d  %s' % (v, ' + '.join(used) if used else '(chuẩn hóa dư)'))
    for nm in ['ws', 'quote', 'ref_sheet', 'xlfn', 'ref_junk_lost3d', 'ref_junk_referr']:
        print('    - %s: %s' % (nm, NORM_DESC[nm]))
    print('  lost_refs: địa chỉ trong lost_refs khớp địa chỉ SheetJS in sau #REF!: %d ô khớp, %d ô lệch' % (cnt['lost3d_addr_match'], cnt['lost3d_addr_nomatch']))
    check(cnt['lost3d_addr_nomatch'] == 0, '(3b) lost_refs trùng địa chỉ SheetJS ở mọi ô sheet đã xóa')
    print('  lost_refs (đếm lại từ byte, tập ô + số ptg + địa chỉ): %d ô trong %d tệp' % (lost_cells_n[0], len(lost_cells_n[1])))
    for u in unexplained[:10]:
        print('  KHÔNG GIẢI THÍCH:', u)
    check(not unexplained and cnt['cells_only_json'] == 0 and cnt['cells_only_sheetjs'] == 0,
          '(3) 0 ô khác ngoài các loại đã liệt kê')

    # ---- (4)
    print('\n== (4) Tham chiếu ngoài')
    ext_forms = 0
    ext_all = 0
    ext_files = []
    for e in ok_files:
        ext_all += e['formulas_ext']
    for h in forms_hash.values():
        if h in by_sha:
            ext_forms += by_sha[h][0]['formulas_ext']
            if by_sha[h][0]['formulas_ext']:
                ext_files.append(by_sha[h][0]['source_name'])
    sup_files_forms = len([1 for h in forms_hash.values() if h in by_sha and by_sha[h][0]['supbooks_external']])
    print('  ô có tham chiếu ngoài: tập _forms %d (%d tệp), toàn nguồn %d; tệp có SUPBOOK ngoài (tập _forms): %d' % (
        ext_forms, len(ext_files), ext_all, sup_files_forms))
    # kiểm chỉ số SUPBOOK: mọi ext_ref phải trỏ SUPBOOK kind=external và token mang [ext:n] khớp
    bad = 0
    for e in ok_files:
        j = data_cache[e['key']]
        for f in j['formulas']:
            for r in f['ext_refs']:
                sb = j['supbooks'][r['supbook']] if r['supbook'] < len(j['supbooks']) else None
                if sb is None or sb['kind'] != 'external' or ('[ext:%d]' % r['supbook']) not in r['token']:
                    bad += 1
    print('  ext_ref sai chỉ số SUPBOOK/không external: %d' % bad)
    # kiểm tay 5 ô
    picks = [('_2_lncn_psc_tt80_dk_284.xls1', '01-2', 'L11')]
    seen = set(picks)
    for e in ok_files:
        if e['formulas_ext'] and len(picks) < 5 and e['source_name'] not in [p[0] for p in picks]:
            j = data_cache[e['key']]
            for f in j['formulas']:
                if f['ext_refs']:
                    picks.append((e['source_name'], f['sheet'], f['cell']))
                    break
    hand_ok = True
    for name, sh, cell in picks[:5]:
        e = next(x for x in ok_files if x['source_name'] == name)
        j = data_cache[e['key']]
        f = next(x for x in j['formulas'] if x['sheet'] == sh and x['cell'] == cell)
        refs = [(r['supbook'], j['supbooks'][r['supbook']]['path'], j['supbooks'][r['supbook']]['sheets']) for r in f['ext_refs']]
        print('  tay: %s | %s!%s = %s' % (name, sh, cell, f['formula']))
        for sbi, path, shs in refs:
            print('       SUPBOOK[%d] path=%s sheets=%s' % (sbi, path, shs))
    e = next(x for x in ok_files if x['source_name'] == '_2_lncn_psc_tt80_dk_284.xls1')
    j = data_cache[e['key']]
    f = next(x for x in j['formulas'] if x['sheet'] == '01-2' and x['cell'] == 'L11')
    sb = j['supbooks'][f['ext_refs'][0]['supbook']]
    hand_ok = ('02_LNCN_VSP_TT80_DK_284' in sb['path'] and "[ext:%d]'01'!" % sb['index'] in f['formula'])
    print('  mẫu bắt buộc: formula=%s ; đích sheet "01" của %s -> %s' % (f['formula'], sb['path'], 'ĐÚNG' if hand_ok else 'SAI'))
    check(ext_forms >= EXPECT['ext_min'] and bad == 0 and hand_ok, '(4) ô tham chiếu ngoài trên _forms = %d (>= %d), chỉ số SUPBOOK đúng, mẫu 01-2!L11 đúng' % (ext_forms, EXPECT['ext_min']))

    # ---- (5)
    print('\n== (5) Defined name do người dùng đặt')
    grp = collections.Counter()
    for p, h in forms_hash.items():
        if h in by_sha:
            e = by_sha[h][0]
            g = 'xls/xls1' if p.endswith(('.xls', '.xls1')) else ('xlsx' if e['format'] == 'xlsx' else 'ole đuôi khác')
            grp[(g, 'tên')] += e['defined_names_user']
            grp[(g, 'tệp có tên')] += 1 if e['defined_names_user'] else 0
    nu_all = sum(e['defined_names_user'] for e in ok_files)
    nb_all = sum(e['defined_names_builtin'] for e in ok_files)
    nf_all = sum(1 for e in ok_files if e['defined_names_user'])
    for g in ('xls/xls1', 'ole đuôi khác', 'xlsx'):
        print('  _forms, nhóm %-14s: %d tên người dùng trong %d tệp' % (g, grp[(g, 'tên')], grp[(g, 'tệp có tên')]))
    print('  toàn nguồn: %d tên người dùng trong %d tệp (+%d tên dựng sẵn)' % (nu_all, nf_all, nb_all))
    print('  giải thích: nhóm .xls/.xls1 khớp số kỳ vọng từ manifest; tổng toàn nguồn gồm thêm tệp OLE đuôi .x/.xl/.xl1/không đuôi (84) và .xlsx (32)')
    check((grp[('xls/xls1', 'tên')], grp[('xls/xls1', 'tệp có tên')]) == (EX['names_user'], EX['names_files']),
          '(5) defined name người dùng trong .xls/.xls1 của _forms = %d tên / %d tệp (kỳ vọng %d / %d)' % (
              grp[('xls/xls1', 'tên')], grp[('xls/xls1', 'tệp có tên')], EX['names_user'], EX['names_files']))

    # ---- (6)
    print('\n== (6) catalogue*.xls (catalogue_*.xls và catalogue.xls)')
    cat = [e for e in ok_files if re.match(r'catalogue(_|\.)', e['source_name'])]
    ct = sum(e['formulas'] for e in cat)
    print('  %d tệp catalogue, tổng công thức %d' % (len(cat), ct))
    check(len(cat) == EXPECT['catalogue_files'] and ct == EXPECT['catalogue_formulas'], '(6) 42 catalogue, tổng %d (kỳ vọng %d)' % (ct, EXPECT['catalogue_formulas']))
    for nm in ('header.xls', 'in_tmp.xls'):
        print('  %s: %s' % (nm, 'có' if any(e['source_name'] == nm for e in ok_files) else 'KHÔNG CÓ'))

    # ---- chung
    print('\n== Khác')
    print('  ô có CR/LF trong công thức (JSON đã escape): %d; formula None: %d' % (all_stats['formula_has_crlf'], all_stats['formula_none']))
    sz = 0
    for dp, dn, fn in os.walk(a.formulas):
        for f in fn:
            sz += os.path.getsize(os.path.join(dp, f))
    print('  số tệp JSON: %d (+ index.json); kích thước %s: %.1f MB' % (len(ok_files), a.formulas, sz / 1e6))
    if a.determinism:
        print('\n== Chạy trích hai lần')
        tool = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'biff_formulas.py')
        hs = []
        with tempfile.TemporaryDirectory() as t1, tempfile.TemporaryDirectory() as t2:
            for t in (t1, t2):
                subprocess.run([sys.executable, '-I', tool, '--src', a.src, '--menu', os.path.join(a.src, 'menu.xml'), '--out', t],
                               check=True, stdout=subprocess.DEVNULL)
                h = hashlib.sha256()
                for f in sorted(os.listdir(t)):
                    h.update(f.encode())
                    h.update(open(os.path.join(t, f), 'rb').read())
                hs.append(h.hexdigest())
            same_as_committed = all(open(os.path.join(t1, f), 'rb').read() == open(os.path.join(a.formulas, f), 'rb').read()
                                    for f in os.listdir(t1))
        print('  hash lần 1 = %s\n  hash lần 2 = %s' % (hs[0][:16], hs[1][:16]))
        check(hs[0] == hs[1] and same_as_committed, 'chạy hai lần giống từng byte và khớp bản đã commit')
    print('\nKẾT QUẢ: %s' % ('ĐẠT' if not FAIL else 'KHÔNG ĐẠT (%d mục)' % len(FAIL)))
    sys.exit(1 if FAIL else 0)


if __name__ == '__main__':
    main()
