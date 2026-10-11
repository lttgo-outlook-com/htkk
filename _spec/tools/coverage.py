#!/usr/bin/env python3
"""Ma trận phủ theo ô cho các tờ khai HTKK 5.7.9.

Chạy từ gốc worktree:
  nice -n 15 python3 -I _spec/tools/coverage.py [--out _spec/coverage] [--sample 10 --seed 20260511]

Khóa form: (MaFile, VersionXML) theo menu.xml (242 form).
Nguồn đầu vào (chỉ đọc):
  _forms/_manifest.csv + _forms/*/ *.xml (mẫu ô, Sections), _spec/formulas (T1), _spec/config/mapcongthucchosua (T4),
  _spec/xsd (T5), _spec/dv (T9), T2 rules, T3 validations (đường dẫn tuyệt đối của worktree khác, tham số --t2 --t3).
Quy ước:
  CellID = ô trong .xls ("xls"); CellID2 = ô trong .xls1 ("xls1").
  Ô dữ liệu = ô có Path trong mẫu XML mà đường dẫn đầy đủ có trong XSD (T5). Ô có Path không giải được vào XSD
  được ghi riêng (path_ngoai_xsd), không tính vào tổng ô dữ liệu.
  Ô của bảng động (Dynamic=1) = một mục cho mỗi cột (khóa "COT:<cột>"), hàng xét tại hàng bắt đầu dữ liệu.
"""
import argparse, csv, glob, json, os, random, re, sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

ROOT = os.getcwd()
MENU = '/Users/tuanle/codes/HTKK/_extracted/menu.xml'
T2 = '/Users/tuanle/codes/HTKK/.claude/worktrees/agent-a69710347074048d8/_spec/rules/'
T3 = '/Users/tuanle/codes/HTKK/.claude/worktrees/agent-ae09dfc0b7c50411a/_spec/validations/'
# StatusID chỉ được parse vào HTKKCMCCellInfo, không có mã C# nào đọc lại; hằng HTKKConstant.cs:49-51:
# STATUS_ID_LOAD_DT_YES="0001" (ô nạp dữ liệu từ tệp), STATUS_ID_LOAD_DT_NO="0002". Dùng 0001 làm ô nhập (heuristic).
INPUT_STATUS = {'0001'}


def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


A1 = re.compile(r'^([A-Z]+)(\d+)$')
RANGE = re.compile(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$')


def jload(p):
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def read_xml(p):
    b = open(p, 'rb').read()
    if b[:2] in (b'\xff\xfe', b'\xfe\xff'):
        t = b.decode('utf-16')
    else:
        t = b.decode('utf-8-sig')
    t = re.sub(r'^<\?xml[^>]*\?>', '', t.lstrip())
    return ET.fromstring(t)


def norm_formula(f):
    if f is None:
        return None
    s = str(f).strip()
    if s.startswith('='):
        s = s[1:]
    return re.sub(r'[\s$]', '', s).upper()


def nodots(v):
    return v.replace('.', '')


# ---------------------------------------------------------------- tải dữ liệu
def load_menu():
    root = ET.parse(MENU).getroot()
    forms = {}
    for m in root.iter('Menu'):
        if not m.get('MaFile'):
            continue
        key = (m.get('MaFile'), m.get('VersionXML'))
        val = m.find('Validity')
        sheets = []
        if val is not None:
            for s in val.findall('Sheet'):
                sheets.append({'ID': s.get('ID'), 'SheetName': s.get('SheetName'),
                               'DataFile': s.get('DataFile'), 'XMLnode': s.get('XMLnode')})
        forms[key] = {'kieu_khbs': m.get('KieuKHBS'), 'menu_id': m.get('ID'), 'folder': val.get('FolderName') if val is not None else None,
                      'sheets': sheets}
    return forms


def load_manifest(forms_dir):
    rows = list(csv.DictReader(open(os.path.join(forms_dir, '_manifest.csv'), encoding='utf-8')))
    return rows


def manifest_by_form(rows):
    by = defaultdict(list)
    for r in rows:
        m = re.match(r'menu (\S+) MaFile=(\S+) FolderName=(\S+)', r['evidence'])
        if m:
            by[(m.group(2), r['version'])].append(r)
    return by


def load_formulas(path_dir):
    idx = jload(os.path.join(path_dir, 'index.json'))
    by = defaultdict(dict)  # (mafile, ver) -> {'xls': (src, json), 'xls1': ...}
    for f in idx['files']:
        if f.get('mafile') and f.get('status') == 'ok':
            n = f['source_name'].lower()
            ext = n.rsplit('.', 1)[-1]
            # đuôi tên nguồn bị cắt cụt: xls/xl = mẫu .xls; xls1/xl1/x1 = mẫu .xls1; xlsx (bảng kê) bỏ
            if ext in ('xls1', 'xl1', 'x1'):
                kind = 'xls1'
            elif ext in ('xls', 'xl', 'x'):
                kind = 'xls'
            else:
                by[(f['mafile'], f['version_xml'])].setdefault('_bo_qua', []).append(f['source_name'])
                continue
            d = by[(f['mafile'], f['version_xml'])]
            d.setdefault('_ds_' + kind, []).append(f)
    for (mf, vx), d in by.items():
        for kind in ('xls', 'xls1'):
            ds = d.pop('_ds_' + kind, None)
            if not ds:
                continue
            # ưu tiên tệp không phải *view* (trừ form VIEW) và tên không kết thúc bằng '_' trước đuôi
            ds.sort(key=lambda f: (('view' in f['source_name'].lower()) and ('VIEW' not in mf),
                                   re.search(r'_\.[a-z0-9]+$', f['source_name']) is not None, f['source_name']))
            d[kind] = ds[0]
            if len(ds) > 1:
                d.setdefault('_trung_khoa', []).extend(x['source_name'] for x in ds[1:])
    return by


def load_rowreplace(p):
    d = jload(p)
    maps = {}
    for mp in d['root']['children']:
        items = []
        for it in mp.get('children', []):
            a = it.get('attrs', {})
            items.append(a)
        maps.setdefault(mp['attrs']['ID'], []).extend(items)
    return maps


def load_t2(d):
    by = defaultdict(list)
    unattached = []
    for f in sorted(glob.glob(d + '*.json')):
        j = jload(f)
        if not isinstance(j, dict) or not isinstance(j.get('records'), list):
            continue
        k = j.get('key')
        recs = j['records']
        if isinstance(k, dict) and k.get('MaFile') and k.get('in_menu'):
            by[(k['MaFile'], k['VersionXML'])].extend(recs)
        else:
            unattached.append((os.path.basename(f), recs))
    return by, unattached


def load_t3(d):
    by = defaultdict(list)
    unattached = []
    for f in sorted(glob.glob(d + '*.json')):
        j = jload(f)
        if not isinstance(j, dict) or 'methods' not in j:
            continue
        k = j['key']
        checks = []
        for m in j['methods']:
            for c in m['checks']:
                checks.append(c)
        if k.get('MaFile') and k.get('in_menu'):
            by[(k['MaFile'], k['VersionXML'])].extend(checks)
        else:
            unattached.append((os.path.basename(f), checks))
    return by, unattached


# ---------------------------------------------------------------- mẫu XML ô
class Cell:
    __slots__ = ('sheet', 'section', 'dynamic', 'cid', 'cid2', 'path', 'full', 'ctl', 'status',
                 'tpath', 'row0', 'row0_2', 'minv', 'maxv', 'in_xsd', 'xsd', 'calc', 'chk', 'locks', 'region')

    def __init__(self):
        self.calc = defaultdict(list)
        self.chk = defaultdict(list)
        self.locks = Counter()


def parse_cellid(c):
    if not c:
        return None
    m = re.match(r'^([A-Z]+)_(\d+)$', c)
    if m:
        return m.group(1) + m.group(2)
    return None


def build_cells(sheets, sec_files, ver):
    """sec_files: {sheet_idx: path}. Trả (cells, nhật ký)."""
    cells = []
    log = []
    for si, sh in enumerate(sheets):
        p = sec_files.get(si)
        if not p:
            log.append({'sheet': si, 'ID': sh['ID'], 'loi': 'khong_tim_thay_tep_Sections'})
            continue
        try:
            root = read_xml(p)
        except Exception as e:  # noqa
            log.append({'sheet': si, 'ID': sh['ID'], 'loi': 'parse_xml: %s' % e})
            continue
        node = sh.get('XMLnode') or ''
        parent = node.rsplit('/', 1)[0] if '/' in node else ''
        for sec_i, sec in enumerate(root.findall('Section')):
            dyn = sec.get('Dynamic') == '1'
            tpath = sec.get('TablePath') or ''
            loc = sec.find('Location')
            row0 = row02 = None
            if loc is not None:
                def gi(tag):
                    e = loc.find(tag)
                    try:
                        return int(e.text)
                    except Exception:
                        return None
                row0 = gi('DataLocationTemplateRowStart')
                row02 = gi('DataLocationReportRowStart')
            holders = []
            if dyn:
                for tg in ('RowInfo', 'Header', 'Footer'):
                    h = sec.find(tg)
                    if h is not None:
                        holders.append((h, tg == 'RowInfo'))
            else:
                h = sec.find('Cells')
                if h is not None:
                    holders.append((h, False))
            for holder, is_col in [(h, ic) for h, ic in holders]:
              for ce in holder.findall('Cell'):
                c = Cell()
                c.sheet = si
                c.section = sec_i
                c.dynamic = is_col
                c.path = ce.get('Path') or ''
                if not c.path:
                    continue
                c.ctl = ce.get('Controltype')
                c.status = ce.get('StatusID')
                c.minv = ce.get('MinValue')
                c.maxv = ce.get('MaxValue')
                c.tpath = tpath
                c.row0, c.row0_2 = row0, row02
                c.region = holder.tag
                if is_col:
                    c.cid = ce.get('CellID') or None
                    c.cid2 = ce.get('CellID2') or None
                    if tpath.startswith('HSoThueDTu'):
                        c.full = tpath + '/' + c.path
                    else:
                        c.full = (parent + '/' + tpath + '/' + c.path) if parent else (tpath + '/' + c.path)
                else:
                    c.cid = parse_cellid(ce.get('CellID'))
                    c.cid2 = parse_cellid(ce.get('CellID2'))
                    if c.path.startswith('HSoThueDTu'):
                        c.full = c.path
                    else:
                        if node and c.path.split('/')[0] == node.rsplit('/', 1)[-1]:
                            c.full = (parent + '/' + c.path) if parent else c.path
                        else:
                            c.full = (node + '/' + c.path) if node else c.path
                cells.append(c)
    return cells, log


def find_sections(rows, folder_dst_cache, mafile, ver, sheets):
    """Ánh xạ sheet idx -> đường dẫn tệp Sections."""
    vnd = nodots(ver)
    secs = {}
    for r in rows:
        if r['root'] == 'Sections' and r['magic'] == 'XML':
            secs[os.path.basename(r['dst']).lower()] = r['dst']
    out = {}
    for si, sh in enumerate(sheets):
        D = (sh.get('DataFile') or '').lower()
        cands = []
        if si == 0:
            cands += [b for b in secs if b.endswith('_v%s_mau-xml.xml' % ver.lower())]
        stem = D
        for s in (stem, stem[1:], stem[2:]):
            if s:
                cands.append('%s_%s.xml' % (s, vnd))
        for c in cands:
            if c in secs:
                out[si] = os.path.join(ROOT, '_forms', secs[c])
                break
            # tên tệp trích bị cắt còn 27 ký tự
            pre = [k for k in secs if len(k) == 27 and c.startswith(k)]
            if len(set(pre)) == 1:
                out[si] = os.path.join(ROOT, '_forms', secs[pre[0]])
                break
    return out


# ---------------------------------------------------------------- phân giải
def path_index(cells):
    by_sheet = defaultdict(list)
    for c in cells:
        by_sheet[c.sheet].append(c)
    return by_sheet


def resolve_path(cells_sheet, P):
    """Tìm ô tĩnh theo đường dẫn P (khớp đủ hoặc hậu tố)."""
    ex = [c for c in cells_sheet if not c.dynamic and c.full == P]
    if not ex:
        ex = [c for c in cells_sheet if not c.dynamic and c.full.endswith('/' + P)]
    if len(ex) == 1:
        return ex[0], None
    if not ex:
        dyn = [c for c in cells_sheet if c.dynamic and (c.full == P or c.full.endswith('/' + P))]
        return None, 'placeholder_path_dong_dong' if dyn else 'placeholder_path_khong_co_trong_xml'
    return None, 'placeholder_path_mo_ho'


PH_COLROW = re.compile(r'\{col\(colIndex\(([^{}]+?)\)\)\}\{rowIndex\(\1\)\}')
PH_ROW = re.compile(r'([A-Z]+)\{rowIndex\(([^{}]+?)\)\}')
PH_COLONLY = re.compile(r'\{col\(colIndex\(([^{}]+?)\)\)\}')


def classify_ph(t):
    if 'trigger.' in t:
        return 'placeholder_trigger_row_col'
    if 'staticCell' in t:
        return 'placeholder_staticCell'
    if 'SheetInfos' in t or 'RowsValues' in t or 'Location' in t:
        return 'placeholder_vi_tri_bang_dong'
    if re.search(r'\{[^{}]*\b(i|j|k|n|row|num\d*|iRow|iCol)\b', t):
        return 'placeholder_bien_vong_lap_hoac_hang'
    return 'placeholder_chay_luc_chay_khac'


def resolve_text(s, cells_sheet, space):
    """Thay placeholder dạng path trong chuỗi. Trả (chuỗi, lý do lỗi|None)."""
    err = []

    def cell_ref(c):
        return c.cid if space == 0 else c.cid2

    def f1(m):
        c, e = resolve_path(cells_sheet, m.group(1))
        if c is None or not cell_ref(c):
            err.append(e or 'placeholder_o_khong_co_toa_do')
            return m.group(0)
        return cell_ref(c)

    def f2(m):
        c, e = resolve_path(cells_sheet, m.group(2))
        if c is None or not cell_ref(c):
            err.append(e or 'placeholder_o_khong_co_toa_do')
            return m.group(0)
        mm = A1.match(cell_ref(c))
        return m.group(1) + mm.group(2)

    s2 = PH_COLROW.sub(f1, s)
    s2 = PH_ROW.sub(f2, s2)
    if '{' in s2:
        if not err:
            err.append(classify_ph(s2))
        return s2, err[0]
    return s2, None


def cell_matches_coord(c, space, sheet, coord):
    """True nếu ô c nằm ở toạ độ literal (A1) trong không gian space."""
    if c.sheet != sheet:
        return False
    cid = c.cid if space == 0 else c.cid2
    if not cid:
        return False
    if not c.dynamic:
        return cid == coord
    m = A1.match(coord)
    if not m:
        return False
    row0 = c.row0 if space == 0 else c.row0_2
    return m.group(1) == cid and row0 is not None and int(m.group(2)) == row0


def match_literal(cells, space, sheet, coord):
    m = RANGE.match(coord)
    out = []
    if m:
        c1, r1, c2, r2 = col2n(m.group(1)), int(m.group(2)), col2n(m.group(3)), int(m.group(4))
        for c in cells:
            if c.sheet != sheet:
                continue
            cid = c.cid if space == 0 else c.cid2
            if not cid:
                continue
            if c.dynamic:
                row0 = c.row0 if space == 0 else c.row0_2
                if row0 is not None and c1 <= col2n(cid) <= c2 and r1 <= row0 <= r2:
                    out.append(c)
            else:
                mm = A1.match(cid)
                if mm and c1 <= col2n(mm.group(1)) <= c2 and r1 <= int(mm.group(2)) <= r2:
                    out.append(c)
        return out
    for c in cells:
        if cell_matches_coord(c, space, sheet, coord):
            out.append(c)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=os.path.join(ROOT, '_spec', 'coverage'))
    ap.add_argument('--sample', type=int, default=0)
    ap.add_argument('--seed', type=int, default=20260511)
    ap.add_argument('--t2', default=T2)
    ap.add_argument('--t3', default=T3)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    forms_dir = os.path.join(ROOT, '_forms')

    menu = load_menu()
    mrows = load_manifest(forms_dir)
    mby = manifest_by_form(mrows)
    src2dst = {}
    for r in mrows:
        src2dst.setdefault(r['src'], r['dst'])
    fby = load_formulas(os.path.join(ROOT, '_spec', 'formulas'))
    rr = load_rowreplace(os.path.join(ROOT, '_spec', 'config', 'mapcongthucchosua.xml.json'))
    t2, t2_un = load_t2(a.t2)
    t3, t3_un = load_t3(a.t3)
    xidx = jload(os.path.join(ROOT, '_spec', 'xsd', 'index.json'))
    xsd_by = {}
    xsd_khbs = {}
    xsd_old = {}
    for f in xidx['files']:
        k = f.get('key')
        if k and k.get('ma_file') and k.get('version_xml'):
            kh = 'VersionKHBS' in (k.get('origin') or '') or '_KHBS_' in f['json'].upper()
            xsd_old[(k['ma_file'], k['version_xml'])] = f['json']
            (xsd_khbs if kh else xsd_by)[(k['ma_file'], k['version_xml'])] = f['json']

    G = Counter()  # tổng
    unmapped = defaultdict(Counter)  # nhóm -> lý do -> số
    conflicts = defaultdict(list)
    summary = []
    crosstab = Counter()
    ctl16 = Counter()
    space_choice = Counter()
    sheet_log = []
    all_cells_flat = []
    indep_total = 0
    nform = 0

    for key in sorted(menu):
        mafile, ver = key
        info = menu[key]
        nform += 1
        rows = mby.get((mafile, nodots(ver)), [])
        sec_files = find_sections(rows, None, mafile, ver, info['sheets'])
        cells, log = build_cells(info['sheets'], sec_files, ver)
        for l in log:
            l.update({'form': '%s_%s' % key})
            if l.get('loi') == 'khong_tim_thay_tep_Sections' and 'KHBS' in (l.get('ID') or '').upper():
                allrows = [r for (mf, _v), rs in mby.items() if mf == mafile for r in rs]
                kx = sorted({os.path.basename(r['dst']) for r in allrows if 'khbs' in r['dst'].lower() and r['root'] == 'Sections'})
                ks = sorted({os.path.basename(r['dst']) for r in allrows if 'khbs' in r['dst'].lower() and r['root'] == 'schema'})
                l['nhom_khbs'] = ('khbs_chi_co_Sections_phien_ban_khac' if kx else
                                  ('khbs_chi_co_xsd_khong_co_Sections' if ks else 'khbs_khong_co_tep_nao'))
                l['khbs_sections_ton_tai'] = kx
                l['khbs_xsd_ton_tai'] = ks
            sheet_log.append(l)
        khbs_sheets = {l['sheet'] for l in log if l.get('nhom_khbs')}
        khbs = {}

        def kadd(sheet, kind, ref, field, val=1):
            e = khbs.setdefault((sheet, kind, ref), {'nguon_tinh': set(), 'nguon_kiem': set(), 'lock': 0, 'unlock': 0, 'setvalue': 0})
            if field in ('nguon_tinh', 'nguon_kiem'):
                e[field].add(val)
            else:
                e[field] += 1
        # độc lập: đếm Cell có Path bằng regex
        for p in sorted(set(sec_files.values())):  # đếm theo form (tệp dùng chung tính lại mỗi form)
            t = open(p, 'rb').read()
            t = t.decode('utf-16') if t[:2] in (b'\xff\xfe', b'\xfe\xff') else t.decode('utf-8-sig')
            t = re.sub(r'<!--.*?-->', '', t, flags=re.S)
            indep_total += len(re.findall(r'<Cell\b[^<]*?\bPath\s*=\s*"[^"]', t))

        # XSD
        xs = {}
        xj = xsd_by.get(key)
        if xsd_old.get(key) and 'KHBS' in xsd_old[key].upper():
            G['form_xsd_cu_dung_khbs_cho_sheet_chinh'] += 1
        xj_fb = None
        if xj is None:
            G['form_khong_co_xsd_chinh'] += 1
            base = mafile.replace('_VIEW', '')
            if base != mafile and xsd_by.get((base, ver)):
                xj = xj_fb = xsd_by[(base, ver)]
                G['form_xsd_theo_form_goc_khong_VIEW'] += 1
        if xj and os.path.exists(os.path.join(ROOT, '_spec', 'xsd', xj)):
            for e in jload(os.path.join(ROOT, '_spec', 'xsd', xj))['elements']:
                xs[e['path'].lstrip('/')] = e
        for c in cells:
            e = xs.get(c.full)
            if e is None and c.path:
                # Path/TablePath đã bắt đầu từ gốc XSD (ví dụ TKhai/DLTKhai/...) hoặc ghép bị lặp đoạn liền kề
                cands = [c.path, (c.tpath + '/' + c.path) if c.dynamic else None]
                parts = c.full.split('/')
                col = [q for i, q in enumerate(parts) if i == 0 or q != parts[i - 1]]
                cands.append('/'.join(col))
                for cand in cands:
                    if cand and cand in xs:
                        e = xs[cand]
                        c.full = cand
                        break
            if e is None and c.path and not c.dynamic:
                node0 = info['sheets'][c.sheet].get('XMLnode') or ''
                par0 = node0.rsplit('/', 1)[0] if '/' in node0 else ''
                tp0 = c.tpath.rsplit('/', 1)[0] if '/' in c.tpath else ''
                for alt in (par0 + '/' + c.path, par0 + '/' + tp0 + '/' + c.path, par0 + '/' + c.tpath + '/' + c.path,
                            'HSoThueDTu/HSoKhaiThue/' + c.path,
                            'HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/' + c.path):
                    if alt in xs:
                        e = xs[alt]
                        c.full = alt
                        break
            if e is None and c.path and not c.dynamic:
                # hậu tố duy nhất trong nhánh XMLnode
                node = info['sheets'][c.sheet].get('XMLnode') or ''
                cand = [p for p in xs if p.startswith(node + '/') and p.endswith('/' + c.path)] if node else []
                if len(cand) == 1:
                    e = xs[cand[0]]
                    c.full = cand[0]
            if e is None and c.path:
                # hậu tố duy nhất toàn XSD (XMLnode trong menu có thể thiếu đoạn PLuc/...)
                suf = (c.tpath + '/' + c.path) if c.dynamic else c.path
                if '/' in suf:
                    cand = [p for p in xs if p.endswith('/' + suf)]
                    if len(cand) == 1:
                        e = xs[cand[0]]
                        c.full = cand[0]
            c.in_xsd = e is not None and e.get('kind') != 'attribute'
            c.xsd = e if c.in_xsd else None
            if not c.in_xsd:
                pp = c.full.split('/')
                c.full = '/'.join(q for i, q in enumerate(pp) if i == 0 or q != pp[i - 1])

        data = [c for c in cells if c.in_xsd]
        off = [c for c in cells if not c.in_xsd]
        by_sheet = path_index(data)

        # ---- công thức Excel (T1)
        fx = {}
        sheetnames = {}
        for kind in ('xls', 'xls1'):
            ent = fby.get(key, {}).get(kind)
            fx[kind] = {}
            if not ent:
                continue
            j = jload(os.path.join(ROOT, '_spec', 'formulas', ent['json']))
            sheetnames[kind] = [s['name'] for s in j['sheets']]
            for x in j['formulas']:
                fx[kind][(x['sheet'], x['cell'])] = x['formula']
        for kind in ('xls', 'xls1'):
            nm = sheetnames.get(kind)
            if nm is None:
                continue
            for si, sh in enumerate(info['sheets']):
                if si >= len(nm) or nm[si].replace('-', '_').replace('/', '_').lower() != (sh['ID'] or '').replace('-', '_').lower():
                    nz = lambda x: (x or '').replace('-', '_').replace('/', '_').lower()
                    ids = [nz(q['ID']) for q in info['sheets']]
                    nms = [nz(q) for q in nm]
                    if si >= len(nm):
                        loai = 'thieu_sheet_trong_xls'
                    elif nz(sh['ID']) in nms:
                        loai = 'lech_chi_so'
                    else:
                        loai = 'chi_khac_ten'
                    sheet_log.append({'form': '%s_%s' % key, 'sheet': si, 'ID': sh['ID'], 'phan_loai': loai,
                                      'loi': 'ten_sheet_%s_lech: %s' % (kind, nm[si] if si < len(nm) else None)})

        def xl(kind, c):
            nm = sheetnames.get(kind)
            if nm is None or c.sheet >= len(nm):
                return None
            cid = c.cid if kind == 'xls' else c.cid2
            if not cid:
                return None
            if c.dynamic:
                row0 = c.row0 if kind == 'xls' else c.row0_2
                if row0 is None:
                    return None
                cid = '%s%d' % (cid, row0)
            return fx[kind].get((nm[c.sheet], cid))

        for c in data:
            for kind in ('xls', 'xls1'):
                f = xl(kind, c)
                if f is not None:
                    c.calc[kind].append(f)

        # ---- ROWREPLACE (T4)
        for it in rr.get(info['menu_id'], []):
            try:
                si = int(it.get('SheetIndex'))
            except Exception:
                continue
            if it.get('Dynamic') == '1':
                P = (it.get('TablePath') or '') + '/' + (it.get('MCTPath') or '')
            else:
                P = it.get('MCTPath') or ''
            hit = [c for c in data if c.sheet == si and (c.full == P or c.full.endswith('/' + P))]
            if hit:
                for c in hit:
                    c.calc['rowreplace'].append(it.get('Formula'))
            else:
                unmapped['rowreplace'][('khong_co_trong_xml' if True else '')] += 1

        # ---- C# (T2)
        recs = t2.get(key, [])
        lit_hits = Counter()
        # xác định không gian toạ độ chiếm ưu thế
        probes = []
        for r in recs:
            if r.get('unreachable'):
                continue
            for e in r['effects']:
                t = e.get('target')
                if isinstance(t, dict) and isinstance(t.get('sheet'), int) and A1.match(t.get('cell') or ''):
                    probes.append((t['sheet'], t['cell']))
        for ch in t3.get(key, []):
            for x in ch.get('cells_marked', []):
                m = re.match(r'^S(\d+)!([A-Z]+\d+)$', x)
                if m:
                    probes.append((int(m.group(1)), m.group(2)))
        for sp in (0, 1):
            for (s, co) in probes:
                if any(cell_matches_coord(c, sp, s, co) for c in data):
                    lit_hits[sp] += 1
        # C# sửa bảng tính giao diện (frmWinToKhai.cs:417 WinLoadFromFile(InterfaceTemplateExcel)); CellID = toạ độ
        # mẫu giao diện (TemplateRowStart), CellID2 = mẫu báo cáo (ReportRowStart). Cố định không gian xls (CellID).
        space = 0
        space_choice['xls'] += 1
        if lit_hits[1] > lit_hits[0]:
            space_choice['canh_bao_xls1_khop_nhieu_hon'] += 1
        spname = 'xls1' if space else 'xls'

        def cell_ref(c):
            return c.cid2 if space else c.cid

        for r in recs:
            if r.get('unreachable'):
                G['csharp_unreachable_bo_qua'] += len(r['effects'])
                continue
            for e in r['effects']:
                op = e['op']
                t = e.get('target')
                G['csharp_effect'] += 1
                if not isinstance(t, dict) or t.get('cell') is None:
                    unmapped['csharp'][op + ':khong_co_target'] += 1
                    continue
                sh = t.get('sheet')
                cs = t.get('cell')
                if isinstance(sh, str) and 'activeSheet' in sh:
                    mg = [re.fullmatch(r'activeSheet\(\) == (\d+)', q.strip()) for q in (r.get('guards') or [])]
                    mg = [q for q in mg if q]
                    if len(mg) == 1:
                        sh = int(mg[0].group(1))
                        G['csharp_effect_activeSheet_giai_duoc'] += 1
                if not isinstance(sh, int):
                    unmapped['csharp']['sheet_khong_phai_so:%s' % type(sh).__name__ if not isinstance(sh, str) else 'sheet_dong:%s' % ('sView' if 'sView' in sh else 'activeSheet' if 'active' in sh else 'khac')] += 1
                    continue
                if re.match(r'^\d+:\d+$', cs) or re.match(r'^[A-Z]+:[A-Z]+$', cs):
                    unmapped['csharp']['hang_cot_ca_dong'] += 1
                    continue
                targets = []
                if '{' in cs:
                    cs2, err = resolve_text(cs, by_sheet.get(sh, []), space)
                    if err:
                        unmapped['csharp'][err] += 1
                        continue
                    cs = cs2
                if sh in khbs_sheets and (A1.match(cs) or RANGE.match(cs)):
                    G['csharp_effect_khbs_theo_toa_do'] += 1
                    if op == 'setFormula' and e.get('formula') not in (None, ''):
                        kadd(sh, 'xls', cs, 'nguon_tinh', 'csharp')
                    elif op == 'lock':
                        kadd(sh, 'xls', cs, 'lock')
                    elif op == 'unlock':
                        kadd(sh, 'xls', cs, 'unlock')
                    elif op == 'setValue':
                        kadd(sh, 'xls', cs, 'setvalue')
                    continue
                hit = match_literal(data, space, sh, cs)
                if not hit:
                    unmapped['csharp']['o_literal_khong_phai_o_du_lieu_XML'] += 1
                    continue
                G['csharp_effect_anh_xa'] += 1
                for c in hit:
                    if op == 'setFormula':
                        f = e.get('formula')
                        if f in (None, ''):
                            c.calc['csharp_reset'].append({'hook': r['hook']})
                            continue
                        ftxt, err = resolve_text(str(f), by_sheet.get(sh, []), space) if '{' in str(f) else (str(f), None)
                        c.calc['csharp'].append({
                            'formula': ftxt, 'norm': norm_formula(ftxt) if not err else None,
                            'err': err, 'hook': r['hook'], 'kind': r['kind'], 'at': e.get('at') or 0,
                            'guards': r.get('guards') or [],
                            'when': (r.get('trigger') or {}).get('when') or [],
                            'src': '%s:%s' % (r['source']['file'], '-'.join(str(x) for x in r['source']['lines']))})
                    elif op == 'setValue':
                        c.calc['csharp_value'].append({'hook': r['hook'], 'value': str(e.get('value'))[:80]})
                        if re.fullmatch(r'-?\d+(\.\d+)?|"[^"{}]*"|\'[^\'{}]*\'', str(e.get('value')).strip()):
                            c.calc['csharp_const'].append(str(e.get('value'))[:40])
                    elif op == 'lock':
                        c.locks['lock'] += 1
                        c.calc['lock_info'].append({'guards': r.get('guards') or [],
                                                    'when': (r.get('trigger') or {}).get('when') or [],
                                                    'hook': r['hook'], 'nguon': '%s:%s' % (r['source']['file'], '-'.join(str(q) for q in r['source']['lines']))})
                    elif op == 'unlock':
                        c.locks['unlock'] += 1
                    else:
                        c.calc['csharp_khac_%s' % op].append(1)

        # ---- Validate (T3)
        for ch in t3.get(key, []):
            G['t3_check'] += 1
            cm = ch.get('cells_marked') or []
            if not cm:
                unmapped['t3']['khong_co_cells_marked'] += 1
                continue
            for x in cm:
                m = re.match(r'^S(\d+)!(.+)$', x)
                if not m:
                    unmapped['t3']['sheet_dong(S[..])'] += 1
                    continue
                s, co = int(m.group(1)), m.group(2)
                if '{' in co or '(' in co:
                    unmapped['t3']['toa_do_chay_luc_chay'] += 1
                    continue
                if s in khbs_sheets:
                    G['t3_cell_khbs_theo_toa_do'] += 1
                    kadd(s, 'xls', co, 'nguon_kiem', 'validate')
                    continue
                hit = match_literal(data, space, s, co)
                if not hit:
                    conflicts['validate_o_khong_co_trong_xml'].append({
                        'form': '%s_%s' % key, 'sheet': s, 'cell': co, 'space': spname,
                        'message_id': (ch.get('message') or {}).get('id') if isinstance(ch.get('message'), dict) else None,
                        'nguon': '%s:%s' % (ch['source']['file'], '-'.join(str(q) for q in ch['source']['lines']))})
                    unmapped['t3']['o_literal_khong_co_trong_xml'] += 1
                    continue
                G['t3_cell_anh_xa'] += 1
                for c in hit:
                    c.chk['validate'].append((ch.get('message') or {}).get('id') if isinstance(ch.get('message'), dict) else None)

        # ---- DV (T9)
        ent = fby.get(key, {}).get('xls1')
        dvfile = None
        if ent:
            dst = src2dst.get(ent['source_name'])
            if dst:
                dvfile = os.path.join(ROOT, '_spec', 'dv', '_forms__' + dst.replace('/', '__') + '.json')
        dvs = []
        if dvfile and os.path.exists(dvfile):
            dvs = jload(dvfile).get('dv', [])
        for c in data:
            if not c.cid2:
                continue
            m = A1.match(c.cid2) if not c.dynamic else None
            if c.dynamic:
                if c.row0_2 is None:
                    continue
                colN, rowN = col2n(c.cid2), c.row0_2
            else:
                if not m:
                    continue
                colN, rowN = col2n(m.group(1)), int(m.group(2))
            for d in dvs:
                if d.get('sheet_index') != c.sheet:
                    continue
                for rg in d['ranges']:
                    if rg['col_first'] + 1 <= colN <= rg['col_last'] + 1 and rg['row_first'] + 1 <= rowN <= rg['row_last'] + 1:
                        c.chk['dv'].append(d['type'])
                        break

        # ---- KHBS: phủ theo toạ độ (sheet, ô) từ T1 khi không có Sections
        for kind in ('xls', 'xls1'):
            nm = sheetnames.get(kind)
            if not nm:
                continue
            for si in sorted(khbs_sheets):
                if si >= len(nm):
                    continue
                for (shn, cell), f in fx[kind].items():
                    if shn == nm[si]:
                        kadd(si, kind, cell, 'nguon_tinh', kind)
        for l in log:
            if l.get('nhom_khbs'):
                l['phu_toa_do'] = len([1 for k in khbs if k[0] == l['sheet']])
                l['nguon_mau_khbs'] = ('sheet %s trong .xls của form (T1, T2, T3)' % l['ID'])
                if info.get('kieu_khbs') == '3':
                    l['quy_tac_thay_sheet'] = ('WinFnProcess.cs:69-80: KieuKHBS=="3" và LoaiToKhai=="B": 2 sheet cuối (trước Header) '
                                               'bị thay bằng sheet 0,1 của _forms/_ngoai_menu/KHBS_TT80*/khbs_tt80_<VersionKHBS>.xls; chưa phủ ô của mẫu chung')
                if not l['phu_toa_do']:
                    l['khbs_khong_phu_ly_do'] = 'sheet không có công thức T1, không có effect C# và không có check Validate theo toạ độ literal'
        khbs_out = [{'sheet': k[0], 'sheet_id': info['sheets'][k[0]]['ID'], 'khong_gian': k[1], 'cell': k[2],
                     'nguon_tinh': sorted(v['nguon_tinh']), 'nguon_kiem': sorted(v['nguon_kiem']),
                     'khoa': {'lock': v['lock'], 'unlock': v['unlock']}, 'csharp_setvalue': v['setvalue'],
                     'nhan': 'tinh_tu_dong_hoac_ca_hai' if v['nguon_tinh'] else 'nhap_tay_hoac_chi_kiem'}
                    for k, v in sorted(khbs.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2]))]
        G['khbs_o_theo_toa_do'] += len(khbs_out)
        G['khbs_sheet_co_toa_do'] += len([1 for l in log if l.get('phu_toa_do')])
        G['khbs_sheet_khong_phu'] += len([1 for l in log if l.get('nhom_khbs') and not l.get('phu_toa_do')])

        # ---- XSD + nhãn + xung đột
        out_cells = []
        counts = Counter()
        for c in data:
            xe = c.xsd
            c.chk['xsd'] = [{'type': xe.get('type'), 'pattern': xe.get('pattern'),
                            'required': xe.get('required'), 'repeatable': xe.get('repeatable'),
                            'nillable': xe.get('nillable_effective')}]
            sources = []
            if c.calc.get('xls'):
                sources.append('xls')
            if c.calc.get('xls1'):
                sources.append('xls1')
            if c.calc.get('rowreplace'):
                sources.append('rowreplace')
            if c.calc.get('csharp'):
                sources.append('csharp')
            has_unlock = c.locks['unlock'] > 0
            is_input = c.status in INPUT_STATUS
            hasx = bool(c.calc.get('xls') or c.calc.get('xls1'))
            crosstab[(c.status, 'co_cong_thuc_excel' if hasx else 'khong')] += 1
            if not sources:
                label = 'nhap_tay'
            elif is_input or has_unlock:
                label = 'ca_hai'
            else:
                label = 'tinh_tu_dong'
            counts[label] += 1
            if c.dynamic:
                G['dyn_o'] += 1
                if c.row0 is None or c.row0_2 is None:
                    G['dyn_thieu_hang_mau'] += 1
                if not (c.cid and re.fullmatch(r'[A-Z]+', c.cid)):
                    G['dyn_cot_khong_hop_le'] += 1
                if c.calc.get('xls'):
                    G['dyn_co_T1_xls_hang_mau'] += 1
                if c.calc.get('xls1'):
                    G['dyn_co_T1_xls1_hang_mau'] += 1
                if c.calc.get('csharp') or c.calc.get('csharp_value') or c.locks['lock'] or c.locks['unlock']:
                    G['dyn_co_T2_hang_mau'] += 1
                if c.calc.get('rowreplace'):
                    G['dyn_co_rowreplace'] += 1
                if c.chk.get('validate'):
                    G['dyn_co_validate_hang_mau'] += 1
            if c.ctl == '16':
                ctl16[label] += 1
            cks = []
            if c.chk.get('validate'):
                cks.append('validate')
            if c.chk.get('dv'):
                cks.append('dv')
            if c.xsd:
                cks.append('xsd')
            cid_disp = ('COT:%s' % c.cid if c.dynamic else c.cid)
            cid2_disp = ('COT:%s' % c.cid2 if c.dynamic and c.cid2 else c.cid2)
            rec = {
                'sheet': c.sheet, 'sheet_id': info['sheets'][c.sheet]['ID'],
                'cell_xls': cid_disp, 'cell_xls1': cid2_disp, 'dong_dong': c.dynamic, 'khu_vuc': c.region,
                'path': c.full, 'controltype': c.ctl, 'status_id': c.status,
                'nguon_tinh': sources, 'nguon_kiem': cks, 'nhan': label,
                'xls': c.calc.get('xls', [None])[0], 'xls1': c.calc.get('xls1', [None])[0],
                'rowreplace': c.calc.get('rowreplace', []),
                'csharp': [{k: v for k, v in x.items()} for x in c.calc.get('csharp', [])],
                'csharp_setvalue': len(c.calc.get('csharp_value', [])), 'co_dinh_gia_tri': bool(c.calc.get('csharp_const')),
                'csharp_setformula_reset': len(c.calc.get('csharp_reset', [])),
                'khoa': {'lock': c.locks['lock'], 'unlock': c.locks['unlock']},
                'validate': c.chk.get('validate', []), 'dv': sorted(set(c.chk.get('dv', []))),
                'xsd': c.chk['xsd'][0],
            }
            out_cells.append(rec)
            tag = '%s_%s|S%d|%s' % (mafile, ver, c.sheet, cid_disp or cid2_disp or c.full)
            all_cells_flat.append((tag, rec))
            # xung đột
            cs = c.calc.get('csharp', [])
            exf = c.calc.get(spname)
            exn = norm_formula(exf[0]) if exf else None
            if cs:
                norms = sorted({x['norm'] for x in cs if x['norm']})
                if exn is not None:
                    # chỉ lấy lần ghi cuối của mỗi (hook, tệp, guards, when): lần ghi sau ghi đè lần trước
                    last = {}
                    for x in cs:
                        gk = (x['hook'], x['src'].split(':')[0], tuple(x['guards']), tuple(x['when']))
                        if gk not in last or x['at'] >= last[gk]['at']:
                            last[gk] = x

                    def tuong_duong(x):
                        if x['norm'] == exn:
                            return True
                        if re.fullmatch(r'[A-Z]+\d+', x['norm']):
                            nmx = sheetnames.get(spname)
                            if nmx and c.sheet < len(nmx):
                                f2 = fx[spname].get((nmx[c.sheet], x['norm']))
                                if f2 is not None and norm_formula(f2) == exn:
                                    return True
                        return False
                    diff = [x for x in last.values() if x['norm'] and not tuong_duong(x)]
                    if diff:
                        conflicts['excel_khac_csharp'].append({
                            'form': '%s_%s' % key, 'sheet': c.sheet, 'cell': cid_disp if not space else cid2_disp,
                            'space': spname, 'path': c.full, 'excel': exf[0],
                            'csharp': [{'formula': x['formula'], 'hook': x['hook'], 'guards': x['guards'],
                                        'when': x['when'], 'nguon': x['src']} for x in sorted(diff, key=lambda q: q['src'])],
                            'nguon_excel': 'T1 %s' % fby[key][spname]['json']})
                un = [x for x in cs if x['norm'] and not x['guards'] and not x['when']]
                if len({x['norm'] for x in un}) >= 2:
                    conflicts['nhieu_csharp_khong_dieu_kien'].append({
                        'form': '%s_%s' % key, 'sheet': c.sheet, 'cell': cid_disp if not space else cid2_disp,
                        'path': c.full,
                        'csharp': [{'formula': x['formula'], 'hook': x['hook'], 'nguon': x['src']} for x in un]})
            if c.locks['lock'] > 0 and c.locks['unlock'] == 0 and c.status in INPUT_STATUS:
                li = c.calc.get('lock_info', [])
                cond = bool(li) and all(x['guards'] or x['when'] for x in li)
                conflicts['khoa_co_dieu_kien' if cond else 'khoa_nhung_la_o_nhap'].append({
                    'form': '%s_%s' % key, 'sheet': c.sheet, 'cell': cid_disp if not space else cid2_disp,
                    'path': c.full, 'status_id': c.status, 'lock': c.locks['lock'],
                    'co_nguon_tinh': sources, 'co_dinh_gia_tri': c.calc.get('csharp_const', [])[:3],
                    'guard_khoa': [{'guards': x['guards'], 'when': x['when'], 'nguon': x['nguon']} for x in li[:3]]})

        # ô có Path nhưng không vào XSD
        off_recs = [{'sheet': c.sheet, 'cell_xls': c.cid, 'cell_xls1': c.cid2, 'path': c.full,
                     'controltype': c.ctl, 'dong_dong': c.dynamic} for c in off]
        G['cell_path_ngoai_xsd'] += len(off)
        out = {'header': {'form': {'MaFile': mafile, 'VersionXML': ver, 'MenuID': info['menu_id']},
                          'toa_do_csharp': spname, 'xsd_json': xj, 'xsd_thay_the_form_goc': bool(xj_fb), 'tep_Sections': {str(k): os.path.relpath(v, ROOT) for k, v in sorted(sec_files.items())},
                          'dinh_nghia': 'xem coverage.py'},
               'cells': out_cells, 'path_ngoai_xsd': off_recs, 'khbs_cells_theo_toa_do': khbs_out}
        with open(os.path.join(a.out, '%s_%s.json' % (mafile, ver)), 'w', encoding='utf-8') as f:
            json.dump(out, f, ensure_ascii=False, sort_keys=True, indent=1)
            f.write('\n')
        tot = len(out_cells)
        row = {'MaFile': mafile, 'VersionXML': ver, 'MenuID': info['menu_id'], 'so_sheet': len(info['sheets']),
               'sheet_co_xml': len(sec_files), 'o_du_lieu': tot, 'o_path_ngoai_xsd': len(off),
               'tinh_tu_dong': counts['tinh_tu_dong'], 'nhap_tay': counts['nhap_tay'], 'ca_hai': counts['ca_hai'],
               'co_validate': sum(1 for r in out_cells if 'validate' in r['nguon_kiem']),
               'co_dv': sum(1 for r in out_cells if 'dv' in r['nguon_kiem']),
               'co_xsd': tot, 'toa_do_csharp': spname}
        summary.append(row)
        for k in ('tinh_tu_dong', 'nhap_tay', 'ca_hai'):
            G['nhan_' + k] += counts[k]
        G['o_du_lieu'] += tot

    # nhóm C# không thuộc form
    for fn, recs in t2_un:
        for r in recs:
            for e in r['effects']:
                G['csharp_effect_khong_thuoc_form'] += 1
                unmapped['csharp']['lop_ngoai_menu_hoac_dung_chung(%s)' % ('shared/app' if fn.startswith('_') or fn.startswith('General') else 'ngoai_menu')] += 1
    for fn, chs in t3_un:
        for ch in chs:
            G['t3_check_khong_thuoc_form'] += 1
            unmapped['t3']['lop_ngoai_menu'] += 1

    with open(os.path.join(a.out, 'summary.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(summary[0].keys()), lineterminator='\n')
        w.writeheader()
        for r in summary:
            w.writerow(r)

    for k in conflicts:
        conflicts[k].sort(key=lambda x: json.dumps(x, sort_keys=True, ensure_ascii=False))
    cj = {'dinh_nghia': {
        'excel_khac_csharp': 'ô có công thức Excel (cùng không gian toạ độ C#) và C# setFormula khác sau chuẩn hóa (bỏ khoảng trắng, $, =, chữ hoa); C# setFormula rỗng (reset) không tính',
        'nhieu_csharp_khong_dieu_kien': '>=2 công thức C# chuẩn hóa khác nhau, guards rỗng và trigger.when rỗng',
        'khoa_nhung_la_o_nhap': 'có effect lock không điều kiện (mọi lần khoá không có guards và trigger.when), không có unlock cùng ô, StatusID trong %s' % sorted(INPUT_STATUS),
        'khoa_co_dieu_kien': 'như trên nhưng mọi lần khoá đều có guards hoặc trigger.when (kèm guard_khoa; co_dinh_gia_tri = setValue hằng cùng ô)',
        'validate_o_khong_co_trong_xml': 'check T3 có cells_marked literal S<n>!A1 không trùng ô dữ liệu nào của XML mẫu'},
        'so_luong': {k: len(v) for k, v in sorted(conflicts.items())},
        'unmapped': {g: dict(sorted(v.items())) for g, v in sorted(unmapped.items())},
        'xung_dot': {k: v for k, v in sorted(conflicts.items())}}
    with open(os.path.join(a.out, 'conflicts.json'), 'w', encoding='utf-8') as f:
        json.dump(cj, f, ensure_ascii=False, sort_keys=True, indent=1)
        f.write('\n')
    with open(os.path.join(a.out, 'sheet_log.json'), 'w', encoding='utf-8') as f:
        json.dump(sorted(sheet_log, key=lambda x: json.dumps(x, sort_keys=True, ensure_ascii=False)), f,
                  ensure_ascii=False, sort_keys=True, indent=1)
        f.write('\n')

    # ---- báo cáo
    print('so_form_xu_ly', nform)
    print('form_xsd_cu_dung_khbs (truoc)', G['form_xsd_cu_dung_khbs_cho_sheet_chinh'], '| form_khong_co_xsd_chinh (sau)', G['form_khong_co_xsd_chinh'])
    print('tong_o_du_lieu', G['o_du_lieu'], '| o_path_ngoai_xsd', G['cell_path_ngoai_xsd'])
    print('nhan', {k: G['nhan_' + k] for k in ('tinh_tu_dong', 'nhap_tay', 'ca_hai')},
          'tong_nhan', sum(G['nhan_' + k] for k in ('tinh_tu_dong', 'nhap_tay', 'ca_hai')))
    print('controltype16_theo_nhan', dict(ctl16), 'tong', sum(ctl16.values()))
    print('kiem_cheo_regex_Cell_co_Path (tong ca ngoai XSD)', indep_total,
          '== o_du_lieu + ngoai_xsd?', indep_total == G['o_du_lieu'] + G['cell_path_ngoai_xsd'])
    print('xung_dot', {k: len(v) for k, v in sorted(conflicts.items())})
    print('dong_dong', {k: v for k, v in sorted(G.items()) if k.startswith('dyn_')})
    print('khbs: o_theo_toa_do', G['khbs_o_theo_toa_do'], 'sheet_co_toa_do', G['khbs_sheet_co_toa_do'], 'sheet_khong_phu', G['khbs_sheet_khong_phu'], '| effect_khbs', G['csharp_effect_khbs_theo_toa_do'], 't3_khbs', G['t3_cell_khbs_theo_toa_do'], '| activeSheet_giai_duoc', G['csharp_effect_activeSheet_giai_duoc'])
    print('csharp_effect', G['csharp_effect'], 'anh_xa', G['csharp_effect_anh_xa'],
          'unreachable_bo_qua', G['csharp_unreachable_bo_qua'], 'khong_thuoc_form', G['csharp_effect_khong_thuoc_form'])
    print('t3_check_trong_form', G['t3_check'], 't3_o_anh_xa', G['t3_cell_anh_xa'], 't3_check_khong_thuoc_form', G['t3_check_khong_thuoc_form'])
    for g in sorted(unmapped):
        tot = sum(unmapped[g].values())
        print('unmapped', g, tot)
        for r, n in unmapped[g].most_common():
            print('   ', r, n)
    print('khong_gian_toa_do_csharp', dict(space_choice))
    print('crosstab StatusID x co_cong_thuc_excel', dict(sorted(crosstab.items(), key=lambda x: str(x[0]))))
    print('sheet_log', len(sheet_log), Counter(x['loi'].split(':')[0] for x in sheet_log))
    if a.sample:
        rnd = random.Random(a.seed)
        flat = sorted(all_cells_flat, key=lambda x: x[0])
        for tag, rec in rnd.sample(flat, a.sample):
            print('SAMPLE', tag, json.dumps(rec, ensure_ascii=False, sort_keys=True)[:900])
    # xung đột ví dụ
    for k in ('excel_khac_csharp', 'nhieu_csharp_khong_dieu_kien', 'khoa_nhung_la_o_nhap', 'khoa_co_dieu_kien', 'validate_o_khong_co_trong_xml'):
        for x in conflicts.get(k, [])[:3]:
            print('VIDU', k, json.dumps(x, ensure_ascii=False, sort_keys=True)[:700])


if __name__ == '__main__':
    main()
