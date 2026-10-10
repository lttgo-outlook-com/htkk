#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render_cong_thuc.py: sinh lại mọi _forms/**/cong_thuc.txt từ JSON đặc tả. cong_thuc.txt chỉ là view, không sửa tay.

Chạy (nên dùng venv có olefile để lấy nhãn ô; không có olefile thì cột nhãn để trống):
  /tmp/audit_venv/bin/python -I _spec/tools/render_cong_thuc.py \
      [--repo .] [--rules <dir T2>] [--no-labels] [--dry-run]

Đầu vào (đều tương đối theo --repo):
  _forms/_manifest.csv                      tệp nguồn -> đường dẫn trong _forms (nối theo source_name + sha256)
  _spec/formulas/index.json + <json>        công thức từng tệp (T1)
  _spec/csharp/classes.json                 MaFile + VersionXML -> Class, MenuID (T0)
  _spec/config/mapcongthucchosua.xml.json   công thức ROWREPLACE theo Map ID = MenuID (T4)
  --rules <dir>  (tùy chọn)                 công thức do C# đặt (T2), xem load_rules()
  _forms/<thư mục>/<tệp bảng tính>          chỉ để đọc ô văn bản làm nhãn (BIFF, olefile)

Định dạng một cong_thuc.txt (mỗi tệp bảng tính một khối, khối theo thứ tự đường dẫn: .xls rồi .xls1):
  MaFile: ...            <- dòng đầu của tệp
  VersionXML: ...
  Class: ...
  MenuID: ...
  Nguồn sha256: ...
  Generator: render_cong_thuc.py@<version>
  === FILE: <tên tệp> ===
  ## sheet: <tên sheet>
  REF | công thức | nhãn | nguồn          (ô theo hàng rồi cột; nguồn = xls|xls1|xlsx, C#:Class.Method:Lnn, ROWREPLACE)
  C# có điều kiện: REF | công thức | nhãn | C#:... | khi: <guard>
  => N cell công thức trong file này
  ## ROWREPLACE ...  (nếu MenuID có trong mapcongthucchosua)
  ## SUPBOOK (tham chiếu ngoài)  [ext:n] đường dẫn | sheet
Escape trong công thức và nhãn: CR -> \\r, LF -> \\n, dấu \\ -> \\\\, ký tự | trong nhãn -> ¦.
Tham chiếu ngoài trong công thức đã ở dạng [ext:n] (T1); bảng SUPBOOK in cuối khối.
"""
import os, sys, re, csv, json, struct, argparse, collections

VERSION = '1.0.0'
GEN = 'render_cong_thuc.py@' + VERSION
LABEL_MAX = 60


# ------------------------------------------------------------------ tiện ích
def esc(s):
    return s.replace('\\', '\\\\').replace('\r', '\\r').replace('\n', '\\n')


def esc_label(s):
    s = re.sub(r'\s+', ' ', s).strip().replace('|', '¦')
    if len(s) > LABEL_MAX:
        s = s[:LABEL_MAX - 1] + '…'
    return esc(s)


def clean_path(p):
    return re.sub(r'[\x00-\x1f]+', '/', p or '').strip('/') if p and p.strip() else (p or '')


def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


REF_RE = re.compile(r'^([A-Z]+)(\d+)$')


def split_ref(ref):
    m = REF_RE.match(ref)
    return (int(m.group(2)), col2n(m.group(1))) if m else (10 ** 9, 0)


# ------------------------------------------------------------------ nhãn: ô văn bản từ BIFF8
def _records(buf):
    pos, n = 0, len(buf)
    while pos + 4 <= n:
        rid, ln = struct.unpack_from('<HH', buf, pos)
        yield pos, rid, buf[pos + 4:pos + 4 + ln]
        pos += 4 + ln


def _read_sst(chunks):
    """chunks: [SST data, CONTINUE data...]; trả danh sách chuỗi."""
    ci, off = 0, 8
    cur = chunks[0]
    uniq = struct.unpack_from('<I', cur, 4)[0]
    out = []

    def nxt():
        nonlocal ci, off, cur
        ci += 1
        if ci >= len(chunks):
            raise ValueError('hết CONTINUE')
        cur = chunks[ci]
        off = 0

    for _ in range(uniq):
        if off >= len(cur):
            nxt()
        cch = struct.unpack_from('<H', cur, off)[0]
        fl = cur[off + 2]
        off += 3
        crun = cb = 0
        if fl & 8:
            crun = struct.unpack_from('<H', cur, off)[0]
            off += 2
        if fl & 4:
            cb = struct.unpack_from('<I', cur, off)[0]
            off += 4
        wide = fl & 1
        rem = cch
        parts = []
        while rem > 0:
            if off >= len(cur):
                nxt()
                wide = cur[0] & 1
                off = 1
            avail = (len(cur) - off) // (2 if wide else 1)
            take = min(avail, rem)
            raw = cur[off:off + take * (2 if wide else 1)]
            parts.append(raw.decode('utf-16-le' if wide else 'latin-1', 'replace'))
            off += len(raw)
            rem -= take
        skip = crun * 4 + cb
        while skip > 0:
            if off >= len(cur):
                nxt()
            t = min(skip, len(cur) - off)
            off += t
            skip -= t
        out.append(''.join(parts))
    return out


def biff_labels(path, olefile):
    """{tên sheet: {hàng(1-based): [(cột(1-based), text)]}} từ ô hằng văn bản; ném lỗi nếu không đọc được."""
    ole = olefile.OleFileIO(path)
    try:
        name = 'Workbook' if ole.exists('Workbook') else 'Book'
        buf = ole.openstream(name).read()
    finally:
        ole.close()
    sheets = {}
    sst = []
    pending = None
    recs = list(_records(buf))
    # BOUNDSHEET
    for pos, rid, d in recs:
        if rid == 0x85:
            off, _st, typ = struct.unpack_from('<IBB', d, 0)
            cch, fl = d[6], d[7]
            nm = d[8:8 + cch * (2 if fl & 1 else 1)].decode('utf-16-le' if fl & 1 else 'latin-1', 'replace')
            sheets[off] = (nm, typ)
    # SST + CONTINUE
    i = 0
    while i < len(recs):
        if recs[i][1] == 0xFC:
            chunks = [recs[i][2]]
            j = i + 1
            while j < len(recs) and recs[j][1] == 0x3C:
                chunks.append(recs[j][2])
                j += 1
            sst = _read_sst(chunks)
            break
        i += 1
    out = {}
    cur = None
    for pos, rid, d in recs:
        if rid == 0x809:
            sh = sheets.get(pos)
            cur = out.setdefault(sh[0], {}) if sh and sh[1] == 0 else None
        elif cur is not None and rid == 0xFD and len(d) >= 10:
            r, c, _x, isst = struct.unpack_from('<HHHI', d, 0)
            if isst < len(sst) and sst[isst].strip():
                cur.setdefault(r + 1, []).append((c + 1, sst[isst]))
        elif cur is not None and rid == 0x204 and len(d) >= 8:
            r, c, _x, cch = struct.unpack_from('<HHHH', d, 0)
            fl = d[8] if len(d) > 8 else 0
            s = d[9:9 + cch * (2 if fl & 1 else 1)].decode('utf-16-le' if fl & 1 else 'latin-1', 'replace')
            if s.strip():
                cur.setdefault(r + 1, []).append((c + 1, s))
    return out


def nearest_label(rowcells, col):
    best = None
    for c, t in rowcells:
        if c == col:
            continue
        key = (abs(c - col), 0 if c < col else 1)
        if best is None or key < best[0]:
            best = (key, t)
    return best[1] if best else ''


# ------------------------------------------------------------------ đầu vào T2 (tùy chọn)
def load_rules(rules_dir, key):
    """Công thức do C# đặt (T2). Chưa có định dạng chốt; bộ render đọc <rules_dir>/<key>.json
    (key = khóa T1 của tệp, ví dụ 01_GTGT@v2.6.2@_1_gtgt_262.xls) dạng
      {"rules": [{"sheet","cell","formula","class","method","line","guard"(tùy chọn)}]}
    Không có tệp hoặc không đúng dạng thì bỏ qua (trả danh sách rỗng)."""
    if not rules_dir:
        return []
    p = os.path.join(rules_dir, key + '.json')
    if not os.path.isfile(p):
        return []
    try:
        j = json.load(open(p, encoding='utf-8'))
        return [r for r in j.get('rules', []) if isinstance(r, dict) and r.get('sheet') is not None and r.get('cell') and 'formula' in r]
    except (ValueError, OSError, AttributeError):
        return []


# ------------------------------------------------------------------ dựng khối từng tệp
def render_block(repo, row, ent, jdoc, classes, menu_by_mafile, rowrep, rules_dir, olefile, stats):
    # nguồn của ô gốc: xls | xls1 | xlsx (tên dst có thể bị cắt, ví dụ .xl, .xl1, nên suy theo định dạng và hậu tố 1)
    dext = row['dst'].rsplit('.', 1)[-1]
    ext = 'xlsx' if jdoc.get('format') == 'xlsx' else ('xls1' if dext.endswith('1') else 'xls')
    base = os.path.basename(row['dst'])
    cl = classes.get((ent.get('mafile'), ent.get('version_xml')))
    out = []
    out.append('MaFile: %s' % (ent.get('mafile') or '-'))
    out.append('VersionXML: %s' % (ent.get('version_xml') or '-'))
    out.append('Class: %s' % (cl['class'] if cl else '-'))
    mid_file = cl['MenuID'] if cl else menu_by_mafile.get(ent.get('mafile'))
    out.append('MenuID: %s' % (mid_file or '-'))
    out.append('Nguồn sha256: %s' % ent['source_sha256'])
    out.append('Generator: %s' % GEN)
    out.append('=== FILE: %s ===' % base)

    labels = {}
    if olefile is not None and jdoc.get('format') == 'biff8':
        try:
            labels = biff_labels(os.path.join(repo, '_forms', row['dst']), olefile)
            stats['label_ok'] += 1
        except Exception:
            stats['label_fail'] += 1

    sheet_order = {s['name']: s['index'] for s in jdoc['sheets']}
    rules = load_rules(rules_dir, jdoc['key'])
    by_sheet = collections.defaultdict(list)
    for x in jdoc['formulas']:
        by_sheet[x['sheet']].append(('base', x))
    for r in rules:
        by_sheet[r['sheet']].append(('rule', r))
    ncell = 0
    for sh in sorted(by_sheet, key=lambda n: (sheet_order.get(n, 10 ** 6), n)):
        out.append('## sheet: %s' % esc(sh))
        items = sorted(by_sheet[sh], key=lambda t: (split_ref(t[1]['cell']), 0 if t[0] == 'base' else 1))
        lab = labels.get(sh, {})
        for kind, x in items:
            rc = split_ref(x['cell'])
            label = esc_label(nearest_label(lab.get(rc[0], ()), rc[1])) if lab else ''
            if kind == 'base':
                ncell += 1
                line = '%s | %s | %s | %s' % (x['cell'], esc(x['formula']), label, ext)
                if 'array_ref' in x:
                    line += ' | mảng: %s' % x['array_ref']
                if 'lost_refs' in x:
                    line += ' | mất tham chiếu: %s' % esc(json.dumps(x['lost_refs'], ensure_ascii=False))
            else:
                src = 'C#:%s.%s:L%s' % (x.get('class', '?'), x.get('method', '?'), x.get('line', '?'))
                line = '%s | %s | %s | %s' % (x['cell'], esc(str(x['formula'])), label, src)
                if x.get('guard'):
                    line += ' | khi: %s' % esc(str(x['guard']))
            out.append(line)
    out.append('=> %d cell công thức trong file này' % ncell)
    stats['cells'] += ncell
    stats['files'] += 1

    mid = cl['MenuID'] if cl else None   # ROWREPLACE chỉ gắn với đúng phiên bản có trong menu
    if mid and mid in rowrep:
        out.append('## ROWREPLACE (mapcongthucchosua.xml, Map ID=%s)' % mid)
        for it in rowrep[mid]:
            a = it['attrs']
            lab = 'SheetIndex=%s; Dynamic=%s' % (a.get('SheetIndex', ''), a.get('Dynamic', ''))
            if a.get('TablePath'):
                lab += '; TablePath=%s' % a['TablePath']
            out.append('%s | %s | %s | ROWREPLACE' % (a.get('MCTPath', ''), esc(a.get('Formula', '')), esc(lab)))
            stats['rowreplace'] += 1
    ext_sb = [s for s in jdoc['supbooks'] if s['kind'] != 'internal']
    if ext_sb:
        out.append('## SUPBOOK (tham chiếu ngoài)')
        for s in ext_sb:
            out.append('[ext:%d] %s | sheet: %s' % (s['index'], esc(clean_path(s['path'])), esc(','.join(s['sheets']))))
    return out


# ------------------------------------------------------------------ chính
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', default=os.getcwd())
    ap.add_argument('--rules', default=None, help='thư mục T2 (_spec/rules), tùy chọn')
    ap.add_argument('--no-labels', action='store_true')
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    repo = os.path.abspath(a.repo)
    F = os.path.join(repo, '_forms')
    FS = os.path.join(repo, '_spec', 'formulas')

    olefile = None
    if not a.no_labels:
        try:
            import olefile
        except ImportError:
            print('CẢNH BÁO: không có olefile, cột nhãn để trống (chạy bằng /tmp/audit_venv/bin/python -I)', file=sys.stderr)

    idx = json.load(open(os.path.join(FS, 'index.json'), encoding='utf-8'))['files']
    ent_by = {(e['source_name'], e['source_sha256']): e for e in idx if e['status'] == 'ok' and e['formulas']}
    manifest = list(csv.DictReader(open(os.path.join(F, '_manifest.csv'), encoding='utf-8')))
    classes = {}
    for c in json.load(open(os.path.join(repo, '_spec', 'csharp', 'classes.json'), encoding='utf-8'))['classes']:
        if not c.get('in_menu') or not c.get('MaFile'):
            continue
        k = (c['MaFile'], c['VersionXML'])
        if k not in classes or (c.get('in_menu') and not classes[k].get('in_menu')):
            classes[k] = c
    menu_by_mafile = {c['MaFile']: c['MenuID'] for c in classes.values() if c.get('in_menu')}
    rowrep = {}
    for m in json.load(open(os.path.join(repo, '_spec', 'config', 'mapcongthucchosua.xml.json'), encoding='utf-8'))['root']['children']:
        rowrep[m['attrs']['ID']] = m['children']

    per_dir = collections.defaultdict(list)
    for r in manifest:
        e = ent_by.get((r['src'], r['sha256']))
        if e:
            per_dir[os.path.dirname(r['dst'])].append((r, e))

    stats = {'cells': 0, 'rowreplace': 0, 'label_ok': 0, 'label_fail': 0, 'files': 0}
    created = rewritten = deleted = 0
    for d in sorted(per_dir):
        lines = []
        for r, e in sorted(per_dir[d], key=lambda t: t[0]['dst']):
            jdoc = json.load(open(os.path.join(FS, e['json']), encoding='utf-8'))
            if lines:
                lines.append('')
            lines += render_block(repo, r, e, jdoc, classes, menu_by_mafile, rowrep, a.rules, olefile, stats)
        text = '\n'.join(lines) + '\n'
        p = os.path.join(F, d, 'cong_thuc.txt')
        existed = os.path.exists(p)
        if not a.dry_run:
            with open(p, 'w', encoding='utf-8', newline='\n') as fh:
                fh.write(text)
        if existed:
            rewritten += 1
        else:
            created += 1
    for dp, dn, fn in os.walk(F):
        if 'cong_thuc.txt' in fn and os.path.relpath(dp, F) not in per_dir:
            print('xóa: %s' % os.path.relpath(os.path.join(dp, 'cong_thuc.txt'), repo))
            if not a.dry_run:
                os.remove(os.path.join(dp, 'cong_thuc.txt'))
            deleted += 1
    print('cong_thuc.txt sinh ra: %d (ghi đè %d, tạo mới %d); xóa: %d' % (rewritten + created, rewritten, created, deleted))
    print('tệp nguồn có công thức: %d; dòng ô: %d; dòng ROWREPLACE: %d; nhãn đọc được ở %d tệp, lỗi %d tệp'
          % (stats['files'], stats['cells'], stats['rowreplace'], stats['label_ok'], stats['label_fail']))


if __name__ == '__main__':
    main()
