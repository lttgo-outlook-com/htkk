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
  --rules <dir>  (tùy chọn)                 công thức do C# đặt (T2), mặc định _spec/rules
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
  #   phiên bản ô <ref>#L: <after>   (1.1.0: sau dòng C#, mỗi nhãn cell(...)#L có trong ô đích, công thức,
                                     điều kiện hoặc sheet đích, lấy từ cellVersions của bản ghi T2)
  Đầu vào T2: _spec/rules/<Class>.json sha256_16=...   (1.1.0: dòng header khi khối có class T2)
  => N cell công thức trong file này
  ## ROWREPLACE ...  (nếu MenuID có trong mapcongthucchosua)
  ## SUPBOOK (tham chiếu ngoài)  [ext:n] đường dẫn | sheet
Escape trong công thức và nhãn: CR -> \\r, LF -> \\n, dấu \\ -> \\\\, ký tự | trong nhãn -> ¦.
Tham chiếu ngoài trong công thức đã ở dạng [ext:n] (T1); bảng SUPBOOK in cuối khối.
"""
import os, sys, re, csv, json, struct, argparse, collections, hashlib

VERSION = '1.1.0'
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


# ------------------------------------------------------------------ đầu vào T2 (_spec/rules)
RULES_SKIP = {'index.json', '_app.json', '_shared.json', 'opaque_calls.json', 'formula_points.json'}


def sheet_ref(t):
    """Khóa sheet của target: int (chỉ số sheet) hoặc str (sheet tham số); target rỗng -> None."""
    return t.get('sheet') if isinstance(t, dict) else None


def when_text(rec):
    """Phần điều kiện kích hoạt của bản ghi (trigger) thành văn bản; rỗng nếu không có."""
    tr = rec.get('trigger') or {}
    parts = []
    if tr.get('cells'):
        parts.append('sửa ô ' + ','.join(tr['cells']))
    rr = tr.get('rowRange')
    if rr:
        parts.append('sửa dòng %s..%s cột %s' % (rr.get('from'), rr.get('to'), ','.join(rr.get('cols') or []) or '*'))
    if not parts and tr.get('when'):
        parts.append(' && '.join(tr['when']))
    return ' && '.join(parts)


def load_rules(rules_dir):
    """Đọc <rules_dir>/<Class>.json (T2) theo định dạng thật: {key,file,hooks,records[]}.
    Trả {Class: {'lines': [...], 'by_sheet_other': {sheet: {'setValue':n,'lock':n}}, 'unreach': {sheet: [records, setFormula]},
                  'n_effects': số effect setFormula không unreachable, 'n_lines': số dòng (mỗi nhánh cases một dòng)}}.
    Chỉ lấy op setFormula không unreachable thành dòng; setValue và lock chỉ đếm theo sheet."""
    out = {}
    if not rules_dir:
        return out
    for fn in sorted(os.listdir(rules_dir)):
        if not fn.endswith('.json') or fn in RULES_SKIP:
            continue
        raw = open(os.path.join(rules_dir, fn), 'rb').read()
        j = json.loads(raw.decode('utf-8'))
        if not isinstance(j, dict) or not isinstance(j.get('records'), list):
            continue
        cls = j['key']['Class']
        d = {'lines': [], 'other': collections.defaultdict(lambda: collections.Counter()),
             'unreach': collections.defaultdict(lambda: [0, 0]), 'n_effects': 0, 'n_lines': 0}
        for ri, rec in enumerate(j['records']):
            rsheet = rec.get('sheetIndex')
            if rec.get('unreachable'):
                u = d['unreach'][rsheet]
                u[0] += 1
                u[1] += sum(1 for e in rec['effects'] if e['op'] == 'setFormula')
                continue
            khi_pre = []
            w = when_text(rec)
            if w:
                khi_pre.append('kích hoạt: ' + w)
            khi_pre += [str(g) for g in rec.get('guards') or []]
            cvmap = {}
            for v in rec.get('cellVersions') or []:
                if isinstance(v, dict) and v.get('ref'):
                    cvmap.setdefault(v["ref"], v.get("after", ""))
            for ei, e in enumerate(rec['effects']):
                sh = sheet_ref(e.get('target'))
                if sh is None:
                    sh = rsheet
                if e['op'] in ('setValue', 'lock'):
                    d['other'][sh][e['op']] += 1
                    continue
                if e['op'] != 'setFormula':
                    continue
                d['n_effects'] += 1
                f = e.get('formula')
                if isinstance(f, dict) and isinstance(f.get('cases'), list):
                    cases = [(c.get('when'), c.get('value')) for c in f['cases']]
                else:
                    cases = [(None, f)]
                reset = bool((e.get('params') or {}).get('reset'))
                for ci, (cw, cv) in enumerate(cases):
                    khi = list(khi_pre)
                    if cw is not None:
                        khi.append('ngược lại' if cw == 'else' else str(cw))
                    if cv is None or cv == '':
                        ftxt = '<xóa>'
                    elif isinstance(cv, str):
                        ftxt = cv
                    else:
                        ftxt = json.dumps(cv, ensure_ascii=False, sort_keys=True)
                    cell = (e.get('target') or {}).get('cell', '?')
                    khi_txt = ' && '.join(khi)
                    blob = '%s\n%s\n%s\n%s' % (cell, ftxt, khi_txt, sh if isinstance(sh, str) else '')
                    vers = [(r_, cvmap[r_]) for r_ in sorted(cvmap) if r_ in blob and re.search(re.escape(r_) + r"(?![0-9]|\.[0-9])", blob)]
                    d['lines'].append({'sheet': sh, 'cell': cell, 'formula': ftxt,
                                       'method': rec.get('hook', '?'), 'line': e.get('at', '?'), 'khi': khi_txt,
                                       'ord': (cls, ri, ei, ci), 'versions': vers})
                    d['n_lines'] += 1
        d['cls'] = cls
        d['file'] = fn
        d['sha256_16'] = hashlib.sha256(raw).hexdigest()[:16]
        out[cls] = d
    return out


# ------------------------------------------------------------------ dựng khối từng tệp
def render_block(repo, row, ent, jdoc, classes, cls_of, menu_by_mafile, rowrep, crules, olefile, stats):
    # nguồn của ô gốc: xls | xls1 | xlsx (tên dst có thể bị cắt, ví dụ .xl, .xl1, nên suy theo định dạng và hậu tố 1)
    dext = row['dst'].rsplit('.', 1)[-1]
    ext = 'xlsx' if jdoc.get('format') == 'xlsx' else ('xls1' if dext.endswith('1') else 'xls')
    base = os.path.basename(row['dst'])
    cl = classes.get((ent.get('mafile'), ent.get('version_xml')))
    out = []
    out.append('MaFile: %s' % (ent.get('mafile') or '-'))
    out.append('VersionXML: %s' % (ent.get('version_xml') or '-'))
    out.append('Class: %s' % (cls_of.get((ent.get('mafile'), ent.get('version_xml'))) or '-'))
    mid_file = cl['MenuID'] if cl else menu_by_mafile.get(ent.get('mafile'))
    out.append('MenuID: %s' % (mid_file or '-'))
    out.append('Nguồn sha256: %s' % ent['source_sha256'])
    out.append('Generator: %s' % GEN)
    if crules:
        out.append('Đầu vào T2: _spec/rules/%s sha256_16=%s' % (crules['file'], crules['sha256_16']))
    out.append('=== FILE: %s ===' % base)

    labels = {}
    if olefile is not None and jdoc.get('format') == 'biff8':
        try:
            labels = biff_labels(os.path.join(repo, '_forms', row['dst']), olefile)
            stats['label_ok'] += 1
        except Exception:
            stats['label_fail'] += 1

    sheet_order = {s['name']: s['index'] for s in jdoc['sheets']}
    by_sheet = collections.defaultdict(list)
    for x in jdoc['formulas']:
        by_sheet[x['sheet']].append(('base', x))
    idx_name = {s['index']: s['name'] for s in jdoc['sheets']}

    def sheet_key(sh):
        if isinstance(sh, int):
            return idx_name.get(sh, 'S%d (không có trong tệp)' % sh)
        return sh if sh in sheet_order else '%s (sheet tham số)' % sh

    other = collections.defaultdict(collections.Counter)
    unreach = collections.defaultdict(lambda: [0, 0])
    ncs = 0
    if crules:
        for x in crules['lines']:
            by_sheet[sheet_key(x['sheet'])].append(('rule', x))
        for sh, c in crules['other'].items():
            other[sheet_key(sh)].update(c)
        for sh, u in crules['unreach'].items():
            k = sheet_key(sh)
            unreach[k][0] += u[0]
            unreach[k][1] += u[1]
        for k in list(other) + list(unreach):
            by_sheet.setdefault(k, [])
    ncell = 0
    for sh in sorted(by_sheet, key=lambda n: (sheet_order.get(n, 10 ** 6), n)):
        out.append('## sheet: %s' % esc(sh))
        items = sorted(by_sheet[sh], key=lambda t: ((split_ref(t[1]['cell']), 0, t[1]['cell'], ()) if t[0] == 'base'
                                                    else (split_ref(t[1]['cell']), 1, t[1]['cell'], t[1]['ord'])))
        lab = labels.get(sh, {})
        nsh = 0
        for kind, x in items:
            rc = split_ref(x['cell'])
            label = esc_label(nearest_label(lab.get(rc[0], ()), rc[1])) if lab and rc[0] < 10 ** 9 else ''
            if kind == 'base':
                ncell += 1
                line = '%s | %s | %s | %s' % (x['cell'], esc(x['formula']), label, ext)
                if 'array_ref' in x:
                    line += ' | mảng: %s' % x['array_ref']
                if 'lost_refs' in x:
                    line += ' | mất tham chiếu: %s' % esc(json.dumps(x['lost_refs'], ensure_ascii=False))
            else:
                nsh += 1
                src = 'C#:%s.%s:L%s' % (x['ord'][0], x['method'], x['line'])
                line = '%s | %s | %s | %s' % (x['cell'], esc(x['formula']), label, src)
                if x['khi']:
                    line += ' | khi: %s' % esc(x['khi'])
            out.append(line)
            for r_, af in (x.get('versions') or ()) if kind != 'base' else ():
                out.append('#   phiên bản ô %s: %s' % (esc(r_), esc(af)))
        if nsh or other.get(sh) or unreach.get(sh):
            o, u = other.get(sh, {}), unreach.get(sh, [0, 0])
            out.append('# C# tóm tắt sheet: %d dòng setFormula (liệt kê ở trên); setValue %d, lock %d (không liệt kê); '
                       'bản ghi unreachable bỏ %d (%d setFormula)' % (nsh, o.get('setValue', 0), o.get('lock', 0), u[0], u[1]))
        ncs += nsh
    out.append('=> %d cell công thức trong file này' % ncell)
    if crules:
        out.append('=> %d dòng công thức C# trong file này (%d effect setFormula, class %s)'
                   % (ncs, crules['n_effects'], crules['cls']))
    stats['cells'] += ncell
    stats['cs_lines'] += ncs
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
    ap.add_argument('--rules', default='_spec/rules', help='thư mục T2 (mặc định _spec/rules tương đối theo --repo); rỗng = bỏ C#')
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
    # class không có trong menu (phiên bản cũ của MenuClass, tên <MenuClass>Nckt<số phiên bản>): suy ra (MaFile, VersionXML)
    # qua MenuClass của bản in_menu cùng tên gốc và VersionXML bỏ dấu chấm trùng số hậu tố
    cls_of = {k: c['class'] for k, c in classes.items()}
    mafile_of_menuclass = {c['MenuClass']: c['MaFile'] for c in classes.values()}
    ver_digits = {}
    for e in ent_by.values():
        if e.get('mafile') and e.get('version_xml'):
            ver_digits[(e['mafile'], e['version_xml'].replace('.', ''))] = e['version_xml']
    for c in json.load(open(os.path.join(repo, '_spec', 'csharp', 'classes.json'), encoding='utf-8'))['classes']:
        if c.get('in_menu'):
            continue
        m = re.match(r'^(.*)Nckt(\d+)$', c['class'])
        mf = mafile_of_menuclass.get(m.group(1)) if m else None
        vx = ver_digits.get((mf, m.group(2))) if mf else None
        if vx and (mf, vx) not in cls_of:
            cls_of[(mf, vx)] = c['class']
    menu_by_mafile = {c['MaFile']: c['MenuID'] for c in classes.values() if c.get('in_menu')}
    rowrep = {}
    for m in json.load(open(os.path.join(repo, '_spec', 'config', 'mapcongthucchosua.xml.json'), encoding='utf-8'))['root']['children']:
        rowrep[m['attrs']['ID']] = m['children']

    per_dir = collections.defaultdict(list)
    for r in manifest:
        e = ent_by.get((r['src'], r['sha256']))
        if e:
            per_dir[os.path.dirname(r['dst'])].append((r, e))

    rules_dir = os.path.join(repo, a.rules) if a.rules else None
    crules_all = load_rules(rules_dir)
    # mỗi class gắn đúng một lần: vào tệp đầu tiên (theo đường dẫn dst) của form (MaFile, VersionXML) có class đó
    attach = {}
    for r in sorted(manifest, key=lambda r: r['dst']):
        e = ent_by.get((r['src'], r['sha256']))
        cn = cls_of.get((e.get('mafile'), e.get('version_xml'))) if e else None
        if cn and cn in crules_all and cn not in attach.values():
            attach[(r['src'], r['sha256'])] = cn
    stats = {'cells': 0, 'rowreplace': 0, 'label_ok': 0, 'label_fail': 0, 'files': 0, 'cs_lines': 0}
    created = rewritten = deleted = 0
    for d in sorted(per_dir):
        lines = []
        for r, e in sorted(per_dir[d], key=lambda t: t[0]['dst']):
            jdoc = json.load(open(os.path.join(FS, e['json']), encoding='utf-8'))
            if lines:
                lines.append('')
            lines += render_block(repo, r, e, jdoc, classes, cls_of, menu_by_mafile, rowrep,
                                 crules_all.get(attach.get((r['src'], r['sha256']))), olefile, stats)
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
    if rules_dir:
        mapped = sorted(set(attach.values()))
        in_menu_cls = set(cls_of.values())
        unm = {'không suy ra được (MaFile, VersionXML) có tệp bảng tính (không in_menu, hoặc MenuClass/phiên bản không khớp tệp)': [],
               'form có MaFile + VersionXML nhưng không có tệp bảng tính có công thức Excel (không có cong_thuc.txt)': []}
        for c in sorted(set(crules_all) - set(mapped)):
            unm[list(unm)[0] if c not in in_menu_cls else list(unm)[1]].append(c)
        print('class T2: %d; map được vào form: %d (effect setFormula %d, dòng C# %d); không map được: %d'
              % (len(crules_all), len(mapped), sum(crules_all[c]['n_effects'] for c in mapped),
                 stats['cs_lines'], len(unm[list(unm)[0]]) + len(unm[list(unm)[1]])))
        for k, v in unm.items():
            print('  không map được, lý do "%s": %d class, effect setFormula %d' % (k, len(v), sum(crules_all[c]['n_effects'] for c in v)))


if __name__ == '__main__':
    main()
