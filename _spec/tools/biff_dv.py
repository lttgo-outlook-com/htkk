"""Giải bản ghi DV (0x01BE) và DVAL (0x01B2) của BIFF8 (.xls) ra JSON.
Dùng: python3 -I biff_dv.py <out_dir> <root1> [<root2> ...]
Đọc [MS-XLS] 2.4.86 (DV), 2.4.87 (DVal), 2.5.198 (Formula/ptg). Không sửa nguồn.
Tệp ra: <out_dir>/<tên-tệp-tương-đối>.json và index.json. Khóa tệp = "<root-basename>/<đường dẫn tương đối>".
"""
import sys, os, struct, json, glob
sys.path.insert(0, glob.glob('/tmp/audit_venv/lib/python*/site-packages')[0])
import olefile

VALTYPE = {0: 'any', 1: 'whole', 2: 'decimal', 3: 'list', 4: 'date', 5: 'time', 6: 'textLength', 7: 'custom'}
ERRSTYLE = {0: 'stop', 1: 'warning', 2: 'information'}
OPER = {0: 'between', 1: 'notBetween', 2: 'equal', 3: 'notEqual', 4: 'greaterThan', 5: 'lessThan',
        6: 'greaterThanOrEqual', 7: 'lessThanOrEqual'}
ERRS = {0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A'}


def col(c):
    s = ''
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def ref(r, c, rel=False):
    return f'{col(c & 0x3FFF)}{r + 1}'


def xstr(b, p, cch_size=2):
    cch = struct.unpack_from('<H', b, p)[0]
    fl = b[p + 2]
    p += 3
    if fl & 1:
        s = b[p:p + 2 * cch].decode('utf-16le', 'replace'); p += 2 * cch
    else:
        s = b[p:p + cch].decode('latin-1'); p += cch
    return s, p


def ptg(b):
    """Giải ptg tối thiểu; trả (chuỗi, danh_sách_cảnh_báo)."""
    out = []; warn = []; p = 0
    while p < len(b):
        t = b[p]; p += 1
        base = t if t < 0x20 else ((t & 0x1F) | 0x20)
        if t == 0x17:
            cch = b[p]; fl = b[p + 1]; p += 2
            if fl & 1:
                s = b[p:p + 2 * cch].decode('utf-16le', 'replace'); p += 2 * cch
            else:
                s = b[p:p + cch].decode('latin-1'); p += cch
            out.append('"' + s.replace('"', '""') + '"')
        elif t == 0x1E:
            out.append(str(struct.unpack_from('<H', b, p)[0])); p += 2
        elif t == 0x1F:
            out.append(repr(struct.unpack_from('<d', b, p)[0])); p += 8
        elif t == 0x1D:
            out.append('TRUE' if b[p] else 'FALSE'); p += 1
        elif t == 0x1C:
            out.append(ERRS.get(b[p], f'#ERR{b[p]}')); p += 1
        elif base == 0x24:  # PtgRef
            r, c = struct.unpack_from('<HH', b, p); p += 4
            out.append(('$' if not c & 0x4000 else '') + col(c & 0x3FFF) + ('$' if not c & 0x8000 else '') + str(r + 1))
        elif base == 0x25:  # PtgArea
            r1, r2, c1, c2 = struct.unpack_from('<HHHH', b, p); p += 8
            out.append(f'{col(c1 & 0x3FFF)}{r1 + 1}:{col(c2 & 0x3FFF)}{r2 + 1}')
        elif base == 0x23:  # PtgName
            out.append(f'NAME#{struct.unpack_from("<I", b, p)[0]}'); p += 4
        elif base == 0x3A:  # PtgRef3d
            ix, r, c = struct.unpack_from('<HHH', b, p); p += 6
            out.append(f'[{ix}]!{col(c & 0x3FFF)}{r + 1}')
        elif base == 0x3B:
            ix, r1, r2, c1, c2 = struct.unpack_from('<HHHHH', b, p); p += 10
            out.append(f'[{ix}]!{col(c1 & 0x3FFF)}{r1 + 1}:{col(c2 & 0x3FFF)}{r2 + 1}')
        elif base == 0x3C:  # PtgRefErr3d / 0x3D area err
            p += 6; out.append('#REF!')
        elif base == 0x3D:
            p += 10; out.append('#REF!')
        elif base == 0x29:  # PtgMemFunc: bỏ qua subexpression cce byte (biểu thức thật nằm sau)
            cce = struct.unpack_from('<H', b, p)[0]; p += 2 + cce
        elif base in (0x26, 0x27, 0x28):  # PtgMemArea/MemErr/MemNoMem: 4 byte unused + cce
            cce = struct.unpack_from('<H', b, p + 4)[0]; p += 6 + (cce if base == 0x26 else 0)
        elif base == 0x2C:  # PtgRefN (tương đối theo ô neo): hiển thị hàng/cột thô
            r, c = struct.unpack_from('<HH', b, p); p += 4
            rr = r - 65536 if (c & 0x8000) and r > 32767 else r
            cc = c & 0xFF; cc = cc - 256 if (c & 0x4000) and cc > 127 else cc
            out.append(f'RefN(dr={rr},dc={cc})')
        elif base == 0x2D:  # PtgAreaN
            r1, r2, c1, c2 = struct.unpack_from('<HHHH', b, p); p += 8
            out.append(f'AreaN(r{r1}:{r2},c{c1 & 0xFF}:{c2 & 0xFF})')
        elif base == 0x39:  # PtgNameX
            out.append(f'NAMEX#{struct.unpack_from("<HI", b, p)}'); p += 6
        elif base == 0x2A:
            p += 4; out.append('#REF!')
        elif base == 0x2B:
            p += 8; out.append('#REF!')
        elif base == 0x21:  # PtgFunc
            out.append(f'FUNC#{struct.unpack_from("<H", b, p)[0]}()'); p += 2
        elif base == 0x22:  # PtgFuncVar
            n = b[p] & 0x7F; f = struct.unpack_from('<H', b, p + 1)[0] & 0x7FFF; p += 3
            out.append(f'FUNC#{f}(n={n})')
        elif t == 0x15:
            out.insert(len(out) - 0, '()') if False else None
        elif t == 0x16:
            out.append('')
        elif t == 0x19:  # PtgAttr
            ty = b[p]; p += 1
            if ty & 0x04:
                n = struct.unpack_from('<H', b, p)[0]; p += 2 + 2 * (n + 1)
            else:
                p += 2
        elif 0x03 <= t <= 0x14:
            ops = {3: '+', 4: '-', 5: '*', 6: '/', 7: '^', 8: '&', 9: '<', 10: '<=', 11: '=', 12: '>=', 13: '>', 14: '<>',
                   15: ' ', 16: ',', 17: ':', 18: 'u+', 19: 'u-', 20: '%'}
            out.append(f'<{ops[t]}>')
        elif t == 0x01:  # PtgExp
            p += 4; out.append('<EXP>')
        else:
            warn.append(f'ptg 0x{t:02X} tại {p - 1}')
            break
    return ' '.join(out), warn


def records(d):
    pos = 0
    while pos + 4 <= len(d):
        t, l = struct.unpack_from('<HH', d, pos)
        yield t, d[pos + 4:pos + 4 + l], pos
        pos += 4 + l


def parse_dv(b):
    fl = struct.unpack_from('<I', b, 0)[0]
    p = 4
    strs = []
    for _ in range(4):
        s, p = xstr(b, p)
        strs.append('' if s == '\x00' else s)  # Excel lưu chuỗi rỗng là 1 ký tự NUL
    cce1 = struct.unpack_from('<H', b, p)[0]; p += 4
    f1 = b[p:p + cce1]; p += cce1
    cce2 = struct.unpack_from('<H', b, p)[0]; p += 4
    f2 = b[p:p + cce2]; p += cce2
    cref = struct.unpack_from('<H', b, p)[0]; p += 2
    rngs = []
    for _ in range(cref):
        r1, r2, c1, c2 = struct.unpack_from('<HHHH', b, p); p += 8
        a = f'{col(c1)}{r1 + 1}' + ('' if (r1, c1) == (r2, c2) else f':{col(c2)}{r2 + 1}')
        rngs.append(dict(ref=a, row_first=r1, row_last=r2, col_first=c1, col_last=c2))
    vt = fl & 0xF
    s1, w1 = ptg(f1); s2, w2 = ptg(f2)
    d = dict(
        type=VALTYPE.get(vt, vt), error_style=ERRSTYLE.get((fl >> 4) & 7, (fl >> 4) & 7),
        str_lookup=bool(fl & 0x80), allow_blank=bool(fl & 0x100), suppress_dropdown=bool(fl & 0x200),
        show_input_msg=bool(fl & 0x40000), show_error_msg=bool(fl & 0x80000),
        operator=OPER.get((fl >> 20) & 0xF, (fl >> 20) & 0xF), flags_hex=f'0x{fl:08X}',
        prompt_title=strs[0], error_title=strs[1], prompt=strs[2], error=strs[3],
        formula1_hex=f1.hex(), formula1=s1, formula2_hex=f2.hex(), formula2=s2, ranges=rngs,
        parse_trailing_bytes=len(b) - p)
    if vt == 3 and (fl & 0x80) and f1[:1] == b'\x17':
        d['list_values'] = [x for x in s1[1:-1].replace('""', '"').split('\x00')] if s1.startswith('"') else None
        d['list_values_note'] = 'fStrLookup=1: danh sách chuỗi trong ptgStr, phân tách bằng NUL'
    if w1 or w2:
        d['warnings'] = w1 + w2
    return d


def process(path):
    ole = olefile.OleFileIO(path)
    try:
        s = 'Workbook' if ole.exists('Workbook') else 'Book'
        d = ole.openstream(s).read()
    finally:
        ole.close()
    sheet = -1; sheets = []; cur = None; dvs = []; dvals = []; raw = 0
    depth_sheet_names = []
    for t, b, off in records(d):
        if t == 0x0085:
            cch = b[6]; hi = b[7] & 1
            depth_sheet_names.append(b[8:8 + cch * (2 if hi else 1)].decode('utf-16le' if hi else 'latin-1'))
        elif t == 0x0809:  # BOF
            vt = struct.unpack_from('<H', b, 2)[0] if len(b) >= 4 else None
            if vt == 0x0010:
                sheet += 1
        elif t == 0x01B2:
            dvals.append(dict(sheet_index=sheet, offset=off, raw_hex=b.hex(), dv_count=struct.unpack_from('<I', b, 14)[0]))
        elif t == 0x01BE:
            raw += 1
            try:
                x = parse_dv(b)
            except Exception as e:
                x = dict(parse_error=str(e), raw_hex=b.hex())
            x['sheet_index'] = sheet
            x['sheet'] = depth_sheet_names[sheet] if 0 <= sheet < len(depth_sheet_names) else None
            x['offset'] = off
            dvs.append(x)
    return dict(sheets=depth_sheet_names, dval=dvals, dv=dvs, dv_count=len(dvs), dval_count=len(dvals))


def main():
    out = sys.argv[1]; roots = sys.argv[2:]
    os.makedirs(out, exist_ok=True)
    index = []; tot_dv = tot_dval = 0; fail = []; nfile = 0
    for R in roots:
        base = os.path.basename(R.rstrip('/'))
        for dp, dn, fn in os.walk(R):
            for f in sorted(fn):
                if not (f.endswith('.xls') or f.endswith('.xls1')):
                    continue
                p = os.path.join(dp, f); nfile += 1
                key = base + '/' + os.path.relpath(p, R)
                try:
                    r = process(p)
                except Exception as e:
                    fail.append((key, str(e))); continue
                if r['dv_count'] == 0 and r['dval_count'] == 0:
                    continue
                r['source'] = key
                name = key.replace('/', '__') + '.json'
                with open(os.path.join(out, name), 'w', encoding='utf8') as fh:
                    json.dump(r, fh, ensure_ascii=False, indent=1)
                tot_dv += r['dv_count']; tot_dval += r['dval_count']
                index.append(dict(source=key, json=name, dv_count=r['dv_count'], dval_count=r['dval_count'],
                                  types=sorted({x.get('type', '?') for x in r['dv']})))
    with open(os.path.join(out, 'index.json'), 'w', encoding='utf8') as fh:
        json.dump(dict(roots=roots, files_scanned=nfile, files_with_dv=len(index), total_dv=tot_dv,
                       total_dval=tot_dval, read_failures=fail, files=index), fh, ensure_ascii=False, indent=1)
    print(dict(files_scanned=nfile, files_with_dv=len(index), total_dv=tot_dv, total_dval=tot_dval, failures=len(fail)))


if __name__ == '__main__':
    main()
