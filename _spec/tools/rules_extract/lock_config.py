#!/usr/bin/env python3
"""Đọc bảng cấu hình khóa/mở ô của GeneralFunctions.checkLockUnLockCells trong sheet cuối của mẫu .xls.
Cột AK (dòng 2 trở đi): công thức tham chiếu ô điều khiển (sheet, dòng mẫu); AL: công thức điều kiện (giá trị 0 = khóa);
AM: danh sách ô đích nối bằng &; AN: "1" = xóa dữ liệu khi khóa; BA1: "1" = bật.
Chỉ parse byte BIFF8 (luồng Workbook qua cfb.py); không chạy macro hay mã nào trong tệp.
Dùng: python3 -I lock_config.py <decompiledDir> <classes.json> <manifest.csv> <formsDir> <cfb.py dir> <out.json>"""
import csv, glob, json, os, re, struct, sys

dec, classes_path, manifest_path, forms, cfb_dir, out_path = sys.argv[1:7]
sys.path.insert(0, cfb_dir)
import cfb  # noqa: E402  bộ đọc OLE2 có sẵn trong repo (chỉ đọc byte)

AK, AL, AM, AN, BA, BB = 36, 37, 38, 39, 52, 53
FUNCS = {0: ("COUNT", None), 1: ("IF", None), 2: ("ISNA", 1), 3: ("ISERROR", 1), 4: ("SUM", None), 5: ("AVERAGE", None), 6: ("MIN", None), 7: ("MAX", None),
         8: ("ROW", None), 9: ("COLUMN", None), 10: ("NA", 0), 15: ("SIN", 1), 19: ("PI", 0), 20: ("SQRT", 1), 24: ("ABS", 1), 25: ("INT", 1), 26: ("SIGN", 1),
         27: ("ROUND", 2), 28: ("LOOKUP", None), 29: ("INDEX", None), 30: ("REPT", 2), 31: ("MID", 3), 32: ("LEN", 1), 33: ("VALUE", 1), 34: ("TRUE", 0),
         35: ("FALSE", 0), 36: ("AND", None), 37: ("OR", None), 38: ("NOT", 1), 39: ("MOD", 2), 48: ("TEXT", 2), 64: ("MATCH", None), 65: ("DATE", 3),
         66: ("TIME", 3), 67: ("DAY", 1), 68: ("MONTH", 1), 69: ("YEAR", 1), 70: ("WEEKDAY", None), 71: ("HOUR", 1), 72: ("MINUTE", 1), 73: ("SECOND", 1),
         74: ("NOW", 0), 77: ("COLUMNS", 1), 76: ("ROWS", 1), 78: ("OFFSET", None), 82: ("SEARCH", None), 100: ("CHOOSE", None), 101: ("HLOOKUP", None),
         102: ("VLOOKUP", None), 105: ("ISREF", 1), 111: ("CHAR", 1), 112: ("LOWER", 1), 113: ("UPPER", 1), 115: ("LEFT", None), 116: ("RIGHT", None),
         117: ("EXACT", 2), 118: ("TRIM", 1), 119: ("REPLACE", 4), 120: ("SUBSTITUTE", None), 121: ("CODE", 1), 124: ("FIND", None), 126: ("ISERR", 1),
         127: ("ISTEXT", 1), 128: ("ISNUMBER", 1), 129: ("ISBLANK", 1), 130: ("T", 1), 131: ("N", 1), 140: ("DATEVALUE", 1), 148: ("INDIRECT", None),
         162: ("CLEAN", 1), 169: ("COUNTA", None), 183: ("PRODUCT", None), 190: ("ISNONTEXT", 1), 197: ("TRUNC", None), 198: ("ISLOGICAL", 1),
         212: ("ROUNDUP", 2), 213: ("ROUNDDOWN", 2), 221: ("TODAY", 0), 228: ("SUMPRODUCT", None), 285: ("FLOOR", 2), 288: ("CEILING", 2),
         336: ("CONCATENATE", None), 337: ("POWER", 2), 344: ("SUBTOTAL", None), 345: ("SUMIF", None), 346: ("COUNTIF", 2), 347: ("COUNTBLANK", 1),
         351: ("DATEDIF", 3)}
BIN = {0x03: "+", 0x04: "-", 0x05: "*", 0x06: "/", 0x07: "^", 0x08: "&", 0x09: "<", 0x0A: "<=", 0x0B: "=", 0x0C: ">=", 0x0D: ">", 0x0E: "<>", 0x10: ",", 0x11: ":"}
NEUT = {"=": "==", "<>": "!="}
ERRS = {0: "#NULL!", 7: "#DIV/0!", 15: "#VALUE!", 23: "#REF!", 29: "#NAME?", 36: "#NUM!", 42: "#N/A"}


def letters(c):
    s = ""
    c += 1
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def ustr(b, off, cch_bytes):
    cch = b[off] if cch_bytes == 1 else struct.unpack_from("<H", b, off)[0]
    off += cch_bytes
    fl = b[off]; off += 1
    if fl & 1:
        return b[off:off + 2 * cch].decode("utf-16-le", "replace"), off + 2 * cch
    return b[off:off + cch].decode("latin-1"), off + cch


class Book:
    def __init__(self, path):
        data = cfb.CFB(path).read_stream("Workbook")
        if data is None:
            raise ValueError("không có luồng Workbook")
        recs, pos = [], 0
        while pos + 4 <= len(data):
            rid, ln = struct.unpack_from("<HH", data, pos)
            recs.append((pos, rid, data[pos + 4:pos + 4 + ln]))
            pos += 4 + ln
        self.sheets, self.xti, sst, supb = [], [], [], []
        i = 0
        while i < len(recs):
            pos, rid, b = recs[i]
            if rid == 0x0085:
                off, vis, typ = struct.unpack_from("<IBB", b, 0)
                name, _ = ustr(b, 6, 1)
                self.sheets.append({"off": off, "name": name, "hidden": vis, "type": typ})
            elif rid == 0x01AE:
                supb.append(b)
            elif rid == 0x0017:
                body = b; j = i + 1
                while j < len(recs) and recs[j][1] == 0x003C:
                    body += recs[j][2]; j += 1
                n = struct.unpack_from("<H", body, 0)[0]
                self.xti = [struct.unpack_from("<HhH", body, 2 + 6 * k) for k in range(n)]
            elif rid == 0x00FC:
                chunks = [b[8:]]; j = i + 1
                while j < len(recs) and recs[j][1] == 0x003C:
                    chunks.append(recs[j][2]); j += 1
                sst = self._sst(chunks, struct.unpack_from("<I", b, 4)[0])
            elif rid == 0x000A:
                break
            i += 1
        self.self_book = next((k for k, s in enumerate(supb) if len(s) == 4 and s[2:4] == b"\x01\x04"), 0)
        byoff = {p: k for k, (p, _, _) in enumerate(recs)}
        self.cells = []
        for sh in self.sheets:
            cells = {}
            k = byoff.get(sh["off"])
            if k is not None:
                self._sheet(recs, k + 1, cells, sst)
            self.cells.append(cells)

    @staticmethod
    def _sst(chunks, total):
        # chuỗi SST có thể vắt qua CONTINUE: ghép theo từng ký tự với byte cờ mới ở đầu chunk
        out, ci, off = [], 0, 0
        def take(n):
            nonlocal ci, off
            r = b""
            while n > 0:
                while ci < len(chunks) and off >= len(chunks[ci]):
                    ci += 1; off = 0
                if ci >= len(chunks):
                    raise EOFError
                m = min(n, len(chunks[ci]) - off)
                r += chunks[ci][off:off + m]; off += m; n -= m
            return r
        try:
            for _ in range(total):
                cch = struct.unpack("<H", take(2))[0]; fl = take(1)[0]
                rt = struct.unpack("<H", take(2))[0] if fl & 8 else 0
                sz = struct.unpack("<I", take(4))[0] if fl & 4 else 0
                hi, left, parts = fl & 1, cch, []
                while left > 0:
                    while ci < len(chunks) and off >= len(chunks[ci]):
                        ci += 1; off = 0
                    avail = len(chunks[ci]) - off
                    m = min(left, avail // 2 if hi else avail)
                    raw = take(2 * m if hi else m)
                    parts.append(raw.decode("utf-16-le" if hi else "latin-1", "replace")); left -= m
                    if left > 0:
                        ci += 1; off = 0; hi = take(1)[0] & 1
                take(4 * rt); take(sz)
                out.append("".join(parts))
        except EOFError:
            pass
        return out

    def _sheet(self, recs, k, cells, sst):
        last_formula = None
        while k < len(recs):
            _, rid, b = recs[k]
            if rid == 0x000A:
                break
            try:
                if rid == 0x00FD:
                    r, c, _, idx = struct.unpack_from("<HHHI", b, 0)
                    cells[(r, c)] = {"v": sst[idx] if idx < len(sst) else ""}
                elif rid == 0x0204:
                    r, c, _ = struct.unpack_from("<HHH", b, 0)
                    cells[(r, c)] = {"v": ustr(b, 6, 2)[0]}
                elif rid == 0x0203:
                    r, c, _, v = struct.unpack_from("<HHHd", b, 0)
                    cells[(r, c)] = {"v": v}
                elif rid == 0x027E:
                    r, c, _, rk = struct.unpack_from("<HHHI", b, 0)
                    cells[(r, c)] = {"v": self._rk(rk)}
                elif rid == 0x00BD:
                    r, c0 = struct.unpack_from("<HH", b, 0)
                    n = (len(b) - 6) // 6
                    for j in range(n):
                        _, rk = struct.unpack_from("<HI", b, 4 + 6 * j)
                        cells[(r, c0 + j)] = {"v": self._rk(rk)}
                elif rid == 0x0205:
                    r, c, _, v, e = struct.unpack_from("<HHHBB", b, 0)
                    cells[(r, c)] = {"v": (bool(v) if e == 0 else ERRS.get(v, "#ERR"))}
                elif rid == 0x0006:
                    r, c = struct.unpack_from("<HH", b, 0)
                    cce = struct.unpack_from("<H", b, 20)[0]
                    cell = {"rgce": b[22:22 + cce], "row": r, "col": c}
                    if b[12:14] == b"\xff\xff":
                        t = b[6]
                        cell["v"] = "" if t in (0, 3) else bool(b[8]) if t == 1 else ERRS.get(b[8], "#ERR")
                        if t == 0:
                            cell["pending"] = True
                    else:
                        cell["v"] = struct.unpack_from("<d", b, 6)[0]
                    cells[(r, c)] = cell
                    last_formula = cell
                elif rid == 0x0207 and last_formula is not None and last_formula.get("pending"):
                    last_formula["v"] = ustr(b, 0, 2)[0]; last_formula.pop("pending", None)
                elif rid == 0x04BC:
                    rf, rl, cf, cl = struct.unpack_from("<HHBB", b, 0)
                    cce = struct.unpack_from("<H", b, 8)[0]
                    self.shared = getattr(self, "shared", {})
                    self.shared[(id(cells), rf, cf)] = {"rgce": b[10:10 + cce], "rng": (rf, rl, cf, cl)}
            except (struct.error, IndexError, EOFError):
                pass
            k += 1

    @staticmethod
    def _rk(rk):
        v = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0] if not rk & 2 else float(struct.unpack("<i", struct.pack("<I", rk & 0xFFFFFFFC))[0] >> 2)
        return v / 100 if rk & 1 else v

    def sheet_of_ixti(self, ixti):
        if ixti >= len(self.xti):
            return None
        sb, f, _ = self.xti[ixti]
        return f if sb == self.self_book and f >= 0 else None

    def formula(self, si, r, c):
        """(tokens rgce, gốc r, gốc c, có phải công thức dùng chung) của ô; None nếu không phải công thức"""
        cell = self.cells[si].get((r, c))
        if not cell or "rgce" not in cell:
            return None
        g = cell["rgce"]
        if g and g[0] == 0x01:
            rw, co = struct.unpack_from("<HH", g, 1)
            sh = getattr(self, "shared", {}).get((id(self.cells[si]), rw, co))
            if sh is None:
                raise ValueError("ptgExp không có SHRFMLA (có thể là công thức mảng) tại %s%d" % (letters(c), r + 1))
            return sh["rgce"], r, c, True
        return g, r, c, False


def decode(book, si, rgce, br, bc, shared):
    """trả về (văn bản kiểu Excel, văn bản trung tính, danh sách tham chiếu (sheet, row0, col0), cấu trúc chỉ gồm ref và &)"""
    st, i, refs, only = [], 0, [], True

    def loc(rw, col, rel):
        r_rel, c_rel = col & 0x8000, col & 0x4000
        c = col & 0x3FFF
        if rel:
            r = br + (rw - 0x10000 if rw & 0x8000 else rw) if r_rel else rw
            c8 = c & 0xFF
            c = bc + (c8 - 0x100 if c8 & 0x80 else c8) if c_rel else c8
        else:
            r = rw
        return r, c

    def name_of(s):
        n = book.sheets[s]["name"]
        return n if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", n) else "'" + n.replace("'", "''") + "'"
    while i < len(rgce):
        p = rgce[i]; i += 1
        b = p if p < 0x40 else (p - 0x20 if p < 0x60 else p - 0x40)
        if b in BIN:
            y, x = st.pop(), st.pop()
            op = BIN[b]
            if op != "&":
                only = False
            st.append((x[0] + op + y[0], x[1] + " " + NEUT.get(op, op) + " " + y[1]))
        elif b in (0x12, 0x13):
            x = st.pop(); only = False
            st.append(("+-"[b - 0x12] + x[0], "+-"[b - 0x12] + x[1]))
        elif b == 0x14:
            x = st.pop(); only = False; st.append((x[0] + "%", x[1] + "/100"))
        elif b == 0x15:
            x = st.pop(); only = False; st.append(("(" + x[0] + ")", "(" + x[1] + ")"))
        elif b == 0x16:
            st.append(("", ""))
        elif b == 0x17:
            s, i = ustr(rgce, i, 1); only = False
            st.append(('"' + s.replace('"', '""') + '"', json.dumps(s, ensure_ascii=False)))
        elif b == 0x19:
            g = rgce[i]; i += 1
            if g & 0x04:
                n = struct.unpack_from("<H", rgce, i)[0]; i += 2 + 2 * (n + 1)
            else:
                i += 2
            if g & 0x10:
                x = st.pop(); only = False; st.append(("SUM(" + x[0] + ")", "SUM(" + x[1] + ")"))
        elif b == 0x1C:
            only = False; e = ERRS.get(rgce[i], "#ERR"); i += 1; st.append((e, e))
        elif b == 0x1D:
            only = False; v = "TRUE" if rgce[i] else "FALSE"; i += 1; st.append((v, v.lower()))
        elif b == 0x1E:
            only = False; v = str(struct.unpack_from("<H", rgce, i)[0]); i += 2; st.append((v, v))
        elif b == 0x1F:
            only = False; d = struct.unpack_from("<d", rgce, i)[0]; i += 8
            v = str(int(d)) if d == int(d) else repr(d); st.append((v, v))
        elif b in (0x21, 0x22):
            only = False
            if b == 0x21:
                fi = struct.unpack_from("<H", rgce, i)[0]; i += 2
                nm, argc = FUNCS.get(fi, (None, None))
                if nm is None or argc is None:
                    raise ValueError("hàm cố định chưa biết số đối số iftab=%d" % fi)
            else:
                argc = rgce[i]; fi = struct.unpack_from("<H", rgce, i + 1)[0] & 0x7FFF; i += 3
                nm = FUNCS.get(fi, ("FN%d" % fi, None))[0]
                if fi == 255:
                    argc -= 1
                    nm = st.pop(-argc - 1)[0] if argc >= 0 else nm
            args = st[len(st) - argc:] if argc else []
            del st[len(st) - argc:]
            st.append((nm + "(" + ",".join(a[0] for a in args) + ")", nm + "(" + ", ".join(a[1] for a in args) + ")"))
        elif b in (0x24, 0x2C):
            rw, col = struct.unpack_from("<HH", rgce, i); i += 4
            r, c = loc(rw, col, b == 0x2C or shared)
            refs.append((si, r, c))
            st.append((letters(c) + str(r + 1), "cell(S%d!%s%d)" % (si, letters(c), r + 1)))
        elif b in (0x25, 0x2D):
            rf, rl, cf, cl = struct.unpack_from("<HHHH", rgce, i); i += 8; only = False
            r1, c1 = loc(rf, cf, b == 0x2D or shared); r2, c2 = loc(rl, cl, b == 0x2D or shared)
            st.append(("%s%d:%s%d" % (letters(c1), r1 + 1, letters(c2), r2 + 1), "range(S%d!%s%d:%s%d)" % (si, letters(c1), r1 + 1, letters(c2), r2 + 1)))
        elif b == 0x3A:
            ixti, rw, col = struct.unpack_from("<HHH", rgce, i); i += 6
            s = book.sheet_of_ixti(ixti)
            if s is None:
                raise ValueError("ptgRef3d tới sổ ngoài ixti=%d" % ixti)
            r, c = loc(rw, col, shared)
            refs.append((s, r, c))
            st.append(("%s!%s%d" % (name_of(s), letters(c), r + 1), "cell(S%d!%s%d)" % (s, letters(c), r + 1)))
        elif b == 0x3B:
            ixti, rf, rl, cf, cl = struct.unpack_from("<HHHHH", rgce, i); i += 10; only = False
            s = book.sheet_of_ixti(ixti)
            if s is None:
                raise ValueError("ptgArea3d tới sổ ngoài ixti=%d" % ixti)
            r1, c1 = loc(rf, cf, shared); r2, c2 = loc(rl, cl, shared)
            st.append(("%s!%s%d:%s%d" % (name_of(s), letters(c1), r1 + 1, letters(c2), r2 + 1), "range(S%d!%s%d:%s%d)" % (s, letters(c1), r1 + 1, letters(c2), r2 + 1)))
        elif b in (0x2A, 0x2B, 0x3C, 0x3D):
            i += {0x2A: 4, 0x2B: 8, 0x3C: 6, 0x3D: 10}[b]; only = False; st.append(("#REF!", "#REF!"))
        elif b in (0x26, 0x27, 0x28):
            i += 6
        elif b == 0x29:
            i += 2
        else:
            raise ValueError("token công thức chưa hỗ trợ 0x%02X" % p)
    if len(st) != 1:
        raise ValueError("ngăn xếp công thức còn %d phần tử" % len(st))
    return st[0][0], st[0][1], refs, only


def main():
    classes = json.load(open(classes_path))["classes"]
    man = list(csv.DictReader(open(manifest_path, encoding="utf-8")))
    folder_of = {}
    for m in man:
        mm = re.search(r"MaFile=(\S+) FolderName=(\S+)", m["evidence"])
        if mm:
            folder_of.setdefault(mm.group(1), mm.group(2))
    dsts = {m["dst"] for m in man}
    fam_mafile = {}
    for c in classes:
        mm = re.match(r"(.+?)Nckt(\d+)$", c["class"])
        if mm and c.get("MaFile"):
            fam_mafile.setdefault(mm.group(1), c["MaFile"])
    out = {}
    for c in sorted(classes, key=lambda x: x["class"]):
        src = open(os.path.join(dec, c["path"]), encoding="utf-8-sig").read()
        if "checkLockUnLockCells(" not in src:
            continue
        cname = c["class"]
        rec = {"class": cname}
        out[cname] = rec
        mm = re.match(r"(.+?)Nckt(\d)(\d)(\d)$", cname)
        mafile = c.get("MaFile") or (fam_mafile.get(mm.group(1)) if mm else None)
        rec["MaFile"] = mafile
        rec["MaFile_from"] = "classes.json" if c.get("MaFile") else "class cùng họ " + (mm.group(1) if mm else "?")
        if not mm or not mafile or mafile not in folder_of:
            rec["error"] = "không xác định được MaFile/FolderName của class"
            continue
        ver = "%s.%s.%s" % mm.group(2, 3, 4)
        fold = folder_of[mafile]
        cand = sorted(d for d in dsts if d.endswith("/%s/%s_v%s_mau.xls" % (fold, fold, ver)))
        if not cand:
            rec["error"] = "không có mẫu %s_v%s_mau.xls trong manifest" % (fold, ver)
            continue
        rec["template"] = "_forms/" + cand[0]
        try:
            book = Book(os.path.join(forms, cand[0]))
        except Exception as ex:  # tệp lỗi: ghi lý do
            rec["error"] = "không đọc được mẫu: %s" % ex
            continue
        ws = [k for k, s in enumerate(book.sheets) if s["type"] == 0]
        rec["sheets"] = [book.sheets[k]["name"] for k in ws]
        last = ws[-1]
        rec["lastSheet"] = last
        ba1 = book.cells[last].get((0, BA), {}).get("v")
        rec["BA1"] = ba1 if not isinstance(ba1, float) else (str(int(ba1)) if ba1 == int(ba1) else str(ba1))
        rows = []
        rec["rows"] = rows
        maxr = max([r for (r, _c) in book.cells[last].keys()] + [0])
        for i in range(1, maxr + 1):
            row = {"row": i + 1}
            try:
                f = book.formula(last, i, AK)
                if f is None:
                    rec["stop"] = "dòng %d: ô AK không có công thức, vòng lặp dừng (break)" % (i + 1)
                    break
                ak_x, ak_n, ak_refs, ak_only = decode(book, last, *f)
                row["AK"] = ak_x
                if len(ak_refs) != 1 or not ak_only or "!" not in ak_x:
                    row["skip"] = "AK không phải một tham chiếu sang sheet khác (GetCellValild không tách được)"
                    rows.append(row); continue
                s_ref, r_ref, c_ref = ak_refs[0]
                det = next((k for k, s in enumerate(book.sheets) if s["name"] and s["name"] in ak_x), -1)
                row["controlRow"] = r_ref + 1
                row["controlCol"] = letters(c_ref)
                row["refSheet"] = s_ref
                row["sheet"] = det
                cached = book.cells[last][(i, AK)].get("v")
                tgt = book.cells[s_ref].get((r_ref, c_ref), {}).get("v")
                row["selfCheck"] = (cached == tgt) or (cached in ("", 0.0) and tgt is None)
                if det < 0:
                    row["skip"] = "DetectIndexOfSheet không tìm thấy sheet: continue"
                    rows.append(row); continue
                fa = book.formula(last, i, AL)
                if fa is not None:
                    al_x, al_n, _, _ = decode(book, last, *fa)
                    row["AL"] = al_x; row["ALn"] = al_n
                alv = book.cells[last].get((i, AL), {}).get("v")
                row["ALvalue"] = alv
                fm = book.formula(last, i, AM)
                if fm is None:
                    row["AMempty"] = True
                else:
                    am_x, am_n, am_refs, am_only = decode(book, last, *fm)
                    row["AM"] = am_x
                    if not am_only:
                        row["skip"] = "AM không phải danh sách tham chiếu nối bằng &"
                    else:
                        row["targets"] = [{"row": r + 1, "col": letters(cc), "refSheet": s} for (s, r, cc) in am_refs]
                anc = book.cells[last].get((i, AN), {})
                anv = anc.get("v")
                row["AN"] = (str(int(anv)) if isinstance(anv, float) and anv == int(anv) else anv)
                row["ANformula"] = "rgce" in anc
            except ValueError as ex:
                row["error"] = str(ex)
            rows.append(row)
    json.dump(out, open(out_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1, sort_keys=True)
    n_ok = sum(1 for r in out.values() if "rows" in r)
    n_t = sum(len(x.get("targets", [])) for r in out.values() for x in r.get("rows", []))
    print("class=%d có mẫu=%d dòng cấu hình=%d ô đích=%d lỗi dòng=%d selfCheck sai=%d" % (
        len(out), n_ok, sum(len(r.get("rows", [])) for r in out.values()), n_t,
        sum(1 for r in out.values() for x in r.get("rows", []) if "error" in x),
        sum(1 for r in out.values() for x in r.get("rows", []) if x.get("selfCheck") is False)))
    for k, r in out.items():
        if "error" in r:
            print("  ", k, r["error"])


main()
