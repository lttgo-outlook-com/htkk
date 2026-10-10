# Bộ đọc BIFF8 tối giản: chỉ đọc text ô (SST/LABELSST/LABEL/FORMULA+STRING/RK/NUMBER) theo thứ tự sheet.
# Chỉ parse byte; không chạy macro hay mã nào trong tệp.
import struct

try:
    import olefile
except ImportError:  # máy không có olefile: dùng bộ đọc OLE2 tối giản cùng thư mục
    olefile = None
    import cfb


def _read_stream(path):
    if olefile is None:
        c = cfb.CFB(path)
        for name in ("Workbook", "Book"):
            s = c.read_stream(name)
            if s is not None:
                return s
        return None
    ole = olefile.OleFileIO(path)
    try:
        for name in ("Workbook", "Book"):
            if ole.exists(name):
                return ole.openstream(name).read()
    finally:
        ole.close()
    return None


def _records(data):
    pos, n = 0, len(data)
    while pos + 4 <= n:
        rid, ln = struct.unpack_from("<HH", data, pos)
        yield pos, rid, data[pos + 4: pos + 4 + ln]
        pos += 4 + ln


class _Reader:
    """Đọc chuỗi Unicode BIFF8 trải qua nhiều bản ghi CONTINUE."""

    def __init__(self, chunks):
        self.chunks = chunks
        self.ci = 0
        self.off = 0

    def _need(self):
        while self.ci < len(self.chunks) and self.off >= len(self.chunks[self.ci]):
            self.ci += 1
            self.off = 0

    def take(self, k):
        out = b""
        while k > 0:
            self._need()
            if self.ci >= len(self.chunks):
                raise EOFError
            c = self.chunks[self.ci]
            m = min(k, len(c) - self.off)
            out += c[self.off:self.off + m]
            self.off += m
            k -= m
        return out

    def string(self, cch_bytes=2):
        cch = struct.unpack("<H", self.take(2))[0] if cch_bytes == 2 else self.take(1)[0]
        flags = self.take(1)[0]
        rt = struct.unpack("<H", self.take(2))[0] if flags & 0x08 else 0
        sz = struct.unpack("<I", self.take(4))[0] if flags & 0x04 else 0
        chars = []
        high = flags & 0x01
        left = cch
        while left > 0:
            self._need()
            if self.ci >= len(self.chunks):
                raise EOFError
            c = self.chunks[self.ci]
            avail = len(c) - self.off
            if high:
                m = min(left, avail // 2)
                chars.append(c[self.off:self.off + 2 * m].decode("utf-16-le", "replace"))
                self.off += 2 * m
            else:
                m = min(left, avail)
                chars.append(c[self.off:self.off + m].decode("latin-1"))
                self.off += m
            left -= m
            if left > 0:
                # chuỗi bị cắt sang CONTINUE: byte cờ mới ở đầu chunk kế tiếp
                self.ci += 1
                self.off = 0
                self._need()
                high = self.take(1)[0] & 0x01
        self.take(4 * rt)
        self.take(sz)
        return "".join(chars)


def read_workbook(path):
    """Trả về list[(tên sheet, {(row0, col0): text})] theo thứ tự BOUNDSHEET."""
    data = _read_stream(path)
    if data is None:
        return None
    recs = list(_records(data))
    sheets = []  # (offset, name)
    sst = []
    i = 0
    while i < len(recs):
        pos, rid, body = recs[i]
        if rid == 0x0085:  # BOUNDSHEET
            off = struct.unpack_from("<I", body, 0)[0]
            r = _Reader([body[6:]])
            try:
                name = r.string(cch_bytes=1)
            except EOFError:
                name = ""
            sheets.append((off, name))
        elif rid == 0x00FC:  # SST
            chunks = [body[8:]]
            total = struct.unpack_from("<I", body, 4)[0]
            j = i + 1
            while j < len(recs) and recs[j][1] == 0x003C:
                chunks.append(recs[j][2])
                j += 1
            r = _Reader(chunks)
            for _ in range(total):
                try:
                    sst.append(r.string())
                except EOFError:
                    break
            i = j
            continue
        elif rid == 0x000A and sheets and pos >= sheets[0][0]:
            break
        i += 1
    offs = {p: k for k, (p, _) in enumerate(recs and [(r[0], None) for r in recs])}
    out = []
    for off, name in sheets:
        cells = {}
        k = offs.get(off)
        if k is None:
            out.append((name, cells))
            continue
        k += 1
        pending = None
        while k < len(recs):
            _, rid, body = recs[k]
            if rid == 0x000A:  # EOF
                break
            try:
                if rid == 0x00FD:  # LABELSST
                    r, c, _x, idx = struct.unpack_from("<HHHI", body, 0)
                    if idx < len(sst):
                        cells[(r, c)] = sst[idx]
                elif rid == 0x0204:  # LABEL
                    r, c, _x = struct.unpack_from("<HHH", body, 0)
                    cells[(r, c)] = _Reader([body[6:]]).string()
                elif rid == 0x0006:  # FORMULA
                    r, c = struct.unpack_from("<HH", body, 0)
                    # kết quả chuỗi: byte 6 = 0x00 và byte 12-13 = 0xFFFF; text nằm ở bản ghi STRING kế tiếp
                    if body[12:14] == b"\xff\xff" and body[6] == 0x00:
                        pending = (r, c)
                    else:
                        pending = None
                elif rid == 0x0207 and pending:  # STRING sau FORMULA
                    cells[pending] = _Reader([body]).string()
                    pending = None
            except (EOFError, struct.error):
                pass
            k += 1
        out.append((name, cells))
    return out
