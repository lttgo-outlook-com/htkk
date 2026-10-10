# Bộ đọc Compound File Binary (OLE2) tối giản, dùng khi máy không có gói olefile.
# Chỉ đọc luồng con trực tiếp của root (vd. "Workbook"); chỉ parse byte, không chạy mã trong tệp.
import struct

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF


class CFB:
    def __init__(self, path):
        with open(path, "rb") as f:
            self.data = f.read()
        d = self.data
        if d[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("không phải tệp OLE2")
        self.ss = 1 << struct.unpack_from("<H", d, 0x1E)[0]
        self.mss = 1 << struct.unpack_from("<H", d, 0x20)[0]
        first_dir = struct.unpack_from("<I", d, 0x30)[0]
        self.cutoff = struct.unpack_from("<I", d, 0x38)[0]
        first_minifat = struct.unpack_from("<I", d, 0x3C)[0]
        first_difat, n_difat = struct.unpack_from("<II", d, 0x44)
        difat = [x for x in struct.unpack_from("<109I", d, 0x4C)]
        sec = first_difat
        per = self.ss // 4 - 1
        for _ in range(n_difat):
            if sec in (ENDOFCHAIN, FREESECT):
                break
            vals = struct.unpack_from("<%dI" % (per + 1), d, self._off(sec))
            difat.extend(vals[:per])
            sec = vals[per]
        self.fat = []
        for s in difat:
            if s in (ENDOFCHAIN, FREESECT):
                continue
            self.fat.extend(struct.unpack_from("<%dI" % (self.ss // 4), d, self._off(s)))
        dirdata = self._chain(first_dir)
        self.entries = []
        for i in range(len(dirdata) // 128):
            e = dirdata[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace")
            typ = e[66]
            left, right, child = struct.unpack_from("<III", e, 68)
            start, size = struct.unpack_from("<II", e, 116)
            self.entries.append((name, typ, left, right, child, start, size))
        root = self.entries[0]
        self.ministream = self._chain(root[5])[:root[6]] if root[5] != ENDOFCHAIN else b""
        self.minifat = []
        if first_minifat not in (ENDOFCHAIN, FREESECT):
            mf = self._chain(first_minifat)
            self.minifat = list(struct.unpack_from("<%dI" % (len(mf) // 4), mf, 0))

    def _off(self, sec):
        return (sec + 1) * self.ss

    def _chain(self, start):
        out = []
        sec = start
        seen = set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < len(self.fat) and sec not in seen:
            seen.add(sec)
            o = self._off(sec)
            out.append(self.data[o:o + self.ss])
            sec = self.fat[sec]
        return b"".join(out)

    def _minichain(self, start):
        out = []
        sec = start
        seen = set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < len(self.minifat) and sec not in seen:
            seen.add(sec)
            o = sec * self.mss
            out.append(self.ministream[o:o + self.mss])
            sec = self.minifat[sec]
        return b"".join(out)

    def root_children(self):
        res = []
        stack = [self.entries[0][4]]
        seen = set()
        while stack:
            i = stack.pop()
            if i == FREESECT or i >= len(self.entries) or i in seen:
                continue
            seen.add(i)
            e = self.entries[i]
            res.append(e)
            stack.extend([e[2], e[3]])
        return res

    def read_stream(self, name):
        for e in self.root_children():
            if e[1] == 2 and e[0].lower() == name.lower():
                if e[6] < self.cutoff:
                    return self._minichain(e[5])[:e[6]]
                return self._chain(e[5])[:e[6]]
        return None
