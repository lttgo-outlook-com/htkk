#!/usr/bin/env python3
"""Chuyen danh muc HTKK 5.7.9 (catalogue_*, dm_*, convertdatacommon, ...) sang JSON.

Chay: python3 -I _spec/tools/catalog.py [--src DIR] [--dec DIR] [--out DIR]
Khong chay ma/macro cua nguon; parse XML tu chuoi byte, tu choi DOCTYPE/ENTITY.
.xls: parser BIFF8 tu viet (olefile) lay gia tri o; so o doi chieu voi dem ban ghi BIFF
bang ham doc lap (count_biff_cells). XML: so phan tu doi chieu voi dem the bang regex.
"""
import hashlib
import json
import os
import re
import struct
import sys
import xml.etree.ElementTree as ET

import olefile

GENERATOR = "_spec/tools/catalog.py"
VERSION = 1
SRC_DEFAULT = "/Users/tuanle/codes/HTKK/_extracted"
DEC_DEFAULT = "/Users/tuanle/codes/HTKK/_decompiled"
NAME_CUT = 28  # ten tep trong _extracted bi cat o 28 ky tu

# nhom danh muc theo tien to ten tep (khong phan biet hoa thuong)
GROUPS = [
    ("catalogue", re.compile(r"^catalogue")),
    ("dm", re.compile(r"^dm_|^dmngayle")),
    ("convert", re.compile(r"^(convertdatacommon|converdatafile|convertversion|replacedatafile|createdatafilekhbsmiss)")),
    ("khac", re.compile(r"^(hieulucbanggia|validcombobox|versiontokhai|map)")),
]

# ---------------------------------------------------------------- tien ich


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for blk in iter(lambda: f.read(1 << 20), b""):
            h.update(blk)
    return h.hexdigest()


def sniff(head):
    if head[:4] == bytes.fromhex("D0CF11E0"):
        return "ole"
    if head[:2] == b"PK":
        return "zip"
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "xml"  # BOM UTF-16
    if head[:3] == b"\xef\xbb\xbf" or head.lstrip()[:1] == b"<":
        return "xml"
    if head.lstrip()[:1] in (b"{", b"["):
        return "json"
    return "khac"


def inventory(src):
    out = []
    for name in sorted(os.listdir(src)):
        p = os.path.join(src, name)
        if not os.path.isfile(p):
            continue
        low = name.lower()
        grp = None
        for g, rx in GROUPS:
            if rx.match(low):
                grp = g
                break
        if not grp:
            continue
        with open(p, "rb") as f:
            head = f.read(64)
        out.append((name, grp, sniff(head)))
    return out


# ---------------------------------------------------------------- XML


def decode_xml_bytes(raw):
    """Tra (text_hoac_bytes, ten_ma_hoa). Tu choi DOCTYPE/ENTITY."""
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        enc = "utf-16"
        probe = raw.decode("utf-16", errors="strict")
    elif raw[:3] == b"\xef\xbb\xbf":
        enc = "utf-8-sig"
        probe = raw.decode("utf-8-sig", errors="strict")
    else:
        m = re.match(rb'\s*<\?xml[^>]*encoding\s*=\s*["\']([\w.-]+)["\']', raw)
        enc = m.group(1).decode("ascii").lower() if m else "utf-8 (mac dinh)"
        codec = enc.split(" ")[0]
        probe = raw.decode(codec, errors="strict")
    up = probe.upper()
    if "<!DOCTYPE" in up or "<!ENTITY" in up:
        raise ValueError("co DOCTYPE/ENTITY, tu choi")
    return probe, enc


def parse_xml(raw):
    probe, enc = decode_xml_bytes(raw)
    builder = ET.TreeBuilder(insert_comments=True)
    parser = ET.XMLParser(target=builder)
    data = probe.lstrip("﻿")
    # bo khai bao encoding de ET khong doi chieu lai (chuoi da la unicode)
    data = re.sub(r'^\s*<\?xml[^>]*\?>', "", data, count=1)
    parser.feed(data)
    root = parser.close()
    return root, enc, probe


def el_to_json(el):
    node = {"tag": el.tag}
    if el.attrib:
        node["attrs"] = dict(el.attrib)
    if el.text and el.text.strip():
        node["text"] = el.text
    kids = []
    for ch in el:
        if ch.tag is ET.Comment:
            c = {"comment": ch.text}
            kids.append(c)
        else:
            kids.append(el_to_json(ch))
        if ch.tail and ch.tail.strip():
            kids[-1]["tail"] = ch.tail
    if kids:
        node["children"] = kids
    return node


def json_stats(node):
    """Tra (so_phan_tu, so_item, so_la). Duyet khong de quy."""
    n = items = leaves = 0
    stack = [node]
    while stack:
        x = stack.pop()
        if "tag" not in x:
            continue
        n += 1
        if x["tag"].lower() == "item":
            items += 1
        ch = [c for c in x.get("children", []) if "tag" in c]
        if not ch:
            leaves += 1
        stack.extend(ch)
    return n, items, leaves


def count_xml_tags(probe):
    """Dem phan tu tu nguon bang regex (doc lap voi parser)."""
    s = re.sub(r"<!--.*?-->", "", probe, flags=re.S)
    s = re.sub(r"<!\[CDATA\[.*?\]\]>", "", s, flags=re.S)
    s = re.sub(r"<\?.*?\?>", "", s, flags=re.S)
    return len(re.findall(r"<[A-Za-z_:]", s))


def convert_xml(path, raw):
    root, enc, probe = parse_xml(raw)
    tree = el_to_json(root)
    n, items, leaves = json_stats(tree)
    return {"kind": "xml", "encoding": enc, "root": tree}, {
        "so_phan_tu": n,
        "so_muc": items if items else leaves,
        "so_muc_cach_dem": "phan tu item" if items else "phan tu la",
        "root_tag": root.tag,
    }


# ---------------------------------------------------------------- JSON thuan


def strip_trailing_commas(text):
    """Bo dau phay thua truoc } hoac ] (ngoai chuoi). Tra (text, so_dau_bo)."""
    out = []
    removed = 0
    in_str = False
    esc = False
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if in_str:
            out.append(ch)
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        elif ch == '"':
            in_str = True
            out.append(ch)
        elif ch == ",":
            j = i + 1
            while j < n and text[j] in " \t\r\n":
                j += 1
            if j < n and text[j] in "}]":
                removed += 1
            else:
                out.append(ch)
        else:
            out.append(ch)
        i += 1
    return "".join(out), removed


def convert_json(path, raw):
    text = raw.decode("utf-8-sig")
    note = None
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        text2, removed = strip_trailing_commas(text)
        data = json.loads(text2)
        note = "nguon co %d dau phay thua truoc dong ngoac; da bo de parse" % removed
    st = {
        "so_phan_tu": 0,
        "so_muc": len(data) if hasattr(data, "__len__") else 1,
        "so_muc_cach_dem": "phan tu cap 1",
    }
    if note:
        st["ghi_chu"] = note
    return {"kind": "json", "encoding": "utf-8", "data": data}, st


# ---------------------------------------------------------------- BIFF8

REC_CELL_COUNT = {0xFD, 0x204, 0x203, 0x27E, 0x06, 0x205, 0xD6}  # LABELSST LABEL NUMBER RK FORMULA BOOLERR RSTRING


def col_letters(c):
    s = ""
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def read_stream(path):
    ole = olefile.OleFileIO(path)
    try:
        for nm in ("Workbook", "Book"):
            if ole.exists(nm):
                return ole.openstream(nm).read(), nm
        raise ValueError("khong co stream Workbook/Book")
    finally:
        ole.close()


def iter_records(data):
    pos = 0
    n = len(data)
    while pos + 4 <= n:
        rid, ln = struct.unpack_from("<HH", data, pos)
        if pos + 4 + ln > n:
            break
        yield pos, rid, data[pos + 4: pos + 4 + ln]
        pos += 4 + ln


def count_biff_cells(data):
    """Dem doc lap: so ban ghi o co gia tri (khong tinh BLANK/MULBLANK). MULRK tinh theo so RK."""
    c = {}
    total = 0
    for _, rid, body in iter_records(data):
        if rid in REC_CELL_COUNT:
            c[rid] = c.get(rid, 0) + 1
            total += 1
        elif rid == 0xBD:
            k = (len(body) - 6) // 6
            c[rid] = c.get(rid, 0) + k
            total += k
    return total, c


class Chunks:
    """Doc chuoi unicode BIFF8 qua nhieu CONTINUE."""

    def __init__(self, chunks):
        self.ch = chunks
        self.i = 0
        self.o = 0

    def _avail(self):
        return len(self.ch[self.i]) - self.o

    def _next(self):
        self.i += 1
        self.o = 0

    def take(self, n):
        out = b""
        while n > 0:
            if self._avail() == 0:
                if self.i + 1 >= len(self.ch):
                    raise ValueError("het du lieu SST")
                self._next()
            k = min(n, self._avail())
            out += self.ch[self.i][self.o:self.o + k]
            self.o += k
            n -= k
        return out

    def read_string(self):
        if self._avail() == 0 and self.i + 1 < len(self.ch):
            self._next()
        cch, flags = struct.unpack("<HB", self.take(3))
        rich = flags & 0x08
        ext = flags & 0x04
        runs = struct.unpack("<H", self.take(2))[0] if rich else 0
        extlen = struct.unpack("<I", self.take(4))[0] if ext else 0
        wide = flags & 0x01
        parts = []
        remain = cch
        while remain > 0:
            if self._avail() == 0:
                if self.i + 1 >= len(self.ch):
                    raise ValueError("het du lieu SST (chuoi)")
                self._next()
                wide = self.ch[self.i][0] & 0x01
                self.o = 1
            av = self._avail()
            width = 2 if wide else 1
            k = min(remain, av // width)
            if k == 0:  # ky tu rong nua o ranh gioi: khong xay ra voi du lieu hop le
                raise ValueError("ranh gioi CONTINUE le")
            b = self.ch[self.i][self.o:self.o + k * width]
            self.o += k * width
            parts.append(b.decode("utf-16-le") if wide else b.decode("latin-1"))
            remain -= k
        if runs:
            self.take(4 * runs)
        if extlen:
            self.take(extlen)
        return "".join(parts)


def parse_sst(data, recs):
    """recs: danh sach (rid, body) lien tiep SST + CONTINUE."""
    chunks = [recs[0][1][8:]]
    total, uniq = struct.unpack("<II", recs[0][1][:8])
    for rid, body in recs[1:]:
        chunks.append(body)
    rd = Chunks(chunks)
    out = []
    for _ in range(uniq):
        out.append(rd.read_string())
    return out


def decode_rk(v):
    if v & 2:
        x = v >> 2
        if x & 0x20000000:
            x -= 0x40000000
        num = float(x)
    else:
        num = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    if v & 1:
        num /= 100.0
    return num


def num_out(x):
    if x == x and abs(x) < 1e15 and float(x).is_integer():
        return int(x)
    return x


ERRS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A"}


def read_biff_str16(body, off):
    """Chuoi unicode voi do dai 2 byte (LABEL/STRING). Tra (chuoi, off_moi)."""
    cch = struct.unpack_from("<H", body, off)[0]
    flags = body[off + 2]
    off += 3
    if flags & 0x08:
        off += 2
    if flags & 0x04:
        off += 4
    if flags & 1:
        s = body[off:off + 2 * cch].decode("utf-16-le")
        off += 2 * cch
    else:
        s = body[off:off + cch].decode("latin-1")
        off += cch
    return s, off


def parse_xls(path):
    data, stream = read_stream(path)
    recs = list(iter_records(data))
    # BOF dau tien
    rid0, body0 = recs[0][1], recs[0][2]
    if rid0 != 0x809:
        raise ValueError("khong phai BIFF8 (BOF 0x%X)" % rid0)
    ver, dt = struct.unpack_from("<HH", body0, 0)
    if ver != 0x600:
        raise ValueError("BIFF version 0x%X chua ho tro" % ver)
    # globals
    sheets = []
    sst = []
    encrypted = False
    i = 0
    depth = 0
    n = len(recs)
    while i < n:
        pos, rid, body = recs[i]
        if rid == 0x809:
            depth += 1
        elif rid == 0x0A:
            depth -= 1
            if depth == 0:
                break
        elif rid == 0x2F:
            encrypted = True
        elif rid == 0x85:
            off, state, typ = struct.unpack_from("<IBB", body, 0)
            cch, fl = body[6], body[7]
            nm = body[8:8 + cch * (2 if fl & 1 else 1)].decode("utf-16-le" if fl & 1 else "latin-1")
            sheets.append({"offset": off, "state": state, "type": typ, "name": nm})
        elif rid == 0xFC:
            grp = [(rid, body)]
            j = i + 1
            while j < n and recs[j][1] == 0x3C:
                grp.append((recs[j][1], recs[j][2]))
                j += 1
            sst = parse_sst(data, grp)
            i = j - 1
        i += 1
    if encrypted:
        raise ValueError("tep ma hoa (FILEPASS)")
    pos_to_idx = {r[0]: k for k, r in enumerate(recs)}
    out_sheets = []
    total_cells = 0
    for si, sh in enumerate(sheets):
        entry = {"index": si, "name": sh["name"], "state": sh["state"], "type": sh["type"]}
        if sh["type"] != 0:
            entry["rows"] = []
            entry["so_o"] = 0
            entry["so_hang"] = 0
            out_sheets.append(entry)
            continue
        k = pos_to_idx[sh["offset"]]
        cells = {}
        depth = 0
        while k < n:
            pos, rid, body = recs[k]
            if rid == 0x809:
                depth += 1
            elif rid == 0x0A:
                depth -= 1
                if depth == 0:
                    break
            elif depth == 1:
                if rid == 0xFD:
                    r, c, xf, isst = struct.unpack_from("<HHHI", body, 0)
                    cells[(r, c)] = {"v": sst[isst], "t": "s"}
                elif rid == 0x204:
                    r, c, xf = struct.unpack_from("<HHH", body, 0)
                    s, _ = read_biff_str16(body, 6)
                    cells[(r, c)] = {"v": s, "t": "s"}
                elif rid == 0xD6:
                    r, c, xf = struct.unpack_from("<HHH", body, 0)
                    s, _ = read_biff_str16(body, 6)
                    cells[(r, c)] = {"v": s, "t": "s"}
                elif rid == 0x203:
                    r, c, xf = struct.unpack_from("<HHH", body, 0)
                    cells[(r, c)] = {"v": num_out(struct.unpack_from("<d", body, 6)[0]), "t": "n"}
                elif rid == 0x27E:
                    r, c, xf, rk = struct.unpack_from("<HHHI", body, 0)
                    cells[(r, c)] = {"v": num_out(decode_rk(rk)), "t": "n"}
                elif rid == 0xBD:
                    r, c0 = struct.unpack_from("<HH", body, 0)
                    cnt = (len(body) - 6) // 6
                    for q in range(cnt):
                        rk = struct.unpack_from("<I", body, 4 + q * 6 + 2)[0]
                        cells[(r, c0 + q)] = {"v": num_out(decode_rk(rk)), "t": "n"}
                elif rid == 0x205:
                    r, c, xf, val, isErr = struct.unpack_from("<HHHBB", body, 0)
                    if isErr:
                        cells[(r, c)] = {"v": ERRS.get(val, "#ERR%d" % val), "t": "e"}
                    else:
                        cells[(r, c)] = {"v": bool(val), "t": "b"}
                elif rid == 0x06:
                    r, c, xf = struct.unpack_from("<HHH", body, 0)
                    res = body[6:14]
                    cell = {"f": 1}
                    if res[6:8] == b"\xff\xff":
                        kind = res[0]
                        if kind == 0:  # chuoi o ban ghi STRING ke tiep
                            m = k + 1
                            sval = None
                            while m < n and recs[m][1] in (0x4BC, 0x221, 0x236, 0x207):
                                if recs[m][1] == 0x207:
                                    chunks = [recs[m][2]]
                                    m2 = m + 1
                                    while m2 < n and recs[m2][1] == 0x3C:
                                        chunks.append(recs[m2][2])
                                        m2 += 1
                                    if len(chunks) == 1:
                                        sval, _ = read_biff_str16(chunks[0], 0)
                                    else:
                                        sval = Chunks(chunks).read_string()
                                    break
                                m += 1
                            cell.update({"v": sval if sval is not None else "", "t": "s"})
                        elif kind == 1:
                            cell.update({"v": bool(res[2]), "t": "b"})
                        elif kind == 2:
                            cell.update({"v": ERRS.get(res[2], "#ERR%d" % res[2]), "t": "e"})
                        else:
                            cell.update({"v": "", "t": "s"})
                    else:
                        cell.update({"v": num_out(struct.unpack("<d", res)[0]), "t": "n"})
                    cells[(r, c)] = cell
            k += 1
        rows = {}
        for (r, c) in sorted(cells):
            d = {"a": "%s%d" % (col_letters(c), r + 1)}
            d.update(cells[(r, c)])
            rows.setdefault(r + 1, []).append(d)
        entry["rows"] = [{"r": r, "cells": rows[r]} for r in sorted(rows)]
        entry["so_hang"] = len(rows)
        entry["so_o"] = len(cells)
        total_cells += len(cells)
        out_sheets.append(entry)
    return {"kind": "xls", "encoding": "BIFF8 (UTF-16LE/latin-1 trong SST)", "stream": stream,
            "sheets": out_sheets}, data


def convert_xls(path):
    doc, data = parse_xls(path)
    sheets = doc["sheets"]
    return doc, {
        "so_sheet": len(sheets),
        "so_hang": sum(s["so_hang"] for s in sheets),
        "so_o": sum(s["so_o"] for s in sheets),
        "so_o_biff": count_biff_cells(data)[0],
    }


# ---------------------------------------------------------------- tim noi doc tep trong .cs

_CS_CACHE = None


def load_cs(dec):
    global _CS_CACHE
    if _CS_CACHE is None:
        _CS_CACHE = []
        for dp, dn, fn in os.walk(dec):
            dn.sort()
            for f in sorted(fn):
                if f.endswith(".cs"):
                    p = os.path.join(dp, f)
                    with open(p, "r", encoding="utf-8", errors="replace") as fh:
                        text = fh.read().lower()
                    _CS_CACHE.append((os.path.relpath(p, dec), text, text.split("\n")))
    return _CS_CACHE


def find_refs(name, dec, menu_lines):
    low = name.lower()
    m = re.match(r"^(.*)\.(xml|xls|json)$", low)
    truncated = False
    if m:
        stem = m.group(1)
    else:
        truncated = True
        stem = re.sub(r"\.[a-z]{0,3}$", "", low) if "." in low else low
    if not truncated:
        rx = re.compile(re.escape(stem) + r"(?![a-z0-9_])")
        kieu = "ten_tep"
    else:
        rx = re.compile(re.escape(stem))
        kieu = "ten_tep_bi_cat"
    hits = []
    for rel, text, lines in load_cs(dec):
        if stem not in text:
            continue
        for ln_no, ln in enumerate(lines, 1):
            if stem in ln and rx.search(ln):
                hits.append({"file": rel, "line": ln_no, "kieu": kieu})
    if hits:
        return hits
    # tien to dong: "<tien_to>_" + bien
    toks = stem.split("_")
    for k in range(len(toks) - 1, 1, -1):
        prefix = "_".join(toks[:k]) + "_"
        rx2 = re.compile(re.escape(prefix) + r'"|' + re.escape(prefix[:-1]) + r'"')
        for rel, text, lines in load_cs(dec):
            if prefix[:-1] not in text:
                continue
            for ln_no, ln in enumerate(lines, 1):
                if prefix[:-1] in ln and rx2.search(ln):
                    hits.append({"file": rel, "line": ln_no, "kieu": "tien_to_dong:" + prefix})
        if hits:
            return hits
    # catalogue_MMYYYY_NN: ten lay tu Menu.xml (InterfaceIni) qua GetCatalogueFileName
    if re.match(r"^catalogue(qttn)?_\d", stem):
        for rel, text, lines in load_cs(dec):
            for ln_no, ln in enumerate(lines, 1):
                if "public static string[] getcataloguefilename(" in ln:
                    hits.append({"file": rel, "line": ln_no, "kieu": "gian_tiep_menu_xml"})
                    return hits
    return hits


# ---------------------------------------------------------------- chinh


def dump(obj, path):
    s = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(s)
        f.write("\n")


def main(argv):
    args = dict(zip(argv[1::2], argv[2::2]))
    src = args.get("--src", SRC_DEFAULT)
    dec = args.get("--dec", DEC_DEFAULT)
    here = os.path.dirname(os.path.abspath(__file__))
    out = args.get("--out", os.path.join(os.path.dirname(here), "catalog"))
    os.makedirs(out, exist_ok=True)

    inv = inventory(src)
    index = []
    failed = []
    skipped = []
    xml_counts = []  # (name, json_n, src_n)
    xls_counts = []  # (name, json_cells, biff_cells)
    for name, grp, typ in inv:
        path = os.path.join(src, name)
        if typ not in ("ole", "xml", "json"):
            skipped.append((name, typ))
            continue
        with open(path, "rb") as f:
            raw = f.read()
        sha = hashlib.sha256(raw).hexdigest()
        try:
            if typ == "ole":
                doc, st = convert_xls(path)
                xls_counts.append((name, st["so_o"], st["so_o_biff"]))
            elif typ == "xml":
                doc, st = convert_xml(path, raw)
                probe = decode_xml_bytes(raw)[0]
                xml_counts.append((name, st["so_phan_tu"], count_xml_tags(probe)))
            else:
                doc, st = convert_json(path, raw)
        except Exception as e:  # noqa
            failed.append((name, typ, "%s: %s" % (type(e).__name__, e)))
            continue
        header = {"generator": GENERATOR, "version": VERSION, "source_path": os.path.join(src, name),
                  "source_sha256": sha, "encoding": doc.pop("encoding")}
        header.update(doc)
        dump(header, os.path.join(out, name + ".json"))
        refs = find_refs(name, dec, None)
        ent = {"name": name, "json": name + ".json", "loai": header["kind"], "nhom": grp,
               "bytes": len(raw), "source_sha256": sha}
        ent.update(st)
        ent["doc_boi"] = refs[:8]
        ent["so_vi_tri_doc"] = len(refs)
        index.append(ent)
    index.sort(key=lambda e: e["name"])
    idx = {"generator": GENERATOR, "version": VERSION, "source_dir": src, "decompiled_dir": dec,
           "so_tep_kiem_ke": len(inv), "so_tep_chuyen_duoc": len(index),
           "tep_khong_chuyen": [{"name": n, "ly_do": "magic=" + t} for n, t in skipped] +
                               [{"name": n, "ly_do": r} for n, t, r in failed],
           "files": index}
    dump(idx, os.path.join(out, "index.json"))

    # ---- bao cao doi chieu (doc lai tu JSON da ghi)
    bad_xls = []
    tot_json = tot_biff = 0
    for name, j, b in xls_counts:
        # dem lai tu tep JSON da ghi
        with open(os.path.join(out, name + ".json"), encoding="utf-8") as f:
            d = json.load(f)
        jc = sum(len(r["cells"]) for s in d["sheets"] for r in s["rows"])
        tot_json += jc
        tot_biff += b
        if jc != b:
            bad_xls.append((name, jc, b))
    bad_xml = []
    for name, j, s in xml_counts:
        with open(os.path.join(out, name + ".json"), encoding="utf-8") as f:
            d = json.load(f)
        jn = json_stats(d["root"])[0]
        if jn != s:
            bad_xml.append((name, jn, s))
    nokw = [e["name"] for e in index if e["so_vi_tri_doc"] == 0]
    kinds = {}
    for e in index:
        kinds[e["loai"]] = kinds.get(e["loai"], 0) + 1
    print("so tep kiem ke:", len(inv))
    print("so tep chuyen duoc:", len(index), kinds)
    print("tep khong chuyen:", len(skipped) + len(failed), skipped, failed)
    print("XLS: %d tep, o trong JSON = %d, ban ghi o BIFF = %d, so tep lech = %d %s" %
          (len(xls_counts), tot_json, tot_biff, len(bad_xls), bad_xls))
    print("XML: %d tep, so tep lech phan tu = %d %s" % (len(xml_counts), len(bad_xml), bad_xml))
    print("tep khong thay ma doc (%d):" % len(nokw))
    for n in nokw:
        print("  ", n)
    return 0 if not (bad_xls or bad_xml or failed) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
