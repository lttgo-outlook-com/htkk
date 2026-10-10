#!/usr/bin/env python3
"""Chuyển các tệp XML cấu hình của HTKK 5.7.9 thành JSON có cấu trúc.

Chạy:  python3 -I _spec/tools/config_xml.py [--src DIR] [--cs DIR] [--out DIR]
Không chạy mã của nguồn. Parse bằng defusedxml (cấm DTD/entity).
Ra: <out>/<tên>.json cho mỗi tệp cấu hình và <out>/index.json.
Kết quả xác định (không có thời gian, sắp xếp ổn định).
"""
import collections
import hashlib
import json
import os
import re
import sys
import xml.etree.ElementTree as StdET

from defusedxml import ElementTree as DET

GENERATOR = "config_xml.py"
VERSION = "1.0"
SRC_DEFAULT = "/Users/tuanle/codes/HTKK/_extracted"
CS_DEFAULT = "/Users/tuanle/codes/HTKK/_decompiled"
HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DEFAULT = os.path.join(os.path.dirname(HERE), "config")

# Vai trò từng tệp cấu hình (mô tả ngắn, suy từ nội dung tệp).
ROLES = {
    "app.manifest": "Manifest Windows (assembly, quyền thực thi) của HTKK.exe",
    "appscheduler.xml": "Cấu hình tự cập nhật: URL ServerBalancer, proxy, chu kỳ, thời điểm chạy",
    "button.xml": "Bật/tắt nút thanh công cụ (Thêm phụ lục, Xóa, In, Kết xuất, Nhập từ XML...) theo ID tờ khai",
    "cell_dynamic_format.xml": "Định dạng ô cho bảng động theo tờ khai/sheet (nền 233-245-254, trắng, không trắng)",
    "config.xml": "Thiết lập người dùng cục bộ: đăng ký barcode, debug chuyển đổi, cài mới, kích thước cửa sổ",
    "config_header.xml": "Ánh xạ ô của phần tiêu đề tờ khai tới thuộc tính thông tin (PropertyInfo)",
    "configprinter.xml": "Cấu hình in tờ khai: ô neo, số dòng trên/dưới, ngắt trang theo Menu/Sheet",
    "converdatafile.xml": "Map tên tệp dữ liệu gói cũ sang gói mới (và ô cũ/mới) khi chuyển đổi",
    "convertdatacommon.xml": "Quy tắc chuyển đổi dữ liệu chung giữa phiên bản (bảng động, đường dẫn bảng/tiêu đề)",
    "convertversion.xml": "Danh sách chuyển phiên bản XML theo tệp dữ liệu (VersionXmlOld/New, ConvertType)",
    "createdatafilekhbsmiss.xml": "Tạo tệp dữ liệu KHBS còn thiếu cho từng tờ khai (PL_01_KHBS, VersionXML, VersionKHBS)",
    "cvanconfig.xml": "Cấu hình cổng kết nối C-VAN/BHXH (IP, cổng, tên dịch vụ; giá trị đã mã hóa)",
    "dmngayle.xml": "Danh mục ngày lễ: hạn nộp gốc và ngày hiệu lực dời theo kỳ",
    "header.xml": "Mẫu tiêu đề tờ khai (Sections/Cells NCKT-1.0), bản có thuộc tính hiển thị",
    "header.xml9": "Mẫu tiêu đề tờ khai (NCKT-1.0), bản rút gọn; tên tệp bị cắt đuôi",
    "hieulucbanggia.xml": "Hiệu lực danh mục bảng giá: Map ID -> Validity -> Sheet (mẫu giao diện, lớp, tệp dữ liệu)",
    "htkk.autoupate.exe.config": "Cấu hình .NET của tiến trình tự cập nhật (appSettings)",
    "htkk.exe.config": "Cấu hình .NET của HTKK.exe (appSettings, nhà cung cấp dữ liệu)",
    "htkk.vshost.exe.config": "Cấu hình .NET của vshost (môi trường Visual Studio)",
    "htkk.vshost.exe.manifest": "Manifest Windows của vshost",
    "map.xml": "Bảng ánh xạ chính tờ khai: Map ID, kỳ kê khai, Validity (ngày hiệu lực, DataFile, thuế khấu trừ/phải nộp)",
    "mapcongthucchosua.xml": "Công thức cho dòng động (ROWREPLACE) theo mã tờ khai: Item gồm SheetIndex, MCTPath, TablePath, Formula",
    "mapmafile.xml": "Ánh xạ MapID sang MaFile, phiên bản XML/KHBS, mã tờ khai (MaTK)",
    "mapmct.xml": "Ánh xạ ô (CellID) sang chỉ tiêu (MCT) theo tờ khai, kèm cờ điều chỉnh tăng và chú thích",
    "mapmct_tieumuc.xml": "Ánh xạ chỉ tiêu (MCT) với tiểu mục, đường dẫn bảng/ô mã tiểu mục",
    "mapmct_tm.xml": "Ánh xạ ô sang chỉ tiêu (MCT) cho tiểu mục, tương tự mapmct.xml (tập con)",
    "maptaidulieulon.xml": "Cấu hình tải dữ liệu lớn: cột công thức và cột viền phải theo Map ID/SheetIndex",
    "maptentk.xml": "Tên tờ khai (TenTK) và mô tả biểu mẫu (MoTaBMau, căn cứ thông tư) theo MapID",
    "menu.xml": "Cây menu tờ khai: Menu (ID, PopID, ParentID, MaFile, Class, KieuKy...) -> Validity -> Sheet",
    "menu_backup161118.xml": "Bản sao lưu menu.xml (ngày 16/11/2018 theo tên tệp)",
    "menu_backup2009.xml": "Bản sao lưu menu.xml (2009 theo tên tệp)",
    "menu_backup2011.xml": "Bản sao lưu menu (2011 theo tên tệp); toàn bộ nội dung bị đóng trong chú thích, không có phần tử",
    "message.xml": "Thông báo hệ thống: Infor (ID, Msg) cho hộp thoại, cảnh báo, lỗi",
    "network.xml": "Cấu hình proxy mạng (IP, cổng, tài khoản) của máy chủ",
    "replacedatafile.xml": "Quy tắc thay thế dữ liệu bảng khi chuyển phiên bản (TxtOld -> TxtNew theo TablePath)",
    "schema.dtd": "DTD (không phải XML) mô tả Cell/Section/Sections của mẫu NCKT",
    "schema.dtd1": "Bản sao schema.dtd; tên tệp bị cắt đuôi",
    "ttinchung.xml": "Khuôn TTinChung (thông tin chung tờ khai) dùng làm mẫu khi kết xuất XML",
    "validcombobox.xml": "Quy tắc kiểm tra giá trị combobox (ID ô, bảng động, TablePath, CtPath, CtPathMa)",
    "versiontokhai.xml": "STUB: Root chỉ có chú thích, gần như không có phần tử; không có cấu hình hiệu lực",
}

NOTES = {
    "versiontokhai.xml": "Tệp stub (loại khỏi danh mục cấu hình hiệu lực): Root có 0 phần tử con, chỉ chứa 1 chú thích Menu đã vô hiệu.",
    "menu_backup2011.xml": "Không parse được như XML (chỉ có chú thích); lưu dạng raw_text.",
    "schema.dtd": "Không phải XML (DTD); lưu dạng raw_text. Hai vị trí .cs chỉ là chuỗi hằng DOCTYPE chứa 'Schema.dtd' (chuoi_hang), không đọc tệp.",
    "header.xml9": "Tên bị cắt đuôi; tương ứng Header.xml trong thư mục dữ liệu từng MST (cùng vị trí đọc với header.xml, ref_kind gian_tiep, không tính vào cs_ref_count).",
    "htkk.exe.config": "Runtime .NET nạp; mã đọc giá trị qua ConfigurationManager.AppSettings (ref_kind config_manager) trong HTKKConfig.cs.",
    "schema.dtd1": "Không phải XML (DTD); lưu dạng raw_text.",
    "network.xml": "Chứa tài khoản/mật khẩu proxy dạng rõ trong nguồn; sao chép nguyên văn.",
    "cvanconfig.xml": "Giá trị IP/Port/ServiceName là chuỗi mã hóa (base64) trong nguồn; sao chép nguyên văn.",
    "appscheduler.xml": "Mã đọc dưới tên AppSchedulerCf.xml (ref_kind ten_khac). Có URL máy chủ cập nhật và thông tin proxy trong nguồn; sao chép nguyên văn.",
}

# Nhóm loại khỏi danh mục cấu hình (không chuyển) kèm lý do.
LIBRARY_DOC = re.compile(r"^(farpoint\.|grapecity\.|office\.xml$)")


def classify(name):
    """Trả (nhóm, lý do) cho tệp có nội dung giống XML."""
    low = name.lower()
    if low.endswith(".xsd"):
        return "xsd", "lược đồ XSD"
    if low.startswith("catalogue_") or low.startswith("catalogueqttn_"):
        return "catalogue", "danh mục catalogue (catalogue_/catalogueqttn_)"
    if low.startswith("dm_"):
        return "dm", "danh mục dm_"
    if LIBRARY_DOC.match(low):
        return "thu_vien", "tài liệu XML của thư viện bên thứ ba (FarPoint/GrapeCity/Office interop), không phải cấu hình HTKK"
    if low.startswith("in_bia_"):
        return "mau_tk", "mẫu in bìa tờ khai"
    if (re.match(r"^_\d", low) or low.startswith(("khbs_", "bc", "giay_", "bk_", "ds_", "dk_"))):
        return "mau_tk", "mẫu tờ khai/dữ liệu tờ khai"
    return "cau_hinh", ""


def decode_head(b):
    if b.startswith((b"\xff\xfe", b"\xfe\xff")):
        return b.decode("utf-16", errors="replace")
    if b.startswith(b"\xef\xbb\xbf"):
        b = b[3:]
    return b.decode("utf-8", errors="replace")


def looks_xml(path):
    with open(path, "rb") as f:
        head = f.read(400)
    return decode_head(head).lstrip().startswith("<")


def detect_encoding(b):
    if b.startswith(b"\xff\xfe"):
        return "utf-16-le (BOM)"
    if b.startswith(b"\xfe\xff"):
        return "utf-16-be (BOM)"
    if b.startswith(b"\xef\xbb\xbf"):
        return "utf-8 (BOM)"
    m = re.match(rb'\s*<\?xml[^>]*encoding=["\']([^"\']+)["\']', b[:200])
    return (m.group(1).decode("ascii").lower() + " (khai báo)") if m else "utf-8 (mặc định)"


def decode_text(b):
    if b.startswith((b"\xff\xfe", b"\xfe\xff")):
        return b.decode("utf-16")
    if b.startswith(b"\xef\xbb\xbf"):
        b = b[3:]
    return b.decode("utf-8", errors="replace")


def parse_xml(b):
    """Parse an toàn, giữ chú thích. defusedxml cấm DTD/entity."""
    parser = DET.XMLParser(target=StdET.TreeBuilder(insert_comments=True))
    parser.feed(b)
    return parser.close()


def is_comment(e):
    return not isinstance(e.tag, str)


def ws_only(s):
    return s is None or s.strip() == ""


def to_node(e):
    n = {}
    if is_comment(e):
        n["comment"] = e.text or ""
    else:
        n["tag"] = e.tag
        if e.attrib:
            n["attrs"] = dict(e.attrib)
        if not ws_only(e.text):
            n["text"] = e.text
        kids = [to_node(c) for c in e]
        if kids:
            n["children"] = kids
    if not ws_only(e.tail):
        n["tail"] = e.tail
    return n


def from_node(n, parent=None):
    if "comment" in n:
        e = StdET.Comment(n["comment"])
        if parent is not None:
            parent.append(e)
    else:
        e = StdET.Element(n["tag"], n.get("attrs", {})) if parent is None else StdET.SubElement(parent, n["tag"], n.get("attrs", {}))
        e.text = n.get("text")
        for c in n.get("children", []):
            from_node(c, e)
    e.tail = n.get("tail")
    return e


def compare(a, b, path, diffs):
    """So hai cây: thẻ, thuộc tính (cả thứ tự), text, tail, con. Text chỉ-khoảng-trắng coi như rỗng."""
    ta = a.tag if isinstance(a.tag, str) else "#comment"
    tb = b.tag if isinstance(b.tag, str) else "#comment"
    here = f"{path}/{ta}"
    if ta != tb:
        diffs.append(f"{here}: tag {ta} != {tb}")
        return
    if list(a.attrib.items()) != list(b.attrib.items()):
        diffs.append(f"{here}: attrs")
    na = None if ws_only(a.text) else a.text
    nb = None if ws_only(b.text) else b.text
    if na != nb:
        diffs.append(f"{here}: text")
    ta_ = None if ws_only(a.tail) else a.tail
    tb_ = None if ws_only(b.tail) else b.tail
    if ta_ != tb_:
        diffs.append(f"{here}: tail")
    ca, cb = list(a), list(b)
    if len(ca) != len(cb):
        diffs.append(f"{here}: số con {len(ca)} != {len(cb)}")
        return
    for i, (x, y) in enumerate(zip(ca, cb)):
        compare(x, y, f"{here}[{i}]", diffs)


def stats(root):
    c = collections.Counter(e.tag for e in root.iter() if isinstance(e.tag, str))
    comments = sum(1 for e in root.iter() if not isinstance(e.tag, str))
    return c, comments


READ_KINDS = ("doc_theo_ten", "ten_khac", "config_manager")  # chỉ các loại này tính vào cs_ref_count
CONST_REFS = {"schema.dtd", "schema.dtd1"}
_F = "htkk/HTKK.Forms/frmFormStart.cs"
_U = "htkk/HTKK/frmConfigUpdate.cs"
EXTRA_REFS = {
    # đọc dưới tên khác
    "appscheduler.xml": [(f"{_F}:210", "ten_khac"), (f"{_F}:298", "ten_khac"), (f"{_F}:420", "ten_khac"),
                         (f"{_U}:133", "ten_khac"), (f"{_U}:163", "ten_khac")],
}
CFG_FILE = "tax.htkk.utilities/TAX.HTKK.Utilities.Core/HTKKConfig.cs"


def config_manager_lines(cs_root):
    """Dòng ConfigurationManager.AppSettings trong HTKKConfig.cs (đọc htkk.exe.config qua runtime)."""
    with open(os.path.join(cs_root, CFG_FILE), "rb") as f:
        lines = f.read().decode("utf-8", errors="replace").splitlines()
    return [f"{CFG_FILE}:{i}" for i, l in enumerate(lines, 1) if "ConfigurationManager.AppSettings" in l]


def find_cs_refs(cs_root, names):
    """Tìm tên tệp trong mọi .cs của _decompiled. Khớp chính xác tên tệp (không dính ký tự chữ/số/_ hai bên)."""
    files = []
    for dp, dn, fn in os.walk(cs_root):
        dn.sort()
        for f in sorted(fn):
            if f.endswith(".cs"):
                files.append(os.path.join(dp, f))
    pats = {n: re.compile(r"(?<![A-Za-z0-9_])" + re.escape(n) + r"(?![A-Za-z0-9_])", re.I) for n in names}
    stems = {}
    for n in names:
        stem = n.rsplit(".", 1)[0] if "." in n else n
        if len(stem) >= 8 or "_" in stem:
            stems[n] = re.compile(r'"[^"\n]*(?<![A-Za-z0-9_])' + re.escape(stem) + r'(?![A-Za-z0-9_])[^"\n]*"', re.I)
    exact = {n: [] for n in names}
    indirect = {n: [] for n in names}
    for p in files:
        rel = os.path.relpath(p, cs_root)
        with open(p, "rb") as f:
            lines = f.read().decode("utf-8", errors="replace").splitlines()
        for i, line in enumerate(lines, 1):
            for n in names:
                if pats[n].search(line):
                    exact[n].append(f"{rel}:{i}")
                elif n in stems and stems[n].search(line):
                    indirect[n].append(f"{rel}:{i}")
    return exact, indirect, len(files)


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC_DEFAULT)
    ap.add_argument("--cs", default=CS_DEFAULT)
    ap.add_argument("--out", default=OUT_DEFAULT)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)

    # 1. Kiểm kê toàn bộ tệp giống XML
    inv = collections.Counter()
    cfg = []
    excluded = []
    total_files = 0
    for name in sorted(os.listdir(a.src)):
        p = os.path.join(a.src, name)
        if os.path.isdir(p):
            continue
        total_files += 1
        if name.lower().endswith((".dll", ".png", ".gif", ".jpg", ".rar", ".exe")):
            continue
        if not looks_xml(p) and not name.lower().endswith(".xsd"):
            continue
        grp, why = classify(name)
        inv[grp] += 1
        if grp == "cau_hinh":
            cfg.append(name)
        elif grp in ("thu_vien", "mau_tk") and not re.match(r"^_\d|^khbs_|^bc|^giay_|^bk_|^ds_|^dk_", name.lower()):
            excluded.append({"name": name, "group": grp, "reason": why,
                             "size": os.path.getsize(p)})
        elif grp == "catalogue" and name.lower().startswith("catalogueqttn_"):
            excluded.append({"name": name, "group": grp, "reason": why, "size": os.path.getsize(p)})
    print(f"Tệp trong nguồn: {total_files}; tệp giống XML phân loại: {sum(inv.values())}")
    for k, v in sorted(inv.items()):
        print(f"  {k}: {v}")
    print(f"Tệp cấu hình kiểm kê: {len(cfg)}")
    missing_role = [n for n in cfg if n not in ROLES]
    assert not missing_role, f"thiếu mô tả vai trò: {missing_role}"

    # 2. Chuyển đổi + kiểm vòng
    exact, indirect, ncs = find_cs_refs(a.cs, cfg)
    index = []
    converted = 0
    roundtrip_diffs = 0
    counts = {}
    for name in cfg:
        p = os.path.join(a.src, name)
        raw = open(p, "rb").read()
        sha = hashlib.sha256(raw).hexdigest()
        enc = detect_encoding(raw)
        header = {"generator": GENERATOR, "version": VERSION, "source_path": p,
                  "source_sha256": sha, "encoding": enc,
                  "whitespace_only_text": "bỏ (chỉ lưu text/tail có ký tự không phải khoảng trắng)"}
        entry = {"name": name, "json": name + ".json", "size_bytes": len(raw), "source_sha256": sha,
                 "encoding": enc, "role": ROLES[name]}
        if name in NOTES:
            entry["note"] = NOTES[name]
        try:
            if name.lower().endswith(".dtd") or name.lower().endswith(".dtd1"):
                raise ValueError("DTD")
            root = parse_xml(raw)
            doc = {"header": header, "format": "xml", "root": to_node(root)}
            # kiểm vòng: dựng lại XML từ JSON, parse lại, so với nguồn
            rebuilt = from_node(json.loads(json.dumps(doc, ensure_ascii=False))["root"])
            xml_bytes = StdET.tostring(rebuilt, encoding="utf-8")
            reparsed = parse_xml(xml_bytes)
            diffs = []
            compare(root, reparsed, "", diffs)
            c, ncom = stats(root)
            entry.update({"format": "xml", "root_tag": root.tag,
                          "root_children": len(list(root)), "element_count": sum(c.values()),
                          "comment_count": ncom, "tag_counts": dict(c.most_common(6))})
            counts[name] = c
        except Exception as ex:  # không phải XML hợp lệ: lưu nguyên văn
            text = decode_text(raw)
            doc = {"header": header, "format": "raw_text", "parse_error": str(ex), "text": text}
            diffs = [] if decode_text(raw) == doc["text"] else ["raw_text"]
            entry.update({"format": "raw_text", "root_tag": None, "root_children": 0,
                          "element_count": 0, "comment_count": 0, "tag_counts": {}})
        roundtrip_diffs += len(diffs)
        entry["roundtrip_diffs"] = len(diffs)
        with open(os.path.join(a.out, name + ".json"), "w", encoding="utf-8", newline="\n") as f:
            json.dump(doc, f, ensure_ascii=False, indent=1)
            f.write("\n")
        converted += 1
        refs = [{"loc": r, "ref_kind": "doc_theo_ten"} for r in exact[name]]
        if name in EXTRA_REFS:
            for r, kind in EXTRA_REFS[name]:
                refs.append({"loc": r, "ref_kind": kind})
        if name in CONST_REFS:  # khớp tên nhưng chỉ là chuỗi hằng, không đọc tệp
            for r in refs:
                if r["ref_kind"] == "doc_theo_ten":
                    r["ref_kind"] = "chuoi_hang"
        if name == "htkk.exe.config":
            refs += [{"loc": r, "ref_kind": "config_manager"} for r in config_manager_lines(a.cs)]
        if name == "header.xml9":
            refs += [{"loc": r, "ref_kind": "gian_tiep"} for r in exact["header.xml"]]
        if indirect[name] and not exact[name]:
            refs += [{"loc": r, "ref_kind": "gian_tiep"} for r in indirect[name]]
        entry["cs_refs"] = refs
        entry["cs_ref_count"] = sum(1 for r in refs if r["ref_kind"] in READ_KINDS)
        index.append(entry)

    idx = {"header": {"generator": GENERATOR, "version": VERSION, "source_dir": a.src, "cs_dir": a.cs,
                      "cs_files_scanned": ncs},
           "inventory": {"files_in_source": total_files, "xml_like_by_group": dict(sorted(inv.items())),
                         "config_files": len(cfg), "converted": converted},
           "excluded_non_config": sorted(excluded, key=lambda x: x["name"]),
           "files": index}
    with open(os.path.join(a.out, "index.json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)
        f.write("\n")

    # 3. Báo cáo
    print(f"Đã chuyển: {converted}/{len(cfg)}; lệch kiểm vòng: {roundtrip_diffs}")
    for n, tag in (("mapcongthucchosua.xml", "Item"), ("maptaidulieulon.xml", "Item"),
                   ("validcombobox.xml", "Item"), ("message.xml", "Infor")):
        raw = open(os.path.join(a.src, n), "rb").read()
        txt = decode_text(raw)
        by_regex = len(re.findall(r"<" + tag + r"(?=[\s/>])", txt))
        by_regex_prefix = len(re.findall(r"<" + tag, txt))
        print(f"  {n}: {counts[n][tag]} {tag} (phần tử parse); regex '<{tag}' thô = {by_regex_prefix}; regex có ranh giới = {by_regex}")
    with_cs = [e["name"] for e in index if e["cs_ref_count"] > 0]
    no_cs = [e["name"] for e in index if e["cs_ref_count"] == 0]
    print(f"Tệp có >=1 vị trí .cs đọc (khớp tên đầy đủ): {len(with_cs)}/{len(index)}")
    print(f"Tệp không thấy mã đọc ({len(no_cs)}): " + ", ".join(no_cs))
    ind = [e["name"] for e in index if e.get("cs_refs_indirect_stem")]
    print(f"Trong đó có khớp gián tiếp theo tên không đuôi ({len(ind)}): " + ", ".join(ind))
    print(f"Tệp loại khỏi danh mục cấu hình đã ghi trong index: {len(excluded)}")


if __name__ == "__main__":
    sys.exit(main())
