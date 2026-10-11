# Dựng _spec/validations/*.json từ bản trích thô của roslyn_validations.
# Dùng: python -I build.py <raw.json> <data_root (_forms,_extracted)> <out_dir> <classes.json>
# Chỉ đọc: _spec/csharp/classes.json, _extracted/message.xml, _forms/**/*_mau.xls. Không ghi ngoài out_dir.
import glob
import json
import os
import re
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import biff  # noqa: E402

raw_path, repo, out_dir, classes_path = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
raw = json.load(open(raw_path, encoding="utf-8"))["classes"]
cls_info = {c["class"]: c for c in json.load(open(classes_path, encoding="utf-8"))["classes"]}

# ---- message.xml: lấy bản ghi đầu tiên theo ID (XMLMessage.GetMessage dừng ở lần khớp đầu) ----
xml_msgs = {}
for el in ET.parse(os.path.join(repo, "_extracted/message.xml")).getroot().iter("Infor"):
    xml_msgs.setdefault(el.get("ID"), el.get("Msg"))

# ---- template theo MaFile + VersionXML ----
# menu.xml: MenuClass -> {(MaFile, XsdSchema)}; dùng XsdSchema làm tên dự phòng khi _forms không có thư mục MaFile
menu_pairs = {}
folders_cls = {}      # Class -> {FolderName} (Validity/@FolderName; runtime mở <FolderName>_<VersionXML>.xls)
folders_cls_ver = {}  # (Class, VersionXML) -> {FolderName}
for el in ET.parse(os.path.join(repo, "_extracted/menu.xml")).getroot().iter("Menu"):
    if el.get("Class") and el.get("MaFile"):
        menu_pairs.setdefault(el.get("Class"), set()).add((el.get("MaFile"), el.get("XsdSchema") or ""))
        for v in el.iter("Validity"):
            if v.get("FolderName"):
                folders_cls.setdefault(el.get("Class"), set()).add(v.get("FolderName"))
                folders_cls_ver.setdefault((el.get("Class"), el.get("VersionXML") or ""), set()).add(v.get("FolderName"))

# chỉ mục tên tệp (chữ thường) trong _forms/*/*/ và _extracted/
file_index = {}
for d in sorted(glob.glob(os.path.join(repo, "_forms", "*", "*"))) + [os.path.join(repo, "_extracted")]:
    if os.path.isdir(d):
        for fn in sorted(os.listdir(d)):
            file_index.setdefault(fn.lower(), []).append(os.path.join(d, fn))


def _is_ole(path):
    try:
        with open(path, "rb") as f:
            return f.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    except OSError:
        return False


def _norm2(name):
    return re.sub(r"(^|_)0+(?=[0-9a-z])", r"\1", name.lower())


def template_key(cname):
    """Trả về (danh sách (MaFile, XsdSchema) ứng viên, phiên bản, cách suy)."""
    c = cls_info.get(cname)
    m = re.match(r"^(.*)Nckt(\d{3})$", cname)
    if c and c["in_menu"]:
        xs = sorted(x for (mf, x) in menu_pairs.get(c["MenuClass"], set()) if mf == c["MaFile"]) or [""]
        return [(c["MaFile"], xs[0])], c["VersionXML"], "menu"
    if not m:
        return [], None, "tên class không có hậu tố NcktNNN"
    base, ver = m.group(1), ".".join(m.group(2))
    pairs = sorted(menu_pairs.get(base, set()))
    if not pairs:
        return [], ver, "MenuClass %s không có trong menu" % base
    return pairs, ver, "suy từ MenuClass + hậu tố Nckt"


def _norm(name):
    toks = name.lower().split("_")
    toks = [t.lstrip("0") or "0" if t.isdigit() else t for t in toks]
    s = "_".join(toks)
    return "_" + s if s[:1].isdigit() else s


def find_xls(mafile, xsd, ver, folders=()):
    f = _find_xls_old(mafile, xsd, ver)
    if f:
        return f
    vd = ver.replace(".", "")
    for nm in [mafile, xsd] + sorted(folders):
        if not nm:
            continue
        # tên tệp giải nén từ cab mất ký tự đầu (05QTT_TNCN -> _5qtt_tncn, 334_B01 -> _34_b01)
        for base in sorted({nm.lower(), nm.lower()[1:], _norm(nm).lstrip("_"), _norm2(nm)}):
            for pre in ("", "_"):
                for ext in (".xls", ".xl", ".x", ""):  # tên tệp có thể bị cụt đuôi khi giải nén
                    for path in file_index.get(pre + base + "_" + vd + ext, []):
                        if _is_ole(path) and workbook_path(path):
                            return path
    return None


def _find_xls_old(mafile, xsd, ver):
    cands = [os.path.join(repo, "_forms", "*", mafile, "%s_v%s_mau.xls" % (mafile, ver))]
    if xsd:
        cands.append(os.path.join(repo, "_forms", "*", xsd, "%s_v%s_mau.xls" % (xsd, ver)))
    vd = ver.replace(".", "")
    cands.append(os.path.join(repo, "_extracted", "%s_%s.xls" % (_norm(mafile), vd)))
    if xsd:
        cands.append(os.path.join(repo, "_extracted", "%s_%s.xls" % (_norm(xsd), vd)))
    for pat in cands:
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[0]
    return None


_wb_cache = {}


def workbook_path(path):
    if path in _wb_cache:
        return _wb_cache[path]
    try:
        res = biff.read_workbook(path)
    except Exception:
        res = None
    _wb_cache[path] = res
    return res


def templates_for(cname):
    pairs, ver, how = template_key(cname)
    ci = cls_info.get(cname)
    mc = ci["MenuClass"] if ci and ci["in_menu"] else re.sub(r"Nckt\d{3}$", "", cname)
    folders = folders_cls_ver.get((mc, ver or "")) or folders_cls.get(mc, set())
    out = []
    for mf, xs in pairs:
        f = find_xls(mf, xs, ver, folders)
        if f:
            out.append((mf, f))
    return out, pairs, ver, how


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def resolve_part(p, cname):
    out = {"kind": p["kind"], "id": p.get("id"), "gm_line": p["gm_pos"][0], "api": p.get("api"), "pos": p["gm_pos"]}
    if p.get("reason"):
        out["reason"] = p["reason"]
        out["text"] = None
        return out
    if p["kind"] == "literal":
        out["text"] = p.get("text")
        return out
    if p["kind"] == "xml_id":
        t = xml_msgs.get(p["id"])
        out["text"] = t
        if t is None:
            out["reason"] = "ID không có trong message.xml"
        return out
    if p["kind"] == "cell_text":
        out["sheet"] = p.get("sheet")
        tpls, pairs, ver, how = templates_for(cname)
        if not tpls:
            out["text"] = None
            out["reason"] = ("không có template .xls cho %s v%s" % ("/".join(mf for mf, _ in pairs), ver)) if pairs else "không xác định template: " + how
            return out
        mm = re.match(r"^([A-Z]+)_(\d+)$", p["id"])
        found = []
        for mf, f in tpls:
            sheets = workbook_path(f)
            if not sheets:
                continue
            if p.get("sheet") == "Header":
                sh = next((x for x in sheets if x[0] == "Header"), None)
            else:
                sh = sheets[-1]
            if sh is None:
                continue
            t = sh[1].get((int(mm.group(2)) - 1, col_index(mm.group(1))))
            found.append((os.path.relpath(f, repo), sh[0], t))
        texts = sorted(set(t for _, _, t in found if t and t.strip()))
        if found:
            out["template"] = found[0][0]
            out["template_sheet"] = found[0][1]
        if len(texts) == 1:
            out["text"] = texts[0]
            if len(found) > 1:
                out["template"] = ";".join(sorted(f for f, _, _ in found))
        elif len(texts) > 1:
            out["text"] = None
            out["reason"] = "nhiều template ứng viên cho text khác nhau: " + ";".join(sorted(f for f, _, _ in found))
        else:
            out["text"] = None
            out["reason"] = ("ô %s trống trong sheet %s" % (p["id"], found[0][1])) if found else "không đọc được BIFF"
        return out
    out["text"] = None
    out["reason"] = p.get("reason", "dạng không hỗ trợ")
    return out


SHARED = {"GeneralFunctions", "ValidateBase", "Common"}
unresolved = []
unused_all = {}
per_class = {}
app_classes = {}
totals = {"checks": 0}
for c in raw:
    cname = c["class"]
    is_biz = c["assembly"] == "tax.htkk.business"
    mrecs = []
    for m in c["methods"]:
        checks = []
        for k in m["checks"]:
            msg = k["message"]
            parts = [resolve_part(p, cname) for p in msg.get("parts", [])]
            if cname in SHARED or not is_biz:
                for p in parts:
                    if p["kind"] == "cell_text" and p.get("text") is None:
                        p["reason"] = "hàm dùng chung/ứng dụng: ô thông điệp đọc từ template của tờ khai gọi hàm"
            reasons = []
            if parts:
                first = parts[0]
                mo = {"kind": first["kind"], "id": first.get("id"), "text": first.get("text"), "parts": parts}
                if msg.get("composite"):
                    mo["composite"] = True
                    if msg.get("template"):
                        tpl = msg["template"]
                        mo["template"] = tpl
                        filled = tpl
                        for pp in parts:
                            if pp.get("id") and pp.get("text"):
                                filled = filled.replace("{" + pp["id"] + "}", pp["text"])
                        mo["text"] = filled
                    reasons.append("thông điệp ghép thêm giá trị lúc chạy")
                msg_ok = all(p.get("text") for p in parts)
                for p in parts:
                    if not p.get("text"):
                        reasons.append("%s %s: %s" % (p["kind"], p.get("id"), p.get("reason")))
                        unresolved.append({
                            "class": cname, "method": m["method"], "file": "_decompiled/" + c["file"],
                            "line": p["gm_line"], "kind": p["kind"], "id": p.get("id"), "reason": p.get("reason"),
                        })
            elif msg.get("kind") == "literal":
                mo = {"kind": "literal", "id": None, "text": msg.get("text")}
                msg_ok = True
                if msg.get("reason"):
                    reasons.append(msg["reason"])
            else:
                mo = {"kind": "dynamic", "id": None, "text": None, "expr": msg.get("expr")}
                msg_ok = False
                notes = msg.get("notes", [])
                why = "thông điệp là tham số của hàm" if any(n.startswith("param:") for n in notes) else "thông điệp tính lúc chạy"
                reasons.append(why)
                unresolved.append({"class": cname, "method": m["method"], "file": "_decompiled/" + c["file"],
                                   "line": k["lines"][0], "kind": "dynamic", "id": None, "reason": why + ": " + str(msg.get("expr"))})
            for n in msg.get("notes", []):
                if n.startswith("multi:"):
                    reasons.append("biến %s được gán nhiều lần; lấy phép gán gần nhất phía trước" % n[6:])
            if not k["cell_const"]:
                reasons.append("ô bị đánh dấu phụ thuộc chỉ số hàng/cột lúc chạy")
            if not msg_ok:
                status = "unresolved"
            elif not k["cell_const"] or msg.get("composite") or any(n.startswith("multi:") for n in msg.get("notes", [])):
                status = "parametric"
            else:
                status = "resolved"
            rec = {
                "condition": k["condition"],
                "condition_lines": k.get("condition_lines", []),
                "bindings": k["bindings"],
                "binding_info": k.get("binding_info", {}),
                "cells_marked": k["cells_marked"],
                "level": k["level"],
                "mark_api": k["api"],
                "message": mo,
                "source": {"file": "_decompiled/" + c["file"], "lines": k["lines"]},
                "status": status,
                "reason": "; ".join(reasons) if reasons else None,
            }
            if k.get("context"):
                rec["context"] = k["context"]
            if k.get("sheet"):
                rec["sheet"] = k["sheet"]
            checks.append(rec)
        note = None
        if not checks:
            if m["statement_count"] == 0 or m["body_lines"] <= 2:
                note = "thân hàm rỗng (decompile không giữ thân hoặc hàm trống)"
            else:
                note = "không có lệnh đánh dấu ô hay GetMessage trực tiếp; kiểm tra (nếu có) nằm trong các hàm ở calls"
        mrecs.append({
            "method": m["method"], "kind": m["kind"], "in_scope": m["in_scope"], "reached_from": m["reached_from"],
            "return_type": m["return_type"], "sheetIndex": m["sheetIndex"],
            "source": {"file": "_decompiled/" + c["file"], "lines": [m["start_line"], m["end_line"]]},
            "calls": m["calls"], "checks": checks, "note": note, "binding_method": m.get("binding_method"),
        })
        for u in m.get("unused_messages", []):
            r = resolve_part(u, cname)
            unused_all.setdefault(cname, []).append({"method": m["method"], "variable": u.get("variable"), "line": u.get("line"), "usage": u.get("usage"),
                                                     "kind": r["kind"], "id": r.get("id"), "text": r.get("text"), "api": u.get("api"),
                                                     "file": "_decompiled/" + c["file"]})
    rec_list = app_classes if not is_biz else per_class
    rec_list[cname] = mrecs

os.makedirs(out_dir, exist_ok=True)
os.makedirs(os.path.join(out_dir, "_app"), exist_ok=True)


def dump(path, obj):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(obj, ensure_ascii=False, indent=1, sort_keys=True) + "\n")


index = {"classes": {}, "app_classes": {}}
for cname in sorted(cls_info):
    ci = cls_info[cname]
    tpls, pairs, ver, how = templates_for(cname)
    key = {"Class": cname, "MenuID": ci["MenuID"], "MaFile": ci["MaFile"], "VersionXML": ci["VersionXML"], "in_menu": ci["in_menu"]}
    methods = per_class.get(cname, [])
    dump(os.path.join(out_dir, cname + ".json"), {
        "key": key,
        "template": {"candidates": [mf for mf, _ in pairs], "Version": ver, "how": how, "files": [os.path.relpath(f, repo) for _, f in tpls]},
        "methods": methods,
        "unused_messages": unused_all.get(cname, []),
    })
    lv = {}
    st = {}
    for m in methods:
        for k in m["checks"]:
            lv[k["level"]] = lv.get(k["level"], 0) + 1
            st[k["status"]] = st.get(k["status"], 0) + 1
    index["classes"][cname] = {"file": cname + ".json", "in_menu": ci["in_menu"], "MenuID": ci["MenuID"], "MaFile": ci["MaFile"],
                               "VersionXML": ci["VersionXML"], "methods": len(methods),
                               "checks": sum(len(m["checks"]) for m in methods), "by_level": lv, "by_status": st,
                               "template_files": [os.path.relpath(f, repo) for _, f in tpls]}
missing = sorted(set(per_class) - set(cls_info))
for cname in missing:  # class trong thư mục business nhưng ngoài namespace chính (nếu có)
    dump(os.path.join(out_dir, cname + ".json"), {"key": {"Class": cname, "MenuID": None, "MaFile": None, "VersionXML": None, "in_menu": False}, "template": None, "methods": per_class[cname]})
    index["classes"][cname] = {"file": cname + ".json", "in_menu": False, "note": "không có trong classes.json"}
for cname in sorted(app_classes):
    dump(os.path.join(out_dir, "_app", cname + ".json"), {"key": {"Class": cname, "assembly": "htkk"}, "methods": app_classes[cname], "unused_messages": unused_all.get(cname, [])})
    index["app_classes"][cname] = {"file": "_app/" + cname + ".json", "methods": len(app_classes[cname]),
                                   "checks": sum(len(m["checks"]) for m in app_classes[cname])}
index["legend"] = {
    "condition": "cú pháp trung tính: AND/OR/NOT/XOR, =, <>, IF(c,a,b); text(L33)/value(L33) = nội dung ô (cột chữ, hàng đánh số từ 1); "
                 "S1!L33 = ô ở sheet chỉ số 1; {expr +1} = hàng tính lúc chạy (expr là chỉ số 0); col{expr} = cột tính lúc chạy; "
                 "FOR v FROM a WHILE c = vòng lặp; CASE x IN {...} = nhánh switch; message(...) = thông điệp; cellid(E_15) = địa chỉ ô; "
                 "tên hàm khác giữ tên gốc dạng f(args); bindings = biến cục bộ chưa thay vào điều kiện (@L = dòng gán)",
    "pink_count_definition": "level=pink gồm mọi lời gọi SetNodeBackGroundPink: HTKKFile.SetNodeBackGroundPink/4, HTKKFile.SetNodeBackGroundPink/5, GeneralFunctions.SetNodeBackGroundPink/6, cộng HTKKFile.SetNodeBackGroundPink2/5 (đếm riêng theo mark_api); số tổng 7.246 gồm cả lời gọi trong app htkk/ (thư mục _app/); không tính 2 dòng khai báo hàm",
    "unused_messages": "GetMessage không dẫn tới kiểm tra: biến không được dùng, hoặc chỉ đưa vào tập hợp (Add/Insert); không tính là check, liệt kê riêng theo class",
    "source.lines": "dòng của lệnh đánh dấu ô (SetNodeBackGround*/gán BackColor), của lời gọi GetMessage, hoặc của frmMessage; KHÔNG phải dòng câu if",
    "condition_lines": "dòng nguồn của từng vế trong condition, cùng thứ tự (vế ngoài cùng trước); vế 'via helper' lấy dòng lời gọi helper",
    "bindings": "biến chưa thay vào điều kiện; nếu có nhiều phép gán có thể tới điểm dùng (reaching definitions trên ControlFlowGraph, theo ISymbol) thì liệt kê tất cả dạng 'L<dòng> [khi <điều kiện riêng của phép gán>]: <giá trị hoặc +=/-=/++ biểu thức>'; phép gán +=/++ kèm các phép gán tới được nó; khóa 'tên@L<dòng khai báo>' khi một check có hai biến cùng tên khác phạm vi",
    "binding_info": "theo từng khóa binding: method (cfg = reaching definitions trên ControlFlowGraph; fallback = phân tích cú pháp cũ), decl_line (dòng khai báo ISymbol), use_line (dòng điểm dùng nơi tính binding), def_lines (dòng các phép gán tới được)",
    "binding_method": "theo method: cfg nếu dựng được ControlFlowGraph, fallback nếu không",
    "sheet": "sheet của ô bị đánh dấu, trùng tiền tố trong cells_marked: S<n> | S[last] (sheet cuối) | S[<tham số>] khi sheet là tham số của hàm dùng chung và tên hàm không cho chỉ số | S[active] | S[<biểu thức>]",
    "ký hiệu trung tính": "sheet(S<n>) = đối tượng sheet; workbook = bảng tính (FpSpread); sheet_count() = số sheet; visible(x) = x đang hiển thị; now()/today()/date_min()/date_max(); color(pink); key_press; col_index(\"AB\") = chỉ số cột (1 = A); model.old(...) = dữ liệu tờ khai kỳ trước (tKhaiInfoOld); create(Type, ...) = tạo đối tượng",
    "message.template": "thông điệp ghép: {E_xx}/{ID} là phần thông điệp đã tra (thay vào message.text), {biểu thức} là phần tính lúc chạy",
    "via helper": "condition bắt đầu bằng 'via helper X' = thông điệp truyền vào hàm X, ô và điều kiện thật nằm trong X",
    "level": "pink = SetNodeBackGroundPink/Pink2 (lỗi chặn); yellow = SetNodeBackGroundYellow (cảnh báo); other = GetMessage không gắn ô hoặc gán BackColor trực tiếp",
    "message.kind": "xml_id = ID trong _extracted/message.xml; cell_text = text ô ẩn trong template (.xls cùng MaFile+phiên bản, sheet cuối = Header); literal; dynamic",
    "status": "resolved = có text thông điệp và ô hằng; parametric = có text nhưng ô/thông điệp phụ thuộc giá trị lúc chạy; unresolved = chưa có text thông điệp",
}
dump(os.path.join(out_dir, "index.json"), index)
unresolved.sort(key=lambda u: (u["file"], u["line"], u["kind"], str(u["id"])))
dump(os.path.join(out_dir, "messages_unresolved.json"), {"count": len(unresolved), "items": unresolved})
print("classes", len(index["classes"]), "app_classes", len(index["app_classes"]), "unresolved", len(unresolved))
