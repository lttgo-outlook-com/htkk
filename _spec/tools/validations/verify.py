# Kiểm tra tiêu chí cho _spec/validations. Dùng: python -I verify.py <data_root> <spec_root> [--sample]
import collections
import glob
import json
import os
import random
import re
import sys
import xml.etree.ElementTree as ET

data, spec = sys.argv[1], sys.argv[2]
vdir = os.path.join(spec, "validations")
recs = {}
for f in sorted(glob.glob(os.path.join(vdir, "*.json")) + glob.glob(os.path.join(vdir, "_app", "*.json"))):
    b = os.path.basename(f)
    if b in ("index.json", "messages_unresolved.json"):
        continue
    recs[os.path.relpath(f, vdir)] = json.load(open(f, encoding="utf-8"))
methods = [(k, m) for k, r in recs.items() for m in r["methods"]]
checks = [(k, m, c) for k, m in methods for c in m["checks"]]
print("== (0) tệp: %d class nghiệp vụ + %d class ứng dụng; method: %d; check: %d" % (
    sum(1 for k in recs if not k.startswith("_app")), sum(1 for k in recs if k.startswith("_app")), len(methods), len(checks)))

# (1)
vn = sum(1 for _, m in methods if re.match(r"^Validate_\d+$", m["method"]))
va = sum(1 for _, m in methods if m["method"].startswith("Validate"))
idx = json.load(open(os.path.join(spec, "csharp/index.json"), encoding="utf-8"))
ivn = sum(1 for f in idx["files"] for m in f["methods"] if re.match(r"^Validate_\d+$", m["name"]))
iva = sum(1 for f in idx["files"] for m in f["methods"] if m["name"].startswith("Validate"))
print("== (1) Validate_N trong JSON = %d (chỉ mục T0: %d); Validate* = %d (T0: %d)" % (vn, ivn, va, iva))
print("   method theo kind:", dict(sorted(collections.Counter(m["kind"] for _, m in methods).items())))
print("   method Validate*/isValid/checkPL không có check:", sum(1 for _, m in methods if not m["checks"] and m["kind"] in ("Validate", "isValidDuLieuTK", "checkPL")),
      "; trong đó thân rỗng:", sum(1 for _, m in methods if not m["checks"] and (m["note"] or "").startswith("thân hàm rỗng")))

# (2) GetMessage
gm_json = collections.Counter()
gm_kind = collections.Counter()
seen = set()
for k, m, c in checks:
    for p in c["message"].get("parts", []):
        if p.get("api") != "HTKKFile.GetMessage":
            continue
        key = (c["source"]["file"], p["pos"][0], p["pos"][1])
        if key in seen:
            continue
        seen.add(key)
        gm_json[(c["source"]["file"], p["pos"][0])] += 1
        gm_kind[(p["kind"], "có text" if p.get("text") else "không text", c["status"])] += 1
n_unused = 0
for k, r in recs.items():
    for u in r.get("unused_messages", []):
        if u.get("api") == "HTKKFile.GetMessage":
            gm_json[(u["file"], u["line"])] += 1
            n_unused += 1
print("   GetMessage gán biến không dùng (unused_messages, không tính là check):", n_unused)
print("   check có condition 'via helper':", sum(1 for _, _, c in checks if c["condition"].startswith("via helper")))
gm_src = collections.Counter()
by_asm = collections.Counter()
for root in ("_decompiled/tax.htkk.business", "_decompiled/htkk"):
    for f in sorted(glob.glob(os.path.join(data, root, "**", "*.cs"), recursive=True)):
        rel = os.path.relpath(f, data)
        for ln, t in enumerate(open(f, encoding="utf-8-sig"), 1):
            n = t.count("HTKKFile.GetMessage(")
            if n:
                gm_src[(rel, ln)] += n
print("== (2) HTKKFile.GetMessage: nguồn (đếm chuỗi) = %d; JSON (vị trí duy nhất) = %d; dòng lệch = %d" % (
    sum(gm_src.values()), sum(gm_json.values()), sum(1 for k in set(gm_src) | set(gm_json) if gm_src[k] != gm_json[k])))
for k, v in sorted(gm_kind.items()):
    print("   kind=%s %s status=%s: %d" % (k[0], k[1], k[2], v))
st = collections.Counter(c["status"] for _, _, c in checks)
print("   status mọi check:", dict(sorted(st.items())))
print("   message.kind mọi check:", dict(sorted(collections.Counter(c["message"]["kind"] for _, _, c in checks).items())))

# (3) đánh dấu
api = collections.Counter(c["mark_api"] for _, _, c in checks)
print("== (3) mark_api trong JSON:", dict(sorted(api.items())))
print("   level:", dict(sorted(collections.Counter(c["level"] for _, _, c in checks).items())))
src = collections.Counter()
by_asm = collections.Counter()
for root in ("_decompiled/tax.htkk.business", "_decompiled/htkk"):
    for f in sorted(glob.glob(os.path.join(data, root, "**", "*.cs"), recursive=True)):
        for t in open(f, encoding="utf-8-sig"):
            for nm in ("SetNodeBackGroundPink2", "SetNodeBackGroundPink", "SetNodeBackGroundYellow"):
                for mm in re.finditer(r"(void\s+)?\b" + nm + r"\(", t):
                    src[(nm, "khai báo" if mm.group(1) else "lời gọi")] += 1
                    by_asm[(root.split("/")[1], nm, "khai báo" if mm.group(1) else "lời gọi")] += 1
print("   đếm độc lập (regex trên nguồn, business+htkk):", dict(sorted(src.items())))
print("   tách theo assembly:", dict(sorted(by_asm.items())))
jp = sum(v for k, v in api.items() if "SetNodeBackGroundPink/" in k)
jp2 = sum(v for k, v in api.items() if "Pink2" in k)
jy = sum(v for k, v in api.items() if "Yellow" in k)
print("   so khớp lời gọi: Pink JSON=%d nguồn=%d; Pink2 JSON=%d nguồn=%d; Yellow JSON=%d nguồn=%d" % (
    jp, src[("SetNodeBackGroundPink", "lời gọi")], jp2, src[("SetNodeBackGroundPink2", "lời gọi")], jy, src[("SetNodeBackGroundYellow", "lời gọi")]))
neutral_bad = [c for _, _, c in checks if re.search(r"expr\(|lambda\(\.\.\.\)| IS |\?|format\(\"|\bnew ", c["condition"] + json.dumps(c["bindings"], ensure_ascii=False))]
print("   điều kiện/bindings còn mảnh chưa trung tính hoàn toàn (expr(, lambda, IS, ?, format(, new):", len(neutral_bad), "/", len(checks))

# (4) message id
xml_ids = {e.get("ID") for e in ET.parse(os.path.join(data, "_extracted/message.xml")).getroot().iter("Infor")}
xid = collections.Counter()
missing = set()
for k, m, c in checks:
    for p in c["message"].get("parts", []):
        if p["kind"] == "xml_id":
            xid[p.get("id")] += 1
            if p.get("id") not in xml_ids:
                missing.add(p.get("id"))
print("== (4) xml_id: %d lượt, %d ID khác nhau; ID không có trong message.xml: %s" % (sum(xid.values()), len(xid), sorted(map(str, missing))))
ct = collections.Counter()
for k, m, c in checks:
    for p in c["message"].get("parts", []):
        if p["kind"] == "cell_text":
            ct["có text" if p.get("text") else "không text"] += 1
tsheet = collections.Counter()
tsrc = collections.Counter()
for k, m, c in checks:
    for p in c["message"].get("parts", []):
        if p["kind"] == "cell_text" and p.get("text"):
            tsheet[p.get("template_sheet")] += 1
            t = p.get("template") or ""
            tsrc["nhiều ứng viên MaFile cùng text" if ";" in t else "_extracted (tên suy ra)" if t.startswith("_extracted") else "_forms MaFile/Xsd"] += 1
print("   cell_text có text theo sheet đọc:", dict(tsheet.most_common()), "; theo nguồn template:", dict(tsrc))
un = json.load(open(os.path.join(vdir, "messages_unresolved.json"), encoding="utf-8"))
print("   cell_text:", dict(ct), "; messages_unresolved.json:", un["count"], "mục; theo lý do:")
for r, n in collections.Counter((i["kind"], re.sub(r"[0-9A-Z_]{2,}.*$", "…", i["reason"] or "")[:60]) for i in un["items"]).most_common(12):
    print("     ", n, r)

# (5) V1-V22
def find(cls=None, ids=None, lines=None, cond=None, level=None, fn=None):
    out = []
    for k, m, c in checks:
        if cls and not k.startswith(cls + ".json"):
            continue
        if ids and not any(p.get("id") in ids for p in c["message"].get("parts", [])):
            continue
        if lines and not (lines[0] <= c["source"]["lines"][0] <= lines[1]):
            continue
        if cond and not re.search(cond, c["condition"] + " " + json.dumps(c["bindings"], ensure_ascii=False)):
            continue
        if level and c["level"] != level:
            continue
        if fn and not fn(m, c):
            continue
        out.append((k, m, c))
    return out


V = {
    "V1": find("Gtgt01Nckt212", {"E_64"}),
    "V2": find("Gtgt01Nckt212", {"E_15"}, level="yellow"),
    "V3": find("Gtgt01Nckt212", {"E_16"}, level="yellow"),
    "V4": find("Gtgt01Nckt212", {"0071", "E_57", "E_55"}, lines=(414, 546)),
    "V5": find("Gtgt01Nckt212", {"E_96", "E_97", "E_82"}, lines=(548, 675)),
    "V6": find("Gtgt01Nckt212", {"E_98", "E_81"}, lines=(676, 751)),
    "V7": find("Gtgt01Nckt212", {"E_94", "E_99"}, lines=(829, 1035)),
    "V8": find("Gtgt01Nckt212", {"01GTGT0009", "01GTGT0010"}),
    "V9": find("Tndn03TT80Nckt294", {"E_212"}),
    "V10": find("GeneralFunctions", lines=(12100, 12260)),
    "V11": find(None, {"E_227", "E_228"}),
    "V12": find(None, {"E_56", "E_29"}, lines=(771, 815)),
    "V13": find("DkTndn02TT80Nckt276", {"E_25", "E_19", "E_31", "E_32"}, lines=(1971, 2037)),
    "V14": find("DkTndn02TT80Nckt276", {"E_13", "E_19", "E_28"}, lines=(2045, 2122)),
    "V15": find("Tndn02TT89Nckt309", {"E_17", "E_19"}),  # NOTES ghi E_16/E_18 = chỉ số 0 của Cells[16]/[18]; ô thật là E17/E19
    "V16": find("TncnQtt05TT80Nckt293", lines=(1040, 1090)),
    "V17": find(None, cond=r"CheckTinhTheoCQT") + find("GeneralFunctions", fn=lambda m, c: m["method"].startswith("CheckTinhTheoCQT")),
    "V18": find("AcBk01Nckt201", {"E_32", "E_33", "E_34", "E_35", "E_37", "E_38", "E_41", "E_55"}),
    "V19": find("AcBc26Nckt208", lines=(1650, 1720)),
    "V20": find("Tts01TT40Nckt283", lines=(2740, 3240)),
    "V21": find(None, cond=r"CheckMaSoThue\w*\("),
    "V22": find(None, cond=r"\bValidate(Formula|Email|Int|Double|NumberBetween|MaxLen|MaSoThue|CMTND|BangNhau|LonHon|NhoHon|NhoHonHoacBang|MinValue|FullDate)\("),
}
print("== (5) V1-V22")
for v, hits in V.items():
    if hits:
        pk = [h for h in hits if h[2]["level"] in ("pink", "yellow")]
        k, m, c = (pk or hits)[0]
        print("   %s: %d bản ghi; ví dụ %s %s L%s level=%s msg=%s/%s \"%s\"" % (v, len(hits), k, m["method"], c["source"]["lines"][0], c["level"], c["message"]["kind"], c["message"].get("id"), (c["message"].get("text") or "")[:45]))
    else:
        print("   %s: 0 bản ghi" % v)

# (7) kiểm độc lập bindings / định danh thô / ký tự cắt; chỉ đọc JSON và text nguồn .cs (không dùng cờ do công cụ C# tự đặt)
_src_cache = {}


def _clean_lines(path):
    """Text nguồn với chuỗi/ký tự/chú thích thay bằng khoảng trắng (giữ số dòng)."""
    if path in _src_cache:
        return _src_cache[path]
    t = open(os.path.join(data, path), encoding="utf-8-sig").read()
    out = []
    i, n = 0, len(t)
    while i < n:
        ch = t[i]
        if t.startswith("//", i):
            j = t.find("\n", i)
            j = n if j < 0 else j
            out.append(" " * (j - i)); i = j
        elif t.startswith("/*", i):
            j = t.find("*/", i + 2)
            j = n if j < 0 else j + 2
            out.append(re.sub(r"[^\n]", " ", t[i:j])); i = j
        elif ch == '"' or (ch in "@$" and i + 1 < n and t[i + 1] in '"@$'):
            k = i
            verb = False
            while t[k] in "@$":
                verb |= t[k] == "@"; k += 1
            k += 1
            while k < n:
                if verb and t[k] == '"':
                    if k + 1 < n and t[k + 1] == '"':
                        k += 2; continue
                    break
                if not verb and t[k] == "\\":
                    k += 2; continue
                if t[k] == '"' or (not verb and t[k] == "\n"):
                    break
                k += 1
            out.append(re.sub(r"[^\n]", " ", t[i:k + 1])); i = k + 1
        elif ch == "'":
            m = re.match(r"'(\\.|[^\\'])[^']*'", t[i:i + 12])
            if m:
                out.append(" " * len(m.group(0))); i += len(m.group(0))
            else:
                out.append(ch); i += 1
        else:
            out.append(ch); i += 1
    lines = "".join(out).split("\n")
    # cặp ngoặc nhọn: (dòng mở, dòng đóng), dòng đánh số từ 1
    pairs, stack = [], []
    for ln, s in enumerate(lines, 1):
        for c in s:
            if c == "{":
                stack.append(ln)
            elif c == "}" and stack:
                pairs.append((stack.pop(), ln))
    _src_cache[path] = (lines, pairs)
    return _src_cache[path]


def _block_after(pairs, line):
    """Khối { } mở tại dòng >= line gần nhất (thân của for/foreach/while/using/catch bắt đầu ở line)."""
    cands = [p for p in pairs if p[0] >= line]
    return min(cands, key=lambda p: (p[0], -p[1])) if cands else None


def _loops(path):
    lines, pairs = _clean_lines(path)
    res = []
    for ln, s in enumerate(lines, 1):
        if re.match(r"^\s*(for|foreach|while)\s*\(", s) or re.match(r"^\s*do\s*$", s) or re.match(r"^\s*do\s*\{", s):
            # while(...) đóng của do-while: dòng bắt đầu bằng "}" không khớp mẫu ở trên nên không tính trùng
            b = _block_after(pairs, ln)
            if b and b[0] <= ln + 1:
                end = b[1]
                if re.match(r"^\s*do\b", s):
                    end += 1  # dòng while(...) của do-while
                res.append((ln, end))
    return res


def _decl_scope(path, name, use_line, method_lines):
    """Khai báo của name (trong method) có phạm vi chứa use_line: trả (dòng khai báo, đầu phạm vi, cuối phạm vi)."""
    lines, pairs = _clean_lines(path)
    a, b = method_lines
    cands = []
    for ln in range(a, b + 1):
        s = lines[ln - 1]
        if re.search(r"\bforeach\s*\(\s*[\w.\[\],?]+(<[^()]*>)?(\.\w+)*\s+" + re.escape(name) + r"\s+in\b", s) or re.search(r"\bfor\s*\(\s*[\w.\[\],?]+\s+" + re.escape(name) + r"\s*=", s):
            blk = _block_after(pairs, ln)
            if blk:
                cands.append((ln, ln, blk[1]))
            continue
        if re.search(r"\bcatch\s*\(\s*[\w.]+\s+" + re.escape(name) + r"\s*\)", s) or re.search(r"\busing\s*\(\s*[\w.<>]+\s+" + re.escape(name) + r"\s*=", s):
            blk = _block_after(pairs, ln)
            if blk:
                cands.append((ln, ln, blk[1]))
            continue
        decl = re.search(r"(^|[(,;]|\bout\s)\s*(using\s+)?(?!return\b|else\b|new\b|out\b|ref\b|using\b)[A-Za-z_][\w.]*(<[^;=()]*>)?(\.\w+)*(\[\])*\??\s+" + re.escape(name) + r"\s*(=(?!=)|;|,|\))", s)
        pat = re.search(r"\b(is|out)\s+[A-Za-z_][\w.<>]*\s+" + re.escape(name) + r"\b", s)
        if decl or pat:
            inner = [p for p in pairs if p[0] < ln <= p[1] or (p[0] == ln and False)]
            if not inner:
                continue
            blk = min(inner, key=lambda p: p[1] - p[0])
            cands.append((ln, blk[0], blk[1]))
    hits = [c for c in cands if c[1] <= use_line <= c[2]]
    if not hits:
        return None, cands
    return min(hits, key=lambda c: c[2] - c[1]), cands


b_total = b_multi = b_fb = 0
after_bad, after_same_line, merge_bad, no_decl, decl_mismatch, defline_noname = [], [], [], [], [], []
meth_fb = sum(1 for _, m in methods if m.get("binding_method") == "fallback")
for k, m, c in checks:
    src = c["source"]["file"]
    info = c.get("binding_info") or {}
    if set(info) != set(c["bindings"]):
        merge_bad.append((k, c["source"]["lines"][0], "binding_info không khớp bindings"))
    for var, bi in sorted(info.items()):
        b_total += 1
        defs = bi.get("def_lines", [])
        if len(defs) > 1 or "phép gán có thể tới" in c["bindings"].get(var, ""):
            b_multi += 1
        if bi.get("method") == "fallback":
            b_fb += 1
        name = var.split("@")[0]
        U = bi["use_line"]
        loops = _loops(src)
        for d in defs:
            if d > U and not any(a0 <= U <= e0 and a0 <= d <= e0 for a0, e0 in loops):
                after_bad.append((k, c["source"]["lines"][0], var, U, d))
            elif d == U and len(defs) > 1 and not any(a0 <= U <= e0 for a0, e0 in loops):
                after_same_line.append((k, c["source"]["lines"][0], var, U, d))
        lines_src, _ = _clean_lines(src)
        for d in defs:
            if not re.search(r"\b" + re.escape(name) + r"\b", lines_src[d - 1]):
                defline_noname.append((k, c["source"]["lines"][0], var, d))
        sc, cands = _decl_scope(src, name, U, m["source"]["lines"])
        if sc is None:
            no_decl.append((k, c["source"]["lines"][0], var, U, [x[0] for x in cands]))
            continue
        if bi.get("decl_line") is not None and bi["decl_line"] != sc[0]:
            decl_mismatch.append((k, c["source"]["lines"][0], var, bi["decl_line"], sc[0]))
        for d in defs:
            if not (sc[1] <= d <= sc[2]) and d != sc[0]:
                merge_bad.append((k, c["source"]["lines"][0], var, "def L%d ngoài phạm vi khai báo L%d [%d-%d]" % (d, sc[0], sc[1], sc[2])))
print("== (7) bindings: %d biến (%d nhiều phép gán) trong %d check; binding_method=fallback: %d biến (%.2f%%); method fallback (không dựng được CFG): %d/%d" % (
    b_total, b_multi, sum(1 for _, _, c in checks if c["bindings"]), b_fb, 100.0 * b_fb / max(1, b_total), meth_fb, len(methods)))
print("   [kiểm] phép gán sau điểm dùng không qua vòng lặp: %d %s" % (len(after_bad), after_bad[:5]))
print("   [kiểm] phép gán cùng dòng điểm dùng (nhiều phép gán, ngoài vòng lặp): %d %s" % (len(after_same_line), after_same_line[:3]))
print("   [kiểm] binding gộp phép gán ngoài phạm vi khai báo (khác ISymbol): %d %s" % (len(merge_bad), merge_bad[:5]))
print("   [kiểm] decl_line khác khai báo tìm độc lập: %d %s; không tìm được khai báo chứa điểm dùng: %d %s" % (len(decl_mismatch), decl_mismatch[:3], len(no_decl), no_decl[:3]))
print("   [kiểm] dòng phép gán không chứa tên biến: %d %s" % (len(defline_noname), defline_noname[:3]))
ell = ell_tpl = gen_fields = 0
for f in sorted(glob.glob(os.path.join(vdir, "*.json")) + glob.glob(os.path.join(vdir, "_app", "*.json"))):
    ell += open(f, encoding="utf-8").read().count("…")
for k, m, c in checks:
    # text thông điệp chép nguyên văn từ template .xls/message.xml (phần message.text và từng part)
    for p in c["message"].get("parts", []):
        ell_tpl += (p.get("text") or "").count("…")
    ell_tpl += (c["message"].get("text") or "").count("…")
    gen_fields += json.dumps([c["condition"], c["cells_marked"], c["bindings"], c["message"].get("template"), c["message"].get("expr"), c.get("context")], ensure_ascii=False).count("…")
print("   [kiểm] ký tự '…': trong trường do công cụ sinh (condition/cells/bindings/template/expr/context) = %d; toàn bộ tệp JSON = %d, trong đó %d nằm trong text thông điệp chép nguyên từ template/message.xml" % (gen_fields, ell, ell_tpl))
ell = gen_fields + max(0, ell - ell_tpl)
# định danh C# thô theo (e): biến SheetView/FpSpread (khai báo trong tệp nguồn), lớp tĩnh, .Visible/.Sheets, ColumnLettersToIndex, model cũ
_typed = {}


def _sheet_vars(path):
    if path not in _typed:
        lines, _ = _clean_lines(path)
        t = "\n".join(lines)
        _typed[path] = set(re.findall(r"\b(?:SheetView|FpSpread)\s+(\w+)\s*[=;,)]", t))
    return _typed[path]


RAW = re.compile(r"\b(GeneralFunctions|HTKKFile|ValidateBase|ConvertHelpers|DateTime|Color)\.\w|\.Visible\b|\.Sheets\b|\bColumnLettersToIndex\b|\btKhaiInfoOld\b")
raw_bad = collections.Counter()
raw_ex = {}
for k, m, c in checks:
    txt = " ".join([c["condition"]] + c["cells_marked"] + list(c["bindings"].values()))
    txt = re.sub(r'"(?:[^"\\]|\\.)*"', '""', txt)
    txt = re.sub(r"S\[\w+\]", "S[]", txt)
    hit = RAW.search(txt)
    sv = [v for v in _sheet_vars(c["source"]["file"]) if re.search(r"(?<![\w.])" + re.escape(v) + r"\b", txt)]
    if hit or sv:
        key = hit.group(0) if hit else "sheetvar:" + sv[0]
        raw_bad[key] += 1
        raw_ex.setdefault(key, (k, c["source"]["lines"][0]))
print("   [kiểm] check còn định danh C# thô theo (e): %d %s %s" % (sum(1 for _ in raw_bad.elements()), dict(raw_bad.most_common(8)), list(raw_ex.items())[:4]))
ok7 = not after_bad and not merge_bad and ell == 0 and not raw_bad
okc = (sum(gm_json.values()) == 4533 and jp == 7241 and jp2 == 5 and jy == 434 and vn == 496 and va == 580)
print("== (7) KẾT LUẬN: bindings/định danh/ký tự cắt %s; số đếm GetMessage 4533, Pink 7241, Pink2 5, Yellow 434, Validate_N 496, Validate* 580 %s" % ("ĐẠT" if ok7 else "CHƯA ĐẠT", "ĐẠT" if okc else "CHƯA ĐẠT"))

# (6) mẫu đọc tay
if "--sample" in sys.argv:
    rnd = random.Random(int(os.environ.get("SAMPLE_SEED", "20261010")))
    pool = {True: [], False: []}
    for k, m in methods:
        if k.startswith("_app") or not m["checks"]:
            continue
        pool[recs[k]["key"]["in_menu"]].append((k, m))
    pick = rnd.sample(sorted(pool[True], key=lambda x: (x[0], x[1]["source"]["lines"][0])), 18) + \
        rnd.sample(sorted(pool[False], key=lambda x: (x[0], x[1]["source"]["lines"][0])), 12)
    chosen = [(k, m, rnd.choice(m["checks"])) for k, m in pick]
    rnd2 = random.Random(int(os.environ.get("SAMPLE_SEED", "20261010")) + 1)
    used = {(k, m["method"], m["source"]["lines"][0]) for k, m in pick}
    final = []
    for k, m, c in chosen:
        if m["in_scope"]:
            final.append((k, m, c, "goc"))
            continue
        cand = sorted([x for x in pool[recs[k]["key"]["in_menu"]] if x[1]["in_scope"] and (x[0], x[1]["method"], x[1]["source"]["lines"][0]) not in used],
                      key=lambda x: (x[0], x[1]["source"]["lines"][0]))
        k2, m2 = rnd2.choice(cand)
        used.add((k2, m2["method"], m2["source"]["lines"][0]))
        final.append((k2, m2, rnd2.choice(m2["checks"]), "thay(%s.%s)" % (k[:-5], m["method"])))
    for i, (k, m, c, how) in enumerate(final, 1):
        print("#%d|%s|%s|in_menu=%s|L%s-%s|check L%s|%s|%s" % (i, k, m["method"], recs[k]["key"]["in_menu"], *m["source"]["lines"], c["source"]["lines"][0], m["kind"], how))
        print(json.dumps({x: c[x] for x in ("condition", "bindings", "cells_marked", "level", "status")}, ensure_ascii=False))
        print("   msg:", c["message"]["kind"], c["message"].get("id"), (c["message"].get("text") or "")[:60], "| ctx:", c.get("context"))

# (6b) mẫu đọc tay binding nhiều phép gán: in binding, điều kiện riêng và mọi dòng nguồn của method có tên biến (thụt lề = cấu trúc khối)
if "--sample-bindings" in sys.argv:
    rnd = random.Random(int(os.environ.get("BIND_SEED", "20265151")))
    pool = []
    for k, m, c in checks:
        for var, bi in sorted((c.get("binding_info") or {}).items()):
            if len(bi.get("def_lines", [])) >= 2:
                pool.append((k, m, c, var))
    pool.sort(key=lambda x: (x[0], x[2]["source"]["lines"][0], x[3]))
    seen = set()
    picks = []
    for x in rnd.sample(pool, len(pool)):
        key = (x[0], x[1]["method"], x[3], x[2]["binding_info"][x[3]]["use_line"])
        if key in seen:
            continue
        seen.add(key)
        picks.append(x)
        if len(picks) == 30:
            break
    for i, (k, m, c, var) in enumerate(picks, 1):
        bi = c["binding_info"][var]
        name = var.split("@")[0]
        print("#B%d|%s|%s|check L%s|%s|use L%s|defs %s|decl L%s|%s" % (i, k, m["method"], c["source"]["lines"][0], var, bi["use_line"], bi["def_lines"], bi.get("decl_line"), bi["method"]))
        print("   value:", c["bindings"][var])
        src = open(os.path.join(data, c["source"]["file"]), encoding="utf-8-sig").read().split("\n")
        a0, b0 = m["source"]["lines"]
        for ln in range(a0, b0 + 1):
            t = src[ln - 1]
            if re.search(r"\b" + re.escape(name) + r"\b", t) or (ln == bi["use_line"]):
                mark = "U" if ln == bi["use_line"] else ("D" if ln in bi["def_lines"] else " ")
                print("   %s%5d %s" % (mark, ln, t.replace("\t", "  ")[:170]))

mc = os.path.join(vdir, "manual_check.csv")
if os.path.exists(mc):
    import csv
    rows = list(csv.DictReader(open(mc, encoding="utf-8")))
    for vg in sorted({r["vong"] for r in rows}):
        rr = [r for r in rows if r["vong"] == vg]
        n = len(rr)
        ok = sum(1 for r in rr if r["ket_qua_sau_sua"] == "dung")
        part = sum(1 for r in rr if r["ket_qua_sau_sua"] == "dung_mot_phan")
        first = sum(1 for r in rr if r["loi_lan_1"] == "khong")
        print("== (6) manual_check vòng %s (seed %s): %d check; đúng chặt %d (%.1f%%), một phần %d; không lỗi ở lần đọc đầu %d (%.1f%%)" % (
            vg, rr[0]["seed"], n, ok, 100.0 * ok / n, part, first, 100.0 * first / n))
