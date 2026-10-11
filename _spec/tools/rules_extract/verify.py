#!/usr/bin/env python3
"""Kiểm chứng _spec/rules: đếm độc lập bằng grep, so với kết quả công cụ.
Dùng: python3 -I verify.py <decompiledDir> <rulesDir> <classes.json>
Chỉ đọc văn bản; không chạy mã decompile."""
import collections, glob, json, os, re, subprocess, sys

dec, rules, classes_path = sys.argv[1], sys.argv[2], sys.argv[3]
ok = True

def say(flag, msg):
    global ok
    ok &= flag
    print(("ĐẠT " if flag else "TRƯỢT ") + msg)

# (1) điểm .Formula =
cmd = ["grep", "-rn", "--include=*.cs", r"\.Formula =", "tax.htkk.business", "htkk"]
out = subprocess.run(cmd, cwd=dec, capture_output=True, text=True).stdout.splitlines()
grep_lines = collections.Counter()
for l in out:
    f, n, _ = l.split(":", 2)
    grep_lines[(f, int(n))] += 1
print("Lệnh grep: (cd %s && %s | wc -l) -> %d dòng" % (dec, " ".join(cmd[:3] + ["'\\.Formula ='"] + cmd[4:]), len(out)))
pts = json.load(open(os.path.join(rules, "formula_points.json")))
by_status = collections.Counter(p["status"] for p in pts)
pt_lines = collections.Counter((p["file"], p["line"]) for p in pts)
ids = [(p["file"], p["line"], p["col"]) for p in pts]
say(len(ids) == len(set(ids)), "mỗi điểm xuất hiện đúng 1 lần trong formula_points.json (%d điểm)" % len(ids))
say(all(p["status"] in ("resolved", "parametric", "unresolved") for p in pts), "mỗi điểm thuộc đúng 1 nhóm resolved/parametric/unresolved")
print("   nhóm:", dict(sorted(by_status.items())), "tổng", sum(by_status.values()))
multi = {k: v for k, v in pt_lines.items() if v > 1}
missing = set(grep_lines) - set(pt_lines)
extra = set(pt_lines) - set(grep_lines)
say(not missing and not extra, "tập dòng grep trùng tập dòng điểm Roslyn (thiếu %d, thừa %d, dòng có >1 điểm %d)" % (len(missing), len(extra), len(multi)))
say(sum(by_status.values()) == len(out) + sum(v - 1 for v in multi.values()), "tổng các nhóm (%d) = số dòng grep (%d) + điểm phụ trên cùng dòng (%d)" % (sum(by_status.values()), len(out), sum(v - 1 for v in multi.values())))

# (2) hook theo loại: grep khai báo method trong 394 tệp class
classes = json.load(open(classes_path))["classes"]
kinds = ["AutoBindData", "ButtonClick", "CellChange", "CellEditchange", "ClearData", "Refresh", "UpdateFormula"]
rx = re.compile(r"^\s*(?:public|private|protected|internal)?\s*(?:static\s+)?[\w<>\[\]]+\s+(" + "|".join(kinds) + r")_\d+\s*\(", re.M)
gcount = collections.Counter()
for c in classes:
    txt = open(os.path.join(dec, c["path"]), encoding="utf-8-sig").read()
    for m in rx.finditer(txt):
        gcount[m.group(1)] += 1
idx = json.load(open(os.path.join(rules, "index.json")))
for k in kinds:
    say(idx["hook_counts"].get(k) == gcount[k], "hook %s: index.json %s = grep khai báo %d (chỉ mục T0 cho 394 class: %s)" % (k, idx["hook_counts"].get(k), gcount[k], idx["hook_counts_T0_index_394_classes"].get(k)))
# đối chiếu với hook_counts của chỉ mục T0 (mọi tệp): phần chênh nằm ngoài 394 class
t0 = json.load(open(os.path.join(os.path.dirname(classes_path), "index.json")))
paths394 = {c["path"] for c in classes}
outside = collections.Counter(); upd_rt = collections.Counter()
for fe in t0["files"]:
    for m in fe["methods"]:
        if not m.get("is_hook") or m.get("hook_kind") not in kinds:
            continue
        if fe["path"] not in paths394:
            outside[(m["hook_kind"], fe["path"] + ":" + m["name"])] += 1
        elif m["hook_kind"] == "UpdateFormula":
            upd_rt[m.get("return_type")] += 1
print("   T0 hook_counts (mọi tệp):", {k: t0["hook_counts"].get(k) for k in kinds})
for (k, where), n in sorted(outside.items()):
    print("   hook ngoài 394 class: %s %s x%d" % (k, where, n))
print("   UpdateFormula_N trong 394 class theo return_type:", dict(sorted(upd_rt.items())))
for k in kinds:
    n_out = sum(n for (kk, _), n in outside.items() if kk == k)
    say(t0["hook_counts"].get(k) == idx["hook_counts"].get(k) + n_out, "T0 %s %s = công cụ %s + ngoài 394 class %d" % (k, t0["hook_counts"].get(k), idx["hook_counts"].get(k), n_out))
nrec = 0
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b in ("index.json", "opaque_calls.json", "formula_points.json"):
        continue
    nrec += len(json.load(open(f))["records"])
say(nrec == idx["records"], "tổng bản ghi các tệp (%d) = index.json (%d)" % (nrec, idx["records"]))

# (3) ca tham chiếu
d = json.load(open(os.path.join(rules, "DkTndnVsp02TT80Nckt251.json")))
hit = []
for r in d["records"]:
    if r["hook"] != "CellChange_0" or r["sheetIndex"] != 0:
        continue
    for e in r["effects"]:
        if e["op"] == "setFormula" and e["target"] == {"sheet": 0, "cell": "AV101"}:
            hit.append((r, e))
good = [h for h in hit if isinstance(h[1]["formula"], dict) and
        [c["value"] for c in h[1]["formula"]["cases"]] == ["BI101", "BH101"] and
        h[1]["formula"]["cases"][0]["when"] == 'cell(S0!BH92).text == "VND"']
for r, e in hit:
    print("   AV101: trigger=%s guards=%s source=%s at=%s formula=%s status=%s" % (r["trigger"]["cells"] or r["trigger"]["rowRange"], r["guards"], r["source"]["lines"], e["at"], json.dumps(e["formula"], ensure_ascii=False), e["status"]))
say(len(good) == len(hit) == 2 and {tuple(h[0]["trigger"]["cells"]) for h in good} == {("N92",), ("AV99", "AV100")},
    "DkTndnVsp02TT80Nckt251.CellChange_0: AV101 = BI101 nếu ô tiền tệ S0!BH92 == VND, ngược lại BH101; kích hoạt N92 và AV99/AV100")

# (5) không chép đoạn C# dài
bad = []
pat = re.compile(r";|=>|\breturn\b|\n|\bvar\s|\bif\s*\(|\belse\s*(?:\{|if\b)|\bforeach\s*\(|\btry\b|\bcatch\b|\bthrow\b")
def walk(o, path):
    if isinstance(o, dict):
        for k, v in o.items():
            walk(v, path + "." + k)
    elif isinstance(o, list):
        for i, v in enumerate(o):
            walk(v, path)
    elif isinstance(o, str):
        if path.endswith((".formula", ".value")):
            o = " ".join(re.findall(r"\{[^{}]*\}", o))  # công thức/giá trị: chỉ xét phần placeholder {…}
        s2 = re.sub(r'"(?:\\.|[^"\\])*"', '""', o)  # bỏ nội dung hằng chuỗi (dữ liệu, không phải mã)
        if pat.search(s2) and not path.endswith(".file"):
            bad.append((path, s2))
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    walk(json.load(open(f)), os.path.basename(f))
# định danh C# thô trong trường nội dung (trừ tên hàm ở params.call/args và reason)
ident = re.compile(r"hTKKCMC|\.Cells\[|CellTemplate\.|\(int\)|\.Location\.|\bnew [A-Z]")
raw = []
def walk2(o, path):
    if isinstance(o, dict):
        for k, v in o.items():
            if k in ("reason", "call", "args", "file"):
                continue
            walk2(v, path + "." + k)
    elif isinstance(o, list):
        for v in o:
            walk2(v, path)
    elif isinstance(o, str) and ident.search(o):
        raw.append((path, o[:150]))
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    if os.path.basename(f) in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    walk2(json.load(open(f)), os.path.basename(f))
for b in raw[:5]:
    print("   định danh C# thô:", b)
say(not raw, "chuỗi chứa định danh C# thô (hTKKCMC, .Cells[, CellTemplate., (int), .Location., new X): %d" % len(raw))
xs = 0
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    if os.path.basename(f) in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    xs += len(re.findall(r"cell\(S(?:\d+!|\[)", open(f).read()))
print("   tham chiếu ô chéo sheet ghi tường minh cell(Sn!...): %d" % xs)
rxe = re.compile(r"\.(Cells|Rows|Columns)\[.*\]\.(Value|Text|Formula|Locked|CellType|BackColor|ForeColor|HorizontalAlignment|VerticalAlignment|Formatter|Font) = |\.Cells\[.*\]\.Reset(Formula|Value)\(")
tot = cov = drp = 0; miss = []
for c in classes:
    d = json.load(open(os.path.join(rules, c["class"] + ".json")))
    src = open(os.path.join(dec, d["file"]), encoding="utf-8-sig").read().split("\n")
    for h in d["hooks"]:
        ats = {e["at"] for r in d["records"] if r["hook"] == h["hook"] for e in r["effects"] if "via" not in e}
        dats = {x["at"] for x in h.get("dropped", [])}
        for ln in range(h["lines"][0], h["lines"][1] + 1):
            if rxe.search(src[ln - 1]):
                tot += 1
                if ln in ats: cov += 1
                elif ln in dats: drp += 1
                else: miss.append(d["file"] + ":" + str(ln))
for x in miss[:5]:
    print("   thiếu:", x)
say(not miss, "câu lệnh tác dụng trong thân hook (grep): %d, có effect %d, gắn dropped %d, thiếu %d" % (tot, cov, drp, len(miss)))
for b in bad[:10]:
    print("   nghi C#:", b[0], pat.findall(b[1]), b[1][:300])
say(not bad, "không chuỗi nào chứa dấu hiệu câu lệnh C# (; => return if else foreach try catch throw var xuống dòng): %d chuỗi vi phạm" % len(bad))
# ---------------- (7) kiểm độc lập vòng sửa 2 ----------------
def hook_records(d, hook):
    return [r for r in d["records"] if r["hook"] == hook]

# (7a) mọi câu gán thuộc tính / Reset* trên Cells/Rows/Columns trong thân hook: effect hoặc dropped, theo thuộc tính
rx_any = re.compile(r"\.(?:Cells|Rows|Columns)\[.*\]\.(\w+) = |\.(?:Cells|Rows|Columns)\[.*\]\.(Reset\w*)\(\)")
per_prop = collections.Counter(); per_prop_ok = collections.Counter(); miss7 = []
cls_data = {}
for c in classes:
    d = json.load(open(os.path.join(rules, c["class"] + ".json")))
    src = open(os.path.join(dec, d["file"]), encoding="utf-8-sig").read().split("\n")
    cls_data[c["class"]] = (d, src)
    for h in d["hooks"]:
        ats = {e["at"] for r in hook_records(d, h["hook"]) for e in r["effects"] if "via" not in e}
        dats = {x["at"]: x["reason"] for x in h.get("dropped", [])}
        for ln in range(h["lines"][0], h["lines"][1] + 1):
            m = rx_any.search(src[ln - 1])
            if not m:
                continue
            prop = m.group(1) or m.group(2)
            per_prop[prop] += 1
            if ln in ats or (ln in dats and dats[ln]):
                per_prop_ok[prop] += 1
            else:
                miss7.append("%s:%d %s" % (d["file"], ln, prop))
print("   câu gán/Reset theo thuộc tính (nguồn -> effect+dropped có lý do):", {k: "%d->%d" % (v, per_prop_ok[k]) for k, v in sorted(per_prop.items())})
for x in miss7[:8]:
    print("   thiếu:", x)
say(not miss7, "(7a) mọi câu gán thuộc tính/Reset* trên Cells/Rows/Columns trong thân hook có effect hoặc dropped kèm lý do: %d câu, thiếu %d" % (sum(per_prop.values()), len(miss7)))

# (7b) phép chuỗi/chuyển kiểu trong vế phải câu tác dụng (kể cả qua biến cục bộ gán trước đó trong hook) phải có biểu diễn trong JSON
OPS = {
    "Replace": (re.compile(r"\.Replace\("), re.compile(r"\breplace\(")),
    "Substring": (re.compile(r"\.Substring\("), re.compile(r"\bsubstring\(")),
    "Trim": (re.compile(r"\.Trim(?:Start|End)?\("), re.compile(r"\btrim(?:Start|End)?\(")),
    "ToUpper": (re.compile(r"\.ToUpper(?:Invariant)?\("), re.compile(r"\bupper\(")),
    "ToLower": (re.compile(r"\.ToLower(?:Invariant)?\("), re.compile(r"\blower\(")),
    "PadLeft": (re.compile(r"\.Pad(?:Left|Right)\("), re.compile(r"\bpad(?:Left|Right)\(")),
    "ToString(format)": (re.compile(r"\.ToString\([^)\s]"), re.compile(r"\bformat\(")),
    "Split": (re.compile(r"\.Split\("), re.compile(r"\bsplit\(")),
    "string.Format": (re.compile(r"\b[sS]tring\.Format\(\s*\"[^\"]*\{\d+[:,]"), re.compile(r"\bformat\(")),
    "Convert": (re.compile(r"\bConvert(?:Helpers)?\.(?:To|Convert)\w*\(|\b(?:int|long|double|decimal)\.Parse\("), re.compile(r"\b(?:num|int|toDate|bool)\(")),
}
asg = re.compile(r"^\s*(?:[\w.<>\[\],]+\s+)?([A-Za-z_]\w*)\s*(?:\+|-)?=(?!=)\s*(.+?);\s*$")
eff_rhs = re.compile(r"\.(?:Cells|Rows|Columns)\[.*\]\.(?:Value|Text|Formula) = (.+?);\s*$")
src_ops = collections.Counter(); rep_ops = collections.Counter(); miss_ops = []
def leaves(v):
    if isinstance(v, dict):
        return [x for c in v.get("cases", []) for x in leaves(c["value"])]
    return [v] if isinstance(v, str) else []
for cname, (d, src) in cls_data.items():
    for h in d["hooks"]:
        recs = hook_records(d, h["hook"])
        by_line = collections.defaultdict(list)
        for r in recs:
            for e in r["effects"]:
                if "via" not in e:
                    by_line[e["at"]].append(e)
        dats = {x["at"] for x in h.get("dropped", [])}
        varops = {}; varpat = {}
        for ln in range(h["lines"][0], h["lines"][1] + 1):
            t = src[ln - 1]
            em = eff_rhs.search(t)
            am = None if em else asg.match(t)
            rhs = em.group(1) if em else am.group(2) if am else None
            if rhs is None:
                continue
            ops = {k for k, (rs, _) in OPS.items() if rs.search(rhs)}
            pats = set(re.findall(r'\.Replace\("((?:\\.|[^"\\])+)"', rhs))
            for v in re.findall(r"(?<![\w.])([A-Za-z_]\w*)\b", rhs):
                ops |= varops.get(v, set()); pats |= varpat.get(v, set())
            if am:
                pl = re.search(r"\+=|-=", t)
                varops[am.group(1)] = (varops.get(am.group(1), set()) | ops) if pl else ops
                varpat[am.group(1)] = (varpat.get(am.group(1), set()) | pats) if pl else pats
                continue
            if not ops or ln in dats:
                continue
            effs = by_line.get(ln, [])
            for k in sorted(ops):
                src_ops[k] += 1
                js = [json.dumps(e, ensure_ascii=False) for e in effs]
                okk = any(OPS[k][1].search(j) or "opaque(" in j for j in js)
                # Replace trên mẫu hằng đã thay hết (khuôn công thức "###" -> dòng): mẫu không còn trong đầu ra
                if k == "Replace" and pats and effs:
                    okk = okk or all(all(json.loads('"' + p + '"') not in json.dumps(e.get("formula", e.get("value")), ensure_ascii=False) for p in pats) for e in effs)
                # toàn hằng (đã tính sẵn trên hằng số): không còn placeholder lúc chạy
                okk = okk or (effs and all(all("{" not in x for x in leaves(e.get("formula", e.get("value")))) for e in effs))
                if okk:
                    rep_ops[k] += 1
                else:
                    miss_ops.append("%s:%d %s" % (d["file"], ln, k))
print("   phép chuỗi/chuyển kiểu trong câu tác dụng (nguồn -> có biểu diễn):", {k: "%d->%d" % (src_ops[k], rep_ops[k]) for k in sorted(src_ops)})
jtxt = "".join(open(f).read() for f in sorted(glob.glob(os.path.join(rules, "*.json"))) if os.path.basename(f) not in ("opaque_calls.json", "formula_points.json", "index.json"))
print("   số lần xuất hiện trong JSON:", {k: len(OPS[k][1].findall(jtxt)) for k in OPS})
for x in miss_ops[:8]:
    print("   mất phép:", x)
say(not miss_ops, "(7b) phép chuỗi/chuyển kiểu không bị bỏ lặng lẽ (thiếu %d)" % len(miss_ops))

# (7c) biến cục bộ tự do: tên khai báo cục bộ trong tệp xuất hiện trong bản ghi mà không có trong symbols
decl = re.compile(r"^\t{3,}(?:const\s+)?(?!return\b|else\b|case\b|goto\b|throw\b|new\b|await\b|yield\b|using\b)[A-Za-z_][\w.]*(?:<[^;=()]*?>)?(?:\[,*\])*\??\s+([A-Za-z_]\w*)\s*(?:=(?!=)|;)", re.M)
extra_decl = [re.compile(r"\bfor\s*\(\s*(?:int|long|var)\s+(\w+)\s*="), re.compile(r"\bforeach\s*\(\s*[\w.<>\[\], ]+?\s+(\w+)\s+in\b"),
              re.compile(r"\bout\s+(?:[\w.<>\[\]]+\s+)?(\w+)\s*[),]"), re.compile(r"\bcatch\s*\(\s*\w+\s+(\w+)\s*\)")]
member = re.compile(r"^\t\t(?:public|private|protected|internal)[^(=;]*?\s(\w+)\s*(?:;|=|\{|$)", re.M)
tok = re.compile(r"(?<![\w.$])[A-Za-z_]\w*(?![\w(])")
lit = re.compile(r'"(?:\\.|[^"\\])*"')
def placeholders(t):
    out, dep = [], 0
    for ch in t:
        if ch == "{":
            if dep: out.append(ch)
            dep += 1
        elif ch == "}":
            dep -= 1
            out.append(ch if dep else " ")
        elif dep:
            out.append(ch)
    return "".join(out)
def rec_tokens(o, k, acc):
    if isinstance(o, dict):
        for kk, v in o.items():
            if kk in ("reason", "file", "call", "via", "note", "base", "symbols", "source", "key", "hook", "kind", "status", "prop", "after", "targetReason", "template"):
                continue
            rec_tokens(v, kk, acc)
    elif isinstance(o, list):
        for v in o:
            rec_tokens(v, k, acc)
    elif isinstance(o, str):
        t = placeholders(o) if k in ("formula", "value") else o
        t = re.sub(r"field\(\w+\)|obj\(\w*\)|opaque\([^(),]*\)", "", lit.sub('""', t))
        t = re.sub(r"opaque\([^(),]*, ", "(", t)
        acc.update(tok.findall(t))
free_recs = 0; free_ex = []; nrec_all = 0
files = [(c["class"], os.path.join(rules, c["class"] + ".json")) for c in classes]
for cname, path in files:
    d = cls_data[cname][0]
    txt = open(os.path.join(dec, d["file"]), encoding="utf-8-sig").read()
    locs = set(decl.findall(txt))
    for rx in extra_decl:
        locs |= set(rx.findall(txt))
    locs -= set(member.findall(txt)) | {"_"}  # "_" là discard của C#, không phải biến
    for r in d["records"]:
        nrec_all += 1
        acc = set(); rec_tokens(r, "", acc)
        fr = (acc & locs) - {s["name"] for s in r.get("symbols", [])}
        if fr:
            free_recs += 1
            if len(free_ex) < 8:
                free_ex.append("%s %s at=%s %s" % (cname, r["hook"], r["source"]["lines"], sorted(fr)[:6]))
# _shared.json, _app.json: biến cục bộ lấy theo tệp nguồn của từng bản ghi
loc_cache = {}
for extra in ("_shared.json", "_app.json"):
    for r in json.load(open(os.path.join(rules, extra)))["records"]:
        fsrc = r["source"]["file"]
        if fsrc not in loc_cache:
            txt = open(os.path.join(dec, fsrc), encoding="utf-8-sig").read()
            locs = set(decl.findall(txt))
            for rx in extra_decl:
                locs |= set(rx.findall(txt))
            loc_cache[fsrc] = locs - set(member.findall(txt)) - {"_"}
        nrec_all += 1
        acc = set(); rec_tokens(r, "", acc)
        fr = (acc & loc_cache[fsrc]) - {s["name"] for s in r.get("symbols", [])}
        if fr:
            free_recs += 1
            if len(free_ex) < 8:
                free_ex.append("%s %s at=%s %s" % (extra, r["hook"], r["source"]["lines"], sorted(fr)[:6]))
for x in free_ex:
    print("   biến tự do:", x)
say(free_recs == 0, "(7c) bản ghi có biến cục bộ tự do không khai báo trong symbols: %d / %d" % (free_recs, nrec_all))

# (7d) tham chiếu ô thiếu sheet
nos = collections.Counter()
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    if os.path.basename(f) in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    n = len(re.findall(r"\bcell\((?!S\d+!|S\[)", open(f).read()))
    if n:
        nos[os.path.basename(f)] = n
for k, v in nos.most_common(5):
    print("   cell thiếu sheet:", k, v)
say(not nos, "(7d) tham chiếu cell(...) thiếu sheet tường minh: %d lần trong %d tệp" % (sum(nos.values()), len(nos)))

# (7e) đường đi mâu thuẫn được gắn cờ: trigger.col/row == hai hằng khác nhau trong cùng bản ghi
bad_unr = 0; n_unr = 0
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    if os.path.basename(f) in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    for r in json.load(open(f))["records"]:
        n_unr += bool(r.get("unreachable"))
        eq = collections.defaultdict(set)
        for w in r["trigger"]["when"] + r["guards"]:
            m = re.fullmatch(r"(trigger\.(?:row|col)) == (\w+)", w)
            if m:
                eq[m.group(1)].add(m.group(2))
        if any(len(v) > 1 for v in eq.values()) and not r.get("unreachable"):
            bad_unr += 1
print("   bản ghi gắn cờ unreachable: %d" % n_unr)
say(bad_unr == 0, "(7e) bản ghi có trigger == hai hằng khác nhau mà không gắn cờ unreachable: %d" % bad_unr)
# ---------------- (7f)-(7k) kiểm độc lập vòng sửa 3 ----------------
# (7f) lời gọi method cùng class có ghi ô (bắt bằng regex trên văn bản, đóng bắc cầu theo đồ thị gọi) phải có effect qua via <tên>@<dòng gọi> hoặc dropped
mdecl = re.compile(r"^\t(?:(?:public|private|protected|internal|static|virtual|override|new|sealed|async|unsafe)\s+)*[\w<>\[\],.?]+(?:<[^=;()]*>)?\s+(\w+)\s*\([^=]*$")
wdirect = re.compile(r"\.(?:Cells|Rows|Columns)\[.*\]\.\w+ = |\.(?:Cells|Rows|Columns)\[.*\]\.Reset\w*\(\)|\b(?:ChangeValueCell|LockUnLockCols?|changFormatCell|checkLockUnLockCells)\(")
def method_bodies(src):
    res = collections.defaultdict(list)  # tên -> [(dòng đầu, dòng cuối)]
    i = 0
    while i < len(src):
        m = mdecl.match(src[i])
        if m and "=>" not in src[i] and i + 1 < len(src) and src[i + 1] == "\t{":
            j = i + 2
            while j < len(src) and src[j] != "\t}":
                j += 1
            res[m.group(1)].append((i + 1, j + 1))
            i = j
        i += 1
    return res
def writers(src, bodies):
    names = set(bodies)
    callrx = {n: re.compile(r"(?:(?<![\w.])|this\.)" + re.escape(n) + r"\s*\(") for n in names}
    direct, calls = set(), collections.defaultdict(set)
    for n, spans in bodies.items():
        for a, b in spans:
            for ln in range(a + 1, b):
                t = src[ln - 1]
                if wdirect.search(t):
                    direct.add(n)
                if "(" in t:
                    for n2, rx2 in callrx.items():
                        if n2 in t and rx2.search(t):
                            calls[n].add(n2)
    w = set(direct); ch = True
    while ch:
        ch = False
        for n in names:
            if n not in w and calls[n] & w:
                w.add(n); ch = True
    return w, callrx
tot_c = ok_c = drp_c = 0; miss_c = []
for cname, (d, src) in cls_data.items():
    bodies = method_bodies(src)
    w, callrx = writers(src, bodies)
    if not w:
        continue
    for h in d["hooks"]:
        recs = hook_records(d, h["hook"])
        vias = [(e.get("via"), e["op"]) for r in recs for e in r["effects"] if e.get("via")]
        dr = {(x["at"], x.get("call")): x["reason"] for x in h.get("dropped", []) if x.get("kind") == "call"}
        for ln in range(h["lines"][0] + 1, h["lines"][1] + 1):
            t = src[ln - 1]
            for n in sorted(w):
                if n in t and callrx[n].search(t):
                    tot_c += 1
                    pre = "%s@%d" % (n, ln)
                    if any(op != "opaque" and (v == pre or v.startswith(pre + " > ")) for v, op in vias):
                        ok_c += 1
                    elif dr.get((ln, n)):
                        drp_c += 1
                    else:
                        miss_c.append("%s:%d %s" % (d["file"], ln, n))
for x in miss_c[:8]:
    print("   thiếu effect/dropped cho lời gọi:", x)
say(not miss_c and tot_c > 0, "(7f) lời gọi method cùng class có ghi ô trong thân hook: %d, có effect qua via %d, dropped kèm lý do %d, thiếu %d" % (tot_c, ok_c, drp_c, len(miss_c)))

# (7g) lệnh giao diện: màu thẻ sheet, chữ nút, checkLockUnLockCells
UI = {"tab": (re.compile(r"TabStrip\[.*\]\.\w+\s*=(?!=)"), lambda e: e["op"] in ("setTabColor", "setTabProp")),
      "button": (re.compile(r"\b\w*[bB]uttonCellType\w*\.Text\s*=(?!=)"), lambda e: e["op"] == "setButtonText"),
      "chk": (re.compile(r"checkLockUnLockCells\("), lambda e: e["op"] == "lock" and e.get("params", {}).get("via") == "checkLockUnLockCells")}
ui_tot = collections.Counter(); ui_ok = collections.Counter(); ui_drp = collections.Counter(); ui_miss = []
for cname, (d, src) in cls_data.items():
    for h in d["hooks"]:
        recs = hook_records(d, h["hook"])
        dats = {x["at"] for x in h.get("dropped", []) if x.get("reason")}
        for ln in range(h["lines"][0], h["lines"][1] + 1):
            for k, (rx, pred) in UI.items():
                if rx.search(src[ln - 1]):
                    ui_tot[k] += 1
                    if any(e["at"] == ln and "via" not in e and pred(e) for r in recs for e in r["effects"]):
                        ui_ok[k] += 1
                    elif ln in dats:
                        ui_drp[k] += 1
                    else:
                        ui_miss.append("%s:%d %s" % (d["file"], ln, k))
hos = idx.get("hook_other_statements", {})
for k in UI:
    print("   %s: nguồn %d, effect %d, dropped %d; index.json %s" % (k, ui_tot[k], ui_ok[k], ui_drp[k], hos.get(k)))
for x in ui_miss[:8]:
    print("   thiếu:", x)
say(not ui_miss and all(hos.get(k, {}).get("total") == ui_tot[k] for k in UI), "(7g) TabStrip/buttonCellType.Text/checkLockUnLockCells: mỗi câu có effect hoặc dropped, số khớp index.json (thiếu %d)" % len(ui_miss))

# (7h) nhãn opaque(...) không chứa tên biến/trường/method C# của class
mname = re.compile(r"^\t(?:(?:public|private|protected|internal|static|virtual|override|new|sealed|async)\s+)*[\w<>\[\],.?]+\s+(\w+)\s*\(", re.M)
pdecl = re.compile(r"[(,]\s*(?:ref\s+|out\s+|params\s+)?[\w.<>\[\]?]+\s+(\w+)\s*(?:=[^,)]*)?(?=[,)])")
KIND_OK = re.compile(r"^(?:<\w+>\.\w+|[A-Za-z_]\w*(?:\.\w+)+|expr:\w+|binary:\w+|target:\w+|[a-z][a-z0-9-]*|[A-Z][A-Za-z]*)$")
# từ vựng loại của công cụ (nhãn cố định, không lấy từ mã nguồn): không đối chiếu với tên trong tệp
KINDS = {"null-deref", "array-index", "budget", "call", "cells-index", "compound", "div0", "format-template", "Format", "index-out-of-range", "lambda", "letters", "letters<1", "no-case", "pattern", "result", "tab-index", "too-many-cases", "branches"}
lab_bad = []; lab_n = 0
def heads7h(t):
    # nhãn opaque(Tên, đối số...): chỉ xét phần tên (trước dấu phẩy ngoài cùng)
    i = 0
    while True:
        i = t.find("opaque(", i)
        if i < 0:
            return
        j, dep = i + 7, 0
        while j < len(t) and not (dep == 0 and t[j] in ",)"):
            dep += {"(": 1, ")": -1}.get(t[j], 0); j += 1
        yield t[i + 7:j]
        i += 7
names_cache = {}
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    for r in json.load(open(f))["records"]:
        fsrc = r["source"]["file"]
        if fsrc not in names_cache:
            txt = open(os.path.join(dec, fsrc), encoding="utf-8-sig").read()
            nm = set(decl.findall(txt)) | set(member.findall(txt)) | set(mname.findall(txt)) | set(pdecl.findall(txt))
            for rx2 in extra_decl:
                nm |= set(rx2.findall(txt))
            names_cache[fsrc] = nm
        nm = names_cache[fsrc]
        for lab in heads7h(json.dumps({k: v for k, v in r.items() if k not in ("reason", "symbols", "cellVersions")}, ensure_ascii=False)):
            lab_n += 1
            first = lab.split(".")[0].split(":")[-1]
            unq = "." not in lab
            if lab in KINDS or lab.startswith(("op+", "op-", "op*", "op/", "op%")):
                continue
            if not KIND_OK.match(lab) or (unq and lab.split(":")[-1] in nm) or (not unq and not lab.startswith("<") and first in nm and first[:1].islower()):
                lab_bad.append("%s %s opaque(%s)" % (b, r["hook"], lab))
for x in sorted(set(lab_bad))[:8]:
    print("   nhãn lộ tên:", x)
# tên bị bỏ khỏi nhãn phải nằm trong symbols: opaque(branches) -> origin merge, opaque(result) -> method-result, opaque(call) -> method
NEED = {"branches": "merge", "result": "method-result", "call": "method"}
sym_miss = []; sym_n = 0
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    for r in json.load(open(f))["records"]:
        txt = json.dumps({k: v for k, v in r.items() if k not in ("reason", "symbols")}, ensure_ascii=False)
        origins = {s["origin"] for s in r.get("symbols", [])}
        for lab, org in NEED.items():
            if "opaque(%s)" % lab in txt:
                sym_n += 1
                if org not in origins:
                    sym_miss.append("%s %s opaque(%s)" % (b, r["hook"], lab))
for x in sym_miss[:5]:
    print("   thiếu symbols:", x)
say(not sym_miss, "(7h2) bản ghi có opaque(branches/result/call) đều có symbols origin merge/method-result/method: %d lượt, thiếu %d" % (sym_n, len(sym_miss)))
say(not lab_bad, "(7h) nhãn opaque(...) chỉ mô tả loại, không chứa tên biến/method C#: %d nhãn, vi phạm %d" % (lab_n, len(lab_bad)))

# (7i) trạng thái: effect resolved/parametric không được chứa opaque ở giá trị/tham số/đích hoặc điều kiện đường đi
st_bad = []
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    for r in json.load(open(f))["records"]:
        gop = any("opaque(" in g for g in r["guards"] + r["trigger"]["when"])
        for e in r["effects"]:
            if e["status"] in ("resolved", "parametric") and (gop or "opaque(" in json.dumps({k: v for k, v in e.items() if k != "status"}, ensure_ascii=False)):
                st_bad.append("%s %s at=%s" % (b, r["hook"], e["at"]))
        if r["status"] in ("resolved", "parametric") and any(e["status"] not in ("resolved", "parametric", "opaque") for e in r["effects"]):
            st_bad.append("%s %s bản ghi %s có effect unresolved" % (b, r["hook"], r["status"]))
for x in st_bad[:5]:
    print("   trạng thái sai:", x)
say(not st_bad, "(7i) effect resolved/parametric không chứa opaque (kể cả điều kiện): vi phạm %d" % len(st_bad))

# (7j) ca QC lần 3
def recs_of(c):
    return json.load(open(os.path.join(rules, c + ".json")))
cj = []
d1 = recs_of("BctcTt200DdNckt226")
for hk, ln in (("ButtonClick_0", 377), ("ButtonClick_1", 396), ("ButtonClick_2", 415), ("ButtonClick_3", 434)):
    effs = [e for r in d1["records"] if r["hook"] == hk for e in r["effects"] if (e.get("via") or "").startswith("getSoKyTruoc@%d" % ln) and e["op"] == "setValue"]
    cj.append(("(1) %s getSoKyTruoc@%d setValue qua via: %d" % (hk, ln, len(effs)), len(effs) > 0))
d2 = recs_of("Cnkd01TT40Nckt253")
bad2 = [r for r in d2["records"] if r["unreachable"] and any("cell(S0!B10).value" in g for g in r["guards"])]
cj.append(("(2) Cnkd01TT40Nckt253: bản ghi unreachable dựa trên ô B10 (bị ghi ở SetHinhThucKeKhai): %d" % len(bad2), not bad2))
d3 = recs_of("AcTb03Nckt210")
r87 = [r for r in d3["records"] for e in r["effects"] if e["at"] == 87]
cj.append(("(2) AcTb03Nckt210 dòng 87 có effect, không gắn unreachable: %d bản ghi, unreachable %d" % (len(r87), sum(r["unreachable"] for r in r87)), bool(r87) and not any(r["unreachable"] for r in r87)))
d4 = recs_of("CnkdBcSdhd01Nckt250")
tabs = {e["at"] for r in d4["records"] for e in r["effects"] if e["op"] == "setTabColor"}
cj.append(("(3) CnkdBcSdhd01Nckt250 setTabColor tại 2135, 2144: %s" % sorted(tabs & {2135, 2144}), {2135, 2144} <= tabs))
d5 = recs_of("Ttdb01ND373Nckt293")
h5 = [h for h in d5["hooks"] if h["hook"] == "ClearData_1"][0]
cj.append(("(3) Ttdb01ND373Nckt293:1440 checkLockUnLockCells có dropped (ngoại lệ chắc chắn dòng 1439) hoặc effect", any(x["at"] == 1440 for x in h5["dropped"]) or any(e["at"] == 1440 and e["op"] == "lock" for r in d5["records"] for e in r["effects"])))
d6 = recs_of("CnkdTcktHkd01Nckt289")
e6 = [e for r in d6["records"] for e in r["effects"] if e["at"] in (90, 91) and "via" not in e]
cj.append(("(4) CnkdTcktHkd01Nckt289:90-91 ToIntOrDefault -> int(ô): %d effect" % len(e6), bool(e6) and all("int(cell(" in json.dumps(e, ensure_ascii=False) for e in e6)))
d7 = recs_of("TncnXsbhdcHkd01Nckt289")
g7 = [g for r in d7["records"] for g in r["guards"] if "500000000" in g]
cj.append(("(4) TncnXsbhdcHkd01Nckt289:332 điều kiện %s" % sorted(set(g7))[:1], bool(g7) and all("int(cell(" in g for g in g7)))
d8 = recs_of("Gtgt02TT80Nckt251")
e8 = [e for r in d8["records"] for e in r["effects"] if e["at"] in (155, 159, 163, 167) and "via" not in e]
cj.append(("(6) Gtgt02TT80Nckt251:155-167 giữ chỉ số mảng, không còn null[...]: %s" % [json.dumps(e["value"], ensure_ascii=False)[:70] for e in e8[:2]], bool(e8) and all("null[" not in json.dumps(e["value"]) and re.search(r"\[\d+\]", json.dumps(e["value"])) for e in e8)))
for msg, good in cj:
    print("   " + ("đúng " if good else "SAI ") + msg)
say(all(g for _, g in cj), "(7j) các ca QC lần 3 (1)-(6) đúng")

# (7k) không còn dấu ẩn nội bộ trong JSON
leak = [os.path.basename(f) for f in glob.glob(os.path.join(rules, "*.json")) if "\\u001f" in open(f).read().lower() or "\x1f" in open(f).read()]
say(not leak, "(7k) JSON không chứa dấu ẩn nội bộ: %d tệp vi phạm" % len(leak))
# ---------- (8) QC lần 4 ----------
def all_records():
    for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
        b = os.path.basename(f)
        if b in ("formula_points.json", "opaque_calls.json", "index.json", "lock_config.json"):
            continue
        for r in json.load(open(f))["records"]:
            yield b, r
CMP_OPS = [" == ", " != ", " <= ", " >= ", " < ", " > "]
INV = {" == ": " != ", " != ": " == ", " <= ": " > ", " > ": " <= ", " >= ": " < ", " < ": " >= "}
def top_scan(s):
    """vị trí các toán tử ở mức ngoài cùng (ngoài ngoặc và chuỗi)"""
    dep, q, i, ops = 0, False, 0, []
    while i < len(s):
        ch = s[i]
        if q:
            if ch == "\\":
                i += 2; continue
            if ch == '"':
                q = False
        elif ch == '"':
            q = True
        elif ch in "([{":
            dep += 1
        elif ch in ")]}":
            dep -= 1
        elif dep == 0:
            for op in (" && ", " || ") + tuple(CMP_OPS):
                if s.startswith(op, i):
                    ops.append((i, op)); i += len(op) - 1; break
        i += 1
    return ops
def balanced_inner(s):
    if not (s.startswith("!(") and s.endswith(")")):
        return None
    dep = 0
    for i, ch in enumerate(s[1:], 1):
        if ch == "(":
            dep += 1
        elif ch == ")":
            dep -= 1
            if dep == 0 and i != len(s) - 1:
                return None
    return s[2:-1]
def negs(g):
    out = {"!(" + g + ")"}
    inner = balanced_inner(g)
    if inner is not None:
        out.add(inner)
    ops = top_scan(g)
    if len(ops) == 1 and ops[0][1] in INV:
        i, op = ops[0]
        out.add(g[:i] + INV[op] + g[i + len(op):])
    return out
gng_unr_lit = 0; gng_unr_opq = 0
gng = collections.Counter(); gng_unr = 0; gng_unr_k = collections.Counter(); gng_files = set(); gng_ex = []
for b, r in all_records():
    conj = [g for g in r["guards"] if not g.startswith("loop: ")] + list(r["trigger"]["when"])
    cs = set(conj)
    hit = [g for g in conj if negs(g) & cs]
    if hit:
        if r["unreachable"]:
            # định nghĩa đếm (legend.counting): cặp có g chứa opaque( không tính (mỗi lần đánh giá opaque là giá trị riêng)
            hit_real = [g for g in hit if "opaque(" not in g]
            if not hit_real:
                gng_unr_opq += 1
                continue
            if any("!(" + g + ")" in cs or (balanced_inner(g) or "\0") in cs for g in hit_real):
                gng_unr_lit += 1
            gng_unr += 1
            gng_unr_k["cell" if any("cell(" in g for g in hit) else "field" if any("field(" in g for g in hit) else "opaque" if any("opaque(" in g for g in hit) else "khác"] += 1
            continue
        kind = "cell" if any("cell(" in g for g in hit) else "field" if any("field(" in g for g in hit) else "opaque" if any("opaque(" in g for g in hit) else "khác"
        gng[kind] += 1; gng_files.add(b)
        if len(gng_ex) < 6:
            gng_ex.append("%s %s %s %s" % (b, r["hook"], r["source"]["lines"], hit[0][:90]))
for x in gng_ex:
    print("   g và !(g):", x)
print("   (8a) theo loại:", dict(gng), "; bản ghi unreachable có g và !(g) cùng phiên bản (mâu thuẫn thật, đã gắn cờ):", gng_unr, dict(gng_unr_k))
print("   (8a) định nghĩa đếm (legend.counting): literal g/!(g) %d; kể cả cặp so sánh bù %d; cặp chỉ trùng văn bản có opaque( (không tính) %d" % (gng_unr_lit, gng_unr, gng_unr_opq))
say(sum(gng.values()) == 0, "(8a) bản ghi (không gắn unreachable) có g và !(g) cùng phiên bản trong guards/trigger: %d (%d tệp)" % (sum(gng.values()), len(gng_files)))

lk_null = 0; lk_null_noreason = []; lk_tot = 0; lk_chk = 0; lk_chk_tgt = 0
for b, r in all_records():
    for e in r["effects"]:
        if e["op"] in ("lock", "unlock"):
            lk_tot += 1
            p = e.get("params", {})
            if p.get("via") == "checkLockUnLockCells":
                lk_chk += 1
                if e.get("target") is not None:
                    lk_chk_tgt += 1
            if e.get("target") is None:
                lk_null += 1
                if not p.get("targetReason"):
                    lk_null_noreason.append("%s %s at=%s" % (b, r["hook"], e["at"]))
for x in lk_null_noreason[:5]:
    print("   lock thiếu đích không lý do:", x)
say(not lk_null_noreason, "(8b) effect lock/unlock: %d; checkLockUnLockCells %d (có ô đích %d); target null %d, thiếu lý do (params.targetReason) %d" % (lk_tot, lk_chk, lk_chk_tgt, lk_null, len(lk_null_noreason)))

def rec_leaves(o, acc, k=""):
    if isinstance(o, dict):
        for kk, v in o.items():
            if kk not in ("reason", "symbols", "key", "source", "hook", "kind", "status", "note", "via", "file", "base", "targetReason"):
                rec_leaves(v, acc, kk)
    elif isinstance(o, list):
        for v in o:
            rec_leaves(v, acc, k)
    elif isinstance(o, str):
        acc.append(lit.sub('""', o))
def rec_text(r):
    acc = []; rec_leaves(r, acc)
    return "\n".join(acc)
nulls = 0; null_recs = 0; null_ex = []
bare_cl = 0; bare_ex = []; cnt_free = 0; cnt_ex = []
bare_rx = re.compile(r"(?<![\w.\])$@])([A-Za-z_]\w*)(?:@\d+)?\.(Count|Length)\b")
cnt_rx = re.compile(r"\b(?:count|length)\(([A-Za-z_]\w*)(?:@\d+)?\)")
VOCAB = {"tk", "ttc", "period", "header", "sheet", "spread", "trigger"}
for b, r in all_records():
    t = rec_text(r)
    n = len(re.findall(r"(?<![\w.])null\.", t))
    if n:
        nulls += n; null_recs += 1
        if len(null_ex) < 5:
            null_ex.append("%s %s %s" % (b, r["hook"], r["source"]["lines"]))
    syn = {s["name"] for s in r.get("symbols", [])}
    for m in bare_rx.finditer(t):
        bare_cl += 1
        if len(bare_ex) < 5:
            bare_ex.append("%s %s %s" % (b, r["hook"], m.group(0)))
    for m in cnt_rx.finditer(t):
        if m.group(1) not in syn and m.group(1) not in VOCAB:
            cnt_free += 1
            if len(cnt_ex) < 5:
                cnt_ex.append("%s %s count/length(%s)" % (b, r["hook"], m.group(1)))
for x in null_ex:
    print("   null.:", x)
say(nulls == 0, "(8c) \"null.\" trong bản ghi: %d lần, %d bản ghi" % (nulls, null_recs))
for x in bare_ex + cnt_ex:
    print("   Count/Length:", x)
say(bare_cl == 0 and cnt_free == 0, "(8d) .Count/.Length trực tiếp trên tên: %d; count(x)/length(x) với x không khai trong symbols: %d" % (bare_cl, cnt_free))

# (8e) symbols: một tên nhiều nguồn trong cùng bản ghi phải tách khóa tên@dòng (ref khác nhau, có mặt trong văn bản)
ms_bad = []; ms_n = 0
for b, r in all_records():
    by = collections.defaultdict(list)
    for s in r.get("symbols", []):
        by[s["name"]].append(s)
    t = json.dumps({k: v for k, v in r.items() if k not in ("reason", "symbols")}, ensure_ascii=False)
    for nm, l in by.items():
        if len(l) < 2:
            continue
        ms_n += 1
        refs = [s.get("ref") for s in l]
        if None in refs or len(set(refs)) != len(refs) or any(not re.search(r"(?<![\w.])" + re.escape(x) + r"(?!\d)", t) for x in refs):
            ms_bad.append("%s %s %s %s" % (b, r["hook"], nm, refs[:4]))
for x in ms_bad[:5]:
    print("   symbols nhiều nguồn:", x)
say(not ms_bad, "(8e) tên symbols nhiều nguồn trong cùng bản ghi: %d, chưa tách khóa tên@dòng %d" % (ms_n, len(ms_bad)))

# (8f) nhãn opaque mất đối số với hàm đã biết có tham số bắt buộc
sdecl = re.compile(r"public static [\w<>\[\],.?]+ (\w+)\(([^)]*)\)")
KNOWN = {"DateTime.Compare", "File.Exists", "Directory.Exists", "Path.Combine", "string.Join", "String.Join", "Regex.IsMatch", "Regex.Replace", "Regex.Match", "decimal.Round", "Math.Pow", "Convert.ToString"}
for cf in ("GeneralFunctions", "HTKKFile", "HTKKFormater", "ConvertHelpers"):
    p = os.path.join(dec, "tax.htkk.business/TAX.HTKK.Business", cf + ".cs")
    if not os.path.exists(p):
        continue
    arity = collections.defaultdict(set)
    for nm, ps in sdecl.findall(open(p, encoding="utf-8-sig").read()):
        req = [x for x in ps.split(",") if x.strip() and "=" not in x and not x.strip().startswith("params ")]
        arity[nm].add(len(req))
    for nm, s in arity.items():
        if min(s) > 0:
            KNOWN.add(cf + "." + nm)
def opaque_labels(t):
    i = 0
    while True:
        i = t.find("opaque(", i)
        if i < 0:
            return
        j, dep = i + 7, 1
        while j < len(t) and dep:
            dep += {"(": 1, ")": -1}.get(t[j], 0); j += 1
        yield t[i + 7:j - 1]
        i += 7
def rec_text_raw(r):
    acc = []; rec_leaves_raw(r, acc)
    return "\n".join(acc)
def rec_leaves_raw(o, acc):
    if isinstance(o, dict):
        for kk, v in o.items():
            if kk not in ("reason", "symbols", "key", "source", "hook", "kind", "status", "note", "via", "file", "base", "targetReason"):
                rec_leaves_raw(v, acc)
    elif isinstance(o, list):
        for v in o:
            rec_leaves_raw(v, acc)
    elif isinstance(o, str):
        acc.append(o)
# định nghĩa đếm nhãn opaque (legend.counting): số lần 'opaque(' trong lá chuỗi của bản ghi, không bỏ chuỗi literal
op_def = sum(rec_text_raw(r).count("opaque(") for _, r in all_records())
print("   nhãn opaque theo định nghĩa legend.counting: %d" % op_def)
lost = collections.Counter(); op_n = 0
for b, r in all_records():
    for lab in opaque_labels(rec_text(r)):
        op_n += 1
        head = lab.split(", ")[0]
        if head in KNOWN and head == lab:
            lost[head] += 1
print("   (8f) hàm đã biết có tham số bắt buộc: %d; nhãn mất đối số nhiều nhất: %s" % (len(KNOWN), lost.most_common(5)))
say(not lost, "(8f) nhãn opaque(Tên) của hàm đã biết mà không giữ đối số: %d / %d nhãn (tổng theo legend.counting, không bỏ literal: %d; số phân tích được sau khi bỏ chuỗi literal: %d)" % (sum(lost.values()), op_def, op_def, op_n))

# (8g) các ca QC lần 4 (1)-(7)
cq = []
for c in ("AcTb03Nckt210", "AcBc21Nckt210", "BlpBc21Nckt210"):
    dd = recs_of(c)
    rs = [r for r in dd["records"] if r["hook"] == "AutoBindData_0" and any(re.search(r"cell\(S0![A-Z]+8\)#\d+", g) for g in r["guards"]) and r["cellVersions"]]
    cq.append(("(1) %s AutoBindData_0: bản ghi có đọc ô dòng 8 sau lệnh ghi mang nhãn #dòng và cellVersions: %d" % (c, len(rs)), len(rs) > 0))
dd = recs_of("BctcTt200DdNckt226")
cd = [x for h in dd["hooks"] for x in h["dropped"] if x.get("kind") == "callee" and x.get("via", "").startswith("getSoKyTruoc@377") and x["reason"].startswith("nhánh hằng sai")]
cq.append(("(3) BctcTt200DdNckt226 dropped callee qua getSoKyTruoc@377 lý do nhánh hằng sai: %d" % len(cd), len(cd) > 0))
dd = recs_of("Gtgt02TT80Nckt251")
r4 = [r for r in dd["records"] if any(155 <= e["at"] <= 167 and "via" not in e for e in r["effects"])]
it4 = [[(s["at"]) for s in r["symbols"] if s["name"] == "item"] for r in r4]
cq.append(("(4) Gtgt02TT80Nckt251:155-167: %d bản ghi, null. %d, nguồn item %s" % (len(r4), sum(rec_text(r).count("null.") for r in r4), it4[:2]), bool(r4) and all(x == [375] for x in it4) and not any("null." in rec_text(r) for r in r4)))
n5 = sum(len(re.findall(r"(?<![\w.])(?:text3|list|sInfo)(?:@\d+)?\.(?:Count|Length)\b", rec_text(r))) for _, r in all_records())
cq.append(("(5) text3/list/sInfo .Count/.Length còn lại: %d" % n5, n5 == 0))
dc = [lab for _, r in all_records() for lab in opaque_labels(rec_text(r)) if lab.startswith("DateTime.Compare")]
r87 = [e for r in recs_of("AcTb03Nckt210")["records"] for e in r["effects"] if e["at"] == 87]
cq.append(("(6) opaque(DateTime.Compare...) %d nhãn, thiếu đối số %d; AcTb03Nckt210:87 GetStrByLengh có đối số: %s" % (len(dc), sum(1 for x in dc if x == "DateTime.Compare"), any("opaque(GeneralFunctions.GetStrByLengh, " in json.dumps(e, ensure_ascii=False) for e in r87)),
           bool(dc) and all(x != "DateTime.Compare" for x in dc) and any("opaque(GeneralFunctions.GetStrByLengh, " in json.dumps(e, ensure_ascii=False) for e in r87)))
dd = recs_of("Tts01TT40Nckt249")
d7 = [x for h in dd["hooks"] if h["hook"] == "CellChange_1" for x in h["dropped"] if x["at"] == 595]
tabs = {(r["source"]["file"], e["at"]) for _, r in all_records() for e in r["effects"] if e["op"] == "setTabColor" and "via" not in e and r["kind"] != "helper"}
cq.append(("(7) Tts01TT40Nckt249 CellChange_1 dropped@595 nêu RemoveDataRows: %s; setTabColor cấp hook %d vị trí" % (all("RemoveDataRows" in x["reason"] for x in d7) and bool(d7), len(tabs)), bool(d7) and all("RemoveDataRows" in x["reason"] for x in d7) and len(tabs) == 58))
for msg, good in cq:
    print("   " + ("đúng " if good else "SAI ") + msg)
say(all(g for _, g in cq), "(8g) các ca QC lần 4 (1)-(7) đúng")
# ---------- (9) QC lần 5 ----------
src_cache = {}
def src_lines(path):
    if path not in src_cache:
        src_cache[path] = open(os.path.join(dec, path), encoding="utf-8-sig").read().split("\n")
    return src_cache[path]
# (9a) dropped "nhánh hằng" tại điều kiện dựa trên kết quả hộp thoại: hộp thoại có chạy thì điều kiện không thể là hằng;
# hằng chỉ hợp lệ khi hộp thoại không chạy, lúc đó biến giữ giá trị khởi tạo -> hướng (luôn đúng/luôn sai) phải bằng điều kiện tính trên giá trị khởi tạo
dlg_assign = re.compile(r"\b(\w+) = (?:MessageBox\.Show|\(\(Form\)\w+\)\.ShowDialog|\w+\.ShowDialog)\(")
dlg_init = re.compile(r"DialogResult (\w+) = \(DialogResult\)(\d+);")
frm_new = re.compile(r"(\w+) = new frm\w*\(")
cond_cmp = re.compile(r"if \(\(int\)(\w+) (==|!=) (\d+)\)")
dlg_n = 0; dlg_bad = []
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    d = json.load(open(f))
    if "file" not in d:
        continue
    L = src_lines(d["file"])
    for h in d["hooks"]:
        for x in h["dropped"]:
            m = re.match(r"nhánh hằng sai: điều kiện dòng (\d+) luôn (đúng|sai)", x["reason"])
            if not m:
                continue
            n = int(m.group(1)); t = L[n - 1]
            # biến hộp thoại trong cùng method: tìm ngược tới khai báo method
            j = n - 1
            while j > 0 and not re.match(r"^\t(?:public|private|protected|internal)\b", L[j - 1]):
                j -= 1
            body = L[j - 1:n]
            dvars = {mm.group(1) for ln_ in body for mm in dlg_assign.finditer(ln_)}
            fvars = {mm.group(1) for ln_ in body for mm in frm_new.finditer(ln_)}
            cm = cond_cmp.search(t)
            if cm and cm.group(1) in dvars:
                dlg_n += 1
                init = {mm.group(1): int(mm.group(2)) for ln_ in body for mm in dlg_init.finditer(ln_)}
                v0 = init.get(cm.group(1))
                exp = None if v0 is None else ((v0 == int(cm.group(3))) == (cm.group(2) == "=="))
                if exp is None or exp != (m.group(2) == "đúng"):
                    dlg_bad.append("%s %s @%d %s" % (b, h["hook"], x["at"], x["reason"]))
            elif any(re.search(r"\b%s\.(?:Yes|No|DialogResult)\b" % re.escape(v), t) for v in fvars):
                dlg_n += 1; dlg_bad.append("%s %s @%d %s" % (b, h["hook"], x["at"], x["reason"]))
for x in dlg_bad[:5]:
    print("   dropped hằng dựa trên hộp thoại:", x)
print("   (9a) dropped trên nhánh không hiện hộp thoại (biến giữ giá trị khởi tạo, đúng hướng): %d" % (dlg_n - len(dlg_bad)))
say(not dlg_bad, "(9a) dropped có lý do hằng dựa trên ký hiệu hộp thoại (hướng khác giá trị khởi tạo, hoặc điều kiện form.Yes/No): %d" % len(dlg_bad))
# (9b) nhãn opaque của phép tìm/định dạng/hàm vô danh phải giữ đối số
lab_rx = re.compile(r"opaque\((Find|FirstOrDefault|First|Format|lambda)\)")
miss9b = collections.Counter(); lam_n = 0; find_n = 0
for _, r in all_records():
    t = rec_text_raw(r)
    for mm in lab_rx.finditer(t):
        miss9b[mm.group(1)] += 1
    lam_n += t.count("opaque(lambda")
    find_n += len(re.findall(r"opaque\((?:Find|FirstOrDefault|First), ", t))
print("   (9b) opaque(Find|First|FirstOrDefault, ...) có đối số: %d nhãn" % find_n)
say(not miss9b and lam_n == 0, "(9b) opaque(Find|FirstOrDefault|First|Format|lambda) thiếu đối số: %d %s; opaque(lambda...) còn lại %d" % (sum(miss9b.values()), dict(miss9b), lam_n))
# (9c) không còn null[ hoặc null. (truy cập trên null) trong văn bản bản ghi
nn9 = 0; nn9_ex = []
for b, r in all_records():
    t = rec_text_raw(r)
    k = len(re.findall(r"(?<![\w.$@])null[\[.]", t))
    if k:
        nn9 += k
        if len(nn9_ex) < 5:
            nn9_ex.append("%s %s %s" % (b, r["hook"], r["source"]["lines"]))
for x in nn9_ex:
    print("   null[/null.:", x)
say(nn9 == 0, "(9c) \"null[\" hoặc \"null.\" trong bản ghi: %d" % nn9)
# (9d) cellVersions theo định nghĩa chặt: chỉ lệnh đổi GIÁ TRỊ ô tạo phiên bản (kind "ghi"): gán .Value/.Text/.Formula,
# ResetValue/ResetText/ResetFormula, Clear*, chèn/xóa hàng (dịch ô), helper ghi giá trị đã biết; đầu vòng lặp chỉ hợp lệ khi thân
# vòng có lệnh như vậy hoặc lời gọi có thể ghi. Lời gọi chưa rõ có ghi ô (kind "co_the_ghi") phải có mặt ở dòng được trỏ.
# Không tính là dòng ghi: dòng chỉ gán biến cục bộ hoặc chỉ có lời gọi không ghi ô (đọc ô, giao diện, xóa kiểm tra/ghi chú, màu nền),
# dòng chỉ đổi Visible/Height/Width, BackColor/ForeColor/Note/Locked/CellType/Font/căn lề/Border.
UI = {"SuspendLayout", "ResumeLayout", "SetActiveCell", "Focus", "ShowCell", "ShowActiveCell", "Invalidate", "Select", "BringToFront", "SetViewportTopRow", "SetViewportLeftColumn", "SetActiveSheet", "SetNodeBackGroundPink", "SetRowHeight", "SetColumnWidth", "Refresh", "Update", "PerformLayout", "ScrollTo", "SetFocus"}
# lời gọi không ghi ô: đọc ô/giao diện/ghi chú/kiểm tra (đối chiếu thân method trong _decompiled: không gán Value/Text/Formula, không Reset/Clear/chèn hàng)
NOWRITE9 = UI | {"RemoveValidate", "RemoveValidateCell", "GetMessage", "GetRowIndexInSViewByCompares", "ConvertCellValueToText", "UpdateFromSViewToSInfo",
                 "GetValue", "ChangeLockedCell", "GetPreferredHeight", "SetNodeBackGroundWhite", "SetNote", "ResetNote", "GetCurrentMethod", "Log", "LogErrorBigData"}
PURE9 = {"ToString", "ToStringOrDefault", "ToIntOrDefault", "ToBoolOrDefault", "Trim", "TrimEnd", "TrimStart", "ToUpper", "ToLower", "Equals", "Contains", "ContainsKey",
         "TryGetValue", "StartsWith", "EndsWith", "IndexOf", "LastIndexOf", "Substring", "Replace", "Split", "Format", "Concat", "Parse", "TryParse", "ToInt32", "ToInt64",
         "ToInt16", "ToDouble", "ToDecimal", "ToBoolean", "ConvertToInt", "ConvertToLong", "ConvertToDouble", "ConvertToDecimal", "ConvertToDouble1", "ConvertToDecimal1",
         "ConvertToBool", "ConvertToString", "IsNullOrEmpty", "IsNullOrWhiteSpace", "isNullString", "ArrayCellId", "ColumnLettersToIndex", "ColumnIndexToLetters",
         "GetValueCustom", "GetObjectValue", "Round", "Abs", "Max", "Min", "Floor", "Ceiling", "Count", "Any", "ToList", "ToArray", "Where", "Select", "FirstOrDefault",
         "First", "Find", "FindIndex", "FindAll", "PadLeft", "PadRight", "ConvertToDateTime", "ToDateTime", "Exists", "Sum", "GetLength", "AddDays", "AddMonths",
         "ParseExact", "Compare", "DaysInMonth", "Show", "WriteLine", "FromArgb", "FromHtml", "GetType", "Add", "Remove", "RemoveAt", "Insert", "Clear", "Value"}
KW9 = {"if", "for", "foreach", "while", "switch", "catch", "using", "lock", "return", "new", "typeof", "nameof", "else", "do", "sizeof", "default"}
VALW = re.compile(r"\.(?:Value|Text|Formula)\s*=(?!=)|\.Reset(?:Value|Text|Formula)\s*\(|\.Clear(?:Range|Data|Selection|All)\s*\(|\b(?:AddRows?|RemoveRows?|InsertRows?|DeleteRows?|InsertTableRow|RemoveDataRows|RecalculateAll|Recalculate)\s*\(|\.Rows\.(?:Add|Remove|Insert)\s*\(")
PROPW = re.compile(r"\.(?:Visible|Height|Width)\s*=(?!=)")
FMTW = re.compile(r"\.(?:BackColor|ForeColor|Note|NoteStyle|Locked|CellType|Font|HorizontalAlignment|VerticalAlignment|Formatter|Border|Tag|CanFocus)\s*=(?!=)|\]\.Reset(?!Value\b|Text\b|Formula\b)\w+\s*\(")
def calls9(s1):
    s2 = re.sub(r"\bnew\s+[\w.<>]+\s*\(", "(", s1)
    out = []
    for c in re.findall(r"(\w+)\s*\(", s2):
        if c in KW9:
            continue
        if re.match(r"Reset(?!Value$|Text$|Formula$)", c) and "_" not in c:
            continue  # ResetBackColor, ResetNote...: trả định dạng, không đổi giá trị
        out.append(c)
    return out
def maywrite_calls(s1):
    return [c for c in calls9(s1) if c not in NOWRITE9 and c not in PURE9 and not c.startswith(("SetNodeBackGround", "RemoveValidate"))]
def loop_body(L, n):
    # thân vòng lặp bắt đầu ở dòng n (1-based): khối {...} hoặc một câu lệnh
    out = []; depth = 0; started = False
    for j in range(n - 1, min(len(L), n + 3000)):
        t = lit.sub('""', L[j])
        if j > n - 1 or "{" in t:
            out.append(t)
        for ch in t:
            if ch == "{": depth += 1; started = True
            elif ch == "}": depth -= 1
        if started and depth <= 0:
            break
        if not started and j > n - 1 and t.strip().endswith(";"):
            break
    return out
def classify9(L, n, memo={}):
    """ghi | co_the_ghi | loop_ghi | loop_khong | visible | dinh_dang | chi_doc"""
    t = L[n - 1].strip(); s1 = lit.sub('""', t)
    if re.match(r"(for|foreach|while|do)\b", t):
        key = (id(L), n)
        if key not in memo:
            body = loop_body(L, n)
            memo[key] = "loop_ghi" if any(VALW.search(x) or maywrite_calls(x) for x in body) else "loop_khong"
        return memo[key]
    if VALW.search(s1):
        return "ghi"
    if maywrite_calls(s1):
        return "co_the_ghi"
    if PROPW.search(s1):
        return "visible"
    if FMTW.search(s1):
        return "dinh_dang"
    return "chi_doc"
def version_lines(v):
    if "writes" in v:
        return list(v["writes"])
    a2 = re.sub(r"^hợp nhánh dòng \d+: ", "", v["after"])
    return list(map(int, re.findall(r"\d+", " ".join(re.findall(r"dòng ([\d, ]+)", a2)))))
cv_n = 0; cv_bad = []; cv_cat = {}; kind_n = {}; kind_bad = []
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b in ("formula_points.json", "opaque_calls.json", "index.json"):
        continue
    d = json.load(open(f))
    if "records" not in d:
        continue
    for r in d["records"]:
        L = src_lines(r["source"]["file"])
        for v in r["cellVersions"]:
            k = v.get("kind")
            kind_n[k] = kind_n.get(k, 0) + 1
            if k not in ("ghi", "co_the_ghi") or (k == "co_the_ghi" and not v.get("calls")):
                kind_bad.append("%s %s %s kind=%s" % (b, r["hook"], v["ref"][-14:], k))
            for n in version_lines(v):
                cv_n += 1
                c = classify9(L, n) if 0 < n <= len(L) else "ngoai_tep"
                cv_cat[c] = cv_cat.get(c, 0) + 1
                if c not in ("ghi", "co_the_ghi", "loop_ghi"):
                    cv_bad.append("%s %s %s -> dòng %d [%s]: %s" % (b, r["hook"], v["ref"][-14:], n, c, (L[n - 1].strip() if 0 < n <= len(L) else "?")[:60]))
for x in sorted(set(cv_bad))[:6]:
    print("   cellVersions trỏ dòng không ghi giá trị:", x)
print("   (9d) phân loại dòng được trỏ:", json.dumps(dict(sorted(cv_cat.items())), ensure_ascii=False), "| kind:", json.dumps({str(k): n for k, n in sorted(kind_n.items(), key=lambda x: str(x[0]))}, ensure_ascii=False))
say(not cv_bad, "(9d) cellVersions: %d tham chiếu dòng, trỏ dòng không đổi giá trị ô %d (chỉ đọc/gán biến cục bộ %d, Visible/Height/Width %d, BackColor/Note/Locked/CellType %d, vòng lặp không ghi %d)"
    % (cv_n, len(cv_bad), cv_cat.get("chi_doc", 0), cv_cat.get("visible", 0), cv_cat.get("dinh_dang", 0), cv_cat.get("loop_khong", 0)))
say(cv_cat.get("visible", 0) == 0 and cv_cat.get("dinh_dang", 0) == 0, "(9d) 0 phiên bản trỏ dòng Visible/Height/Width/BackColor/Note/Locked/CellType: %d" % (cv_cat.get("visible", 0) + cv_cat.get("dinh_dang", 0)))
for x in kind_bad[:3]:
    print("   kind sai:", x)
say(not kind_bad, "(9d) mỗi phiên bản có kind ghi|co_the_ghi (co_the_ghi kèm calls): sai %d" % len(kind_bad))
print("KẾT LUẬN:", "ĐẠT" if ok else "CÓ MỤC TRƯỢT")
sys.exit(0 if ok else 1)
