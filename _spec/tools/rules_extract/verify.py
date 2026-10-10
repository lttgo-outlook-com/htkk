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
        h[1]["formula"]["cases"][0]["when"] == 'cell(BH92).text == "VND"']
for r, e in hit:
    print("   AV101: trigger=%s guards=%s source=%s at=%s formula=%s status=%s" % (r["trigger"]["cells"] or r["trigger"]["rowRange"], r["guards"], r["source"]["lines"], e["at"], json.dumps(e["formula"], ensure_ascii=False), e["status"]))
say(len(good) == len(hit) == 2 and {tuple(h[0]["trigger"]["cells"]) for h in good} == {("N92",), ("AV99", "AV100")},
    "DkTndnVsp02TT80Nckt251.CellChange_0: AV101 = BI101 nếu ô tiền tệ BH92 == VND, ngược lại BH101; kích hoạt N92 và AV99/AV100")

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
for b in bad[:10]:
    print("   nghi C#:", b[0], pat.findall(b[1]), b[1][:300])
say(not bad, "không chuỗi nào chứa dấu hiệu câu lệnh C# (; => return if else foreach try catch throw var xuống dòng): %d chuỗi vi phạm" % len(bad))
print("KẾT LUẬN:", "ĐẠT" if ok else "CÓ MỤC TRƯỢT")
sys.exit(0 if ok else 1)
