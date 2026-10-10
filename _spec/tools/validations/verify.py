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
