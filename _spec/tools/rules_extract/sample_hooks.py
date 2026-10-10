#!/usr/bin/env python3
"""Chọn 30 hook ngẫu nhiên (seed cố định, phân tầng: 5 hook dài trên 300 dòng đọc 15 dòng tác dụng mỗi hook,
25 hook còn lại chia in_menu true/false; chỉ hook có ít nhất 1 tác dụng),
in dòng nguồn có tác dụng (đầu/giữa/cuối) kèm điều kiện bao quanh và tác dụng trích được ở dòng đó để đọc tay.
Dùng: python3 -I sample_hooks.py <decompiledDir> <rulesDir> [seed]"""
import glob, json, os, random, re, sys

dec, rules = sys.argv[1], sys.argv[2]
seed = int(sys.argv[3]) if len(sys.argv) > 3 else 20261010
EFF = re.compile(r"\.(?:Cells|Rows|Columns)\[.*\]\.\w+ = |\.(Formula|Value|Text|Locked|CellType|BackColor) = |ChangeValueCell\(|LockUnLockCols?\(|changFormatCell\(|\.Reset\w*\(\)")
pool = {True: [], False: []}
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b.startswith("_") or b in ("index.json", "opaque_calls.json", "formula_points.json"):
        continue
    d = json.load(open(f))
    src = open(os.path.join(dec, d["file"]), encoding="utf-8-sig").read().split("\n")
    for h in d["hooks"]:
        s, e = h["lines"]
        lines = [i for i in range(s, e + 1) if EFF.search(src[i - 1])]
        if lines:
            pool[d["key"]["in_menu"]].append((d["key"]["Class"], h["hook"], d["file"], s, e, lines))
rnd = random.Random(seed)
longp = sorted([x for k in (True, False) for x in pool[k] if x[4] - x[3] > 300 and len(x[5]) >= 15])
pl = rnd.sample(longp, 5)
pick = [("long", x) for x in pl]
pick += [(True, x) for x in rnd.sample([x for x in pool[True] if x not in pl], 13)] + [(False, x) for x in rnd.sample([x for x in pool[False] if x not in pl], 12)]
for n, (im, (cls, hook, path, s, e, lines)) in enumerate(pick, 1):
    src = open(os.path.join(dec, path), encoding="utf-8-sig").read().split("\n")
    d = json.load(open(os.path.join(rules, cls + ".json")))
    k = 15 if im == "long" else 6
    chosen = lines if len(lines) <= k else sorted({lines[(len(lines) - 1) * i // (k - 1)] for i in range(k)})
    print("=" * 8, n, "in_menu=%s" % im, cls, hook, "%d-%d" % (s, e), "effect-lines=%d" % len(lines))
    for ln in chosen:
        # điều kiện bao quanh: quét ngược các dòng thụt lề nhỏ hơn
        ctx = []; ind = len(src[ln - 1]) - len(src[ln - 1].lstrip("\t"))
        for j in range(ln - 1, s - 1, -1):
            t = src[j - 1]; ti = len(t) - len(t.lstrip("\t"))
            if t.strip() and ti < ind and re.match(r"\s*(if|else|for|foreach|while|case|switch)\b", t):
                ctx.append("%d:%s" % (j, t.strip()[:150])); ind = ti
        for c in reversed(ctx[:4]):
            print("   ctx", c)
        print("   SRC %d: %s" % (ln, src[ln - 1].strip()[:260]))
        for r in d["records"]:
            if r["hook"] != hook:
                continue
            for ef in r["effects"]:
                if ef["at"] == ln and "via" not in ef:
                    v = ef.get("formula", ef.get("value", ef.get("params", "")))
                    print("   OUT %s S%s!%s %s | trig=%s %s | guards=%s | %s%s" % (ef["op"], (ef.get("target") or {}).get("sheet"), (ef.get("target") or {}).get("cell"), json.dumps(v, ensure_ascii=False)[:200], r["trigger"]["cells"][:4], r["trigger"]["rowRange"] or "", [g[:90] for g in r["guards"]][:4], ef["status"], " UNREACHABLE" if r.get("unreachable") else ""))
