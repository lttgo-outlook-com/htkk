#!/usr/bin/env python3
"""Chọn 30 hook ngẫu nhiên (seed cố định, phân tầng: 5 hook dài trên 300 dòng đọc 15 dòng tác dụng mỗi hook,
25 hook còn lại chia in_menu true/false; ngưỡng dòng: mọi hook trong mẫu, kể cả hook thay slot, có ít nhất MINL dòng tác dụng, mặc định 15),
in dòng nguồn có tác dụng (đầu/giữa/cuối) kèm điều kiện bao quanh và tác dụng trích được ở dòng đó để đọc tay.
Dùng: python3 -I sample_hooks.py <decompiledDir> <rulesDir> [seed] [minLines (mặc định 15; 1 = bỏ ngưỡng)]"""
import glob, json, os, random, re, sys

dec, rules = sys.argv[1], sys.argv[2]
seed = int(sys.argv[3]) if len(sys.argv) > 3 else 20261010
MINL = int(sys.argv[4]) if len(sys.argv) > 4 else 15
EFF = re.compile(r"\.(?:Cells|Rows|Columns)\[.*\]\.\w+ = |\.(Formula|Value|Text|Locked|CellType|BackColor) = |ChangeValueCell\(|LockUnLockCols?\(|changFormatCell\(|\.Reset\w*\(\)|TabStrip\[|checkLockUnLockCells\(")
MDECL = re.compile(r"^\t(?:(?:public|private|protected|internal|static)\s+)*[\w<>\[\],.?]+\s+(\w+)\s*\(", re.M)
CALL = re.compile(r"^\s*(?:[\w<>\[\],.]+\s+\w+\s*=\s*|\w+\s*=\s*)?(?:this\.)?(\w+)\s*\(")
pool = {True: [], False: []}
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b.startswith("_") or b in ("index.json", "opaque_calls.json", "formula_points.json"):
        continue
    d = json.load(open(f))
    src = open(os.path.join(dec, d["file"]), encoding="utf-8-sig").read().split("\n")
    for h in d["hooks"]:
        s, e = h["lines"]
        mnames = set(MDECL.findall("\n".join(src)))
        lines = [i for i in range(s + 1, e + 1) if EFF.search(src[i - 1]) or ((CALL.match(src[i - 1]) or [None, None])[1] in mnames)]
        # ngưỡng dòng áp ngay khi lập pool: mọi tầng và mọi bậc nới khi thay slot đều kế thừa
        if len(lines) >= max(1, MINL):
            pool[d["key"]["in_menu"]].append((d["key"]["Class"], h["hook"], d["file"], s, e, lines))
rnd = random.Random(seed)
longp = sorted([x for k in (True, False) for x in pool[k] if x[4] - x[3] > 300 and len(x[5]) >= max(15, MINL)])
pl = rnd.sample(longp, 5)
pick = [("long", x) for x in pl]
pick += [(True, x) for x in rnd.sample([x for x in pool[True] if x not in pl], 13)] + [(False, x) for x in rnd.sample([x for x in pool[False] if x not in pl], 12)]
# Mẫu mới: loại hook đã đọc ở các vòng trước (manual_check*.csv trong rulesDir, trừ tệp của chính seed này);
# tầng dài: không hai hook cùng khuôn mã (>50% dòng chuẩn hóa giống nhau) hoặc cùng (họ class bỏ hậu tố phiên bản, tên hook);
# chỉ 3 họ có hook trên 300 dòng (DkTndn02, Tndn03, Tts01) nên không thể buộc mỗi họ 1 hook.
# Slot vi phạm thay bằng lượt rút kế tiếp từ Random(seed + 1), giữ nguyên các slot hợp lệ.
import csv
seen = set()
for f in sorted(glob.glob(os.path.join(rules, "manual_check*.csv"))):
    if "seed%d" % seed in os.path.basename(f):
        continue
    for row in csv.DictReader(open(f, encoding="utf-8")):
        seen.add((row["class"], row["hook"]))
def family(c):
    return re.sub(r"(?:TT\d+|ND\d+)?(?:View)?Nckt\d+$", "", c)
sig_cache = {}
def sig(x):
    if x[:2] not in sig_cache:
        src0 = open(os.path.join(dec, x[2]), encoding="utf-8-sig").read().split("\n")
        sig_cache[x[:2]] = {re.sub(r"\d+", "#", t.strip()) for t in src0[x[3] - 1:x[4]] if len(t.strip()) >= 15}
    return sig_cache[x[:2]]
def same_tpl(a, b):
    sa, sb = sig(a), sig(b)
    return len(sa & sb) / max(1, min(len(sa), len(sb))) > 0.5
rnd2 = random.Random(seed + 1)
taken = {x[:2] for _, x in pick}
replaced = []
for i, (im, x) in enumerate(pick):
    ok_long = lambda y, i=i: all((family(y[0]), y[1]) != (family(p[0]), p[1]) and not same_tpl(y, p) for j, (k2, p) in enumerate(pick) if k2 == "long" and j < i)
    if x[:2] not in seen and (im != "long" or ok_long(x)):
        continue
    cand = longp if im == "long" else pool[im]
    order = rnd2.sample(cand, len(cand))
    # nới dần ràng buộc khi hết ứng viên (seed 8086 từng hết ứng viên tầng dài, StopIteration):
    # (a) đủ ràng buộc; (b) bỏ ràng buộc khuôn mã/họ; (c) cho phép hook đã đọc; không còn thì giữ slot cũ
    tiers = [("", lambda y: y[:2] not in seen and (im != "long" or ok_long(y))),
             (" (nới: bỏ ràng buộc khuôn/họ)", lambda y: y[:2] not in seen),
             (" (nới: hook đã đọc ở vòng trước)", lambda y: True)]
    y, note = None, ""
    for note, cond in tiers:
        y = next((y for y in order if y[:2] not in taken and cond(y)), None)
        if y is not None:
            break
    if y is None:
        replaced.append("%d: %s.%s giữ nguyên (hết ứng viên)" % (i + 1, x[0], x[1]))
        continue
    replaced.append("%d: %s.%s -> %s.%s%s" % (i + 1, x[0], x[1], y[0], y[1], note))
    taken.add(y[:2]); pick[i] = (im, y)
print("THAY SLOT:", "; ".join(replaced) if replaced else "không")
for n, (im, (cls, hook, path, s, e, lines)) in enumerate(pick, 1):
    src = open(os.path.join(dec, path), encoding="utf-8-sig").read().split("\n")
    d = json.load(open(os.path.join(rules, cls + ".json")))
    assert len(lines) >= MINL, (cls, hook, len(lines))
    k = 15  # QC4: mỗi hook ít nhất 15 dòng (hook ít dòng tác dụng hơn thì đọc hết)
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
        cm = CALL.match(src[ln - 1])
        if cm and not EFF.search(src[ln - 1]):
            pre = "%s@%d" % (cm.group(1), ln)
            vs = [(r, ef) for r in d["records"] if r["hook"] == hook for ef in r["effects"] if (ef.get("via") or "") == pre or (ef.get("via") or "").startswith(pre + " > ")]
            ops = sorted({ef["op"] for _, ef in vs})
            print("   VIA %s: %d effect %s" % (pre, len(vs), ops))
            for r, ef in [x for x in vs if x[1]["op"] != "opaque"][:2]:
                v = ef.get("formula", ef.get("value", ef.get("params", "")))
                print("      at=%s %s S%s!%s %s | guards=%s | %s%s" % (ef["at"], ef["op"], (ef.get("target") or {}).get("sheet"), (ef.get("target") or {}).get("cell"), json.dumps(v, ensure_ascii=False)[:150], [g[:80] for g in r["guards"]][:3], ef["status"], " UNREACHABLE" if r.get("unreachable") else ""))
        for x in [h for h in d["hooks"] if h["hook"] == hook][0]["dropped"]:
            if x["at"] == ln:
                print("   DROPPED", x["reason"][:160])
        for r in d["records"]:
            if r["hook"] != hook:
                continue
            for ef in r["effects"]:
                if ef["at"] == ln and "via" not in ef:
                    v = ef.get("formula", ef.get("value", ef.get("params", "")))
                    print("   OUT %s S%s!%s %s | trig=%s %s | guards=%s | %s%s" % (ef["op"], (ef.get("target") or {}).get("sheet"), (ef.get("target") or {}).get("cell"), json.dumps(v, ensure_ascii=False)[:200], r["trigger"]["cells"][:4], r["trigger"]["rowRange"] or "", [g[:90] for g in r["guards"]][:4], ef["status"], " UNREACHABLE" if r.get("unreachable") else ""))
