# So sánh trước/sau và kiểm dòng nguồn. Dùng: python -I compare.py <before_dir> <after_dir> <data_root>
import glob
import json
import os
import random
import re
import sys

bdir, adir, data = sys.argv[1], sys.argv[2], sys.argv[3]


def load(d):
    out = {}
    for f in sorted(glob.glob(os.path.join(d, "*.json")) + glob.glob(os.path.join(d, "_app", "*.json"))):
        b = os.path.basename(f)
        if b in ("index.json", "messages_unresolved.json"):
            continue
        r = json.load(open(f, encoding="utf-8"))
        for m in r["methods"]:
            for c in m["checks"]:
                k = (c["source"]["file"], c["source"]["lines"][0], c["mark_api"], tuple(c["cells_marked"]) if False else 0)
                out.setdefault(k, []).append(c)
    return out


B, A = load(bdir), load(adir)
nb = sum(len(v) for v in B.values())
na = sum(len(v) for v in A.values())
old_style = sum(1 for v in B.values() for c in v if any("phép gán gần nhất" in x or "gán trong vòng lặp" in x for x in c["bindings"].values()))
new_multi = sum(1 for v in A.values() for c in v if any("phép gán có thể tới" in x for x in c["bindings"].values()))
changed = 0
for k, v in A.items():
    bv = B.get(k)
    if bv is None:
        continue
    for i, c in enumerate(v):
        if i < len(bv) and (bv[i]["condition"] != c["condition"] or bv[i]["bindings"] != c["bindings"]):
            changed += 1
print("check trước=%d sau=%d" % (nb, na))
print("trước: check có binding dạng 'phép gán gần nhất/gán trong vòng lặp' (một giá trị) = %d" % old_style)
print("sau: check có binding liệt kê nhiều phép gán = %d; check đổi condition hoặc bindings = %d" % (new_multi, changed))
q = sum(1 for v in A.values() for c in v if "?." in c["condition"] + json.dumps(c["bindings"], ensure_ascii=False) + "".join(c["cells_marked"]))
qb = sum(1 for v in B.values() for c in v if "?." in c["condition"] + json.dumps(c["bindings"], ensure_ascii=False) + "".join(c["cells_marked"]))
print("check chứa '?.': trước=%d sau=%d" % (qb, q))
for lab, D in (("trước", B), ("sau", A)):
    cm = [x for v in D.values() for c in v for x in c["cells_marked"]]
    nop = sum(1 for x in cm if not re.match(r"^(S\d+|S\[[^\]]*\]|Sx|[A-Za-z_0-9]+)!", x))
    print("cells_marked %s: %d ô, thiếu tiền tố sheet %d" % (lab, len(cm), nop))
dt = sum(1 for v in A.values() for c in v if "new DateTime" in c["condition"] + json.dumps(c["bindings"], ensure_ascii=False))
print("sau: còn 'new DateTime' = %d" % dt)

# kiểm dòng nguồn: dòng ghi ở source.lines phải chứa lệnh tương ứng
tok = lambda api: "SetNodeBackGround" if "SetNodeBackGround" in api else ("BackColor" if "BackColor" in api else ("frmMessage" if api == "frmMessage" else "GetMessage"))
cache = {}


def line(f, n):
    if f not in cache:
        cache[f] = open(os.path.join(data, f), encoding="utf-8-sig").read().split("\n")
    L = cache[f]
    return L[n - 1] if 0 < n <= len(L) else ""


for lab, D in (("trước", B), ("sau", A)):
    allc = [c for v in D.values() for c in v]
    ok = sum(1 for c in allc if tok(c["mark_api"]) in line(c["source"]["file"], c["source"]["lines"][0]))
    print("dòng lệnh khớp nội dung (%s): %d/%d" % (lab, ok, len(allc)))
allc = sorted([c for v in A.values() for c in v], key=lambda c: (c["source"]["file"], c["source"]["lines"][0]))
cl_tok = re.compile(r"\bif\b|\bfor\b|foreach|while|case |default:|catch|else|&&|\|\||\?|switch|\bdo\b|return|continue|break|\(")
tot = okc = 0
for c in allc:
    for n in c.get("condition_lines", []):
        tot += 1
        okc += bool(cl_tok.search(line(c["source"]["file"], n)))
print("condition_lines chứa từ khóa điều kiện: %d/%d" % (okc, tot))
rnd = random.Random(20263031)
smp = rnd.sample(allc, 50)
ok50 = sum(1 for c in smp if tok(c["mark_api"]) in line(c["source"]["file"], c["source"]["lines"][0]))
okif = 0
for c in smp:
    cl = c.get("condition_lines", [])
    okif += all(cl_tok.search(line(c["source"]["file"], n)) for n in cl)
print("mẫu 50 (seed 20263031): dòng lệnh khớp %d/50; condition_lines đều khớp %d/50" % (ok50, okif))
