# Quét định danh C# còn sót trong condition/bindings và binding "++" thiếu khởi tạo, đếm theo số check.
# Dùng: python -I scan_idents.py <validations_dir> [-v]
import glob
import json
import os
import re
import sys

PATS = {
    ".TotalDays": r"\.TotalDays",
    "ActiveSheetIndex": r"ActiveSheetIndex",
    "ActiveRowIndex": r"ActiveRowIndex",
    "GetType": r"GetType",
    "CultureInfo": r"CultureInfo",
    "DateTimeStyles": r"DateTimeStyles",
    "IsNulOrEmpty": r"IsNulOrEmpty",
    "ToStringOrDefault": r"ToStringOrDefault",
    "ActiveCell.Row.Index": r"ActiveCell\.Row\.Index",
    "ActiveCell.Column.Index": r"ActiveCell\.Column\.Index",
    "ActiveColumnIndex": r"ActiveColumnIndex",
    ".Text": r"\.Text\b",
    "lambda =>": r"=>",
}
VARRE = re.compile(r"(?<![\w.@])((?:num|text|flag|value|obj)\d*)(?![\w(])")
DEFRE = re.compile(r"(?:FOR(?: EACH)?|CASE) (\w+)|[(,] ?(\w+): ")
NOINIT = re.compile(r"^L\d+[^:]*: (\+\+|[+-]=)")


def main():
    d = sys.argv[1]
    cnt = {k: 0 for k in PATS}
    ex = {k: [] for k in PATS}
    tot = noinit = 0
    free = 0
    freeex = []
    for f in sorted(glob.glob(os.path.join(d, "*.json")) + glob.glob(os.path.join(d, "_app", "*.json"))):
        if os.path.basename(f) in ("index.json", "messages_unresolved.json"):
            continue
        for m in json.load(open(f, encoding="utf-8")).get("methods", []):
            for c in m["checks"]:
                tot += 1
                s = c["condition"] + " || " + json.dumps(c["bindings"], ensure_ascii=False)
                for k, p in PATS.items():
                    mm = re.search(".{0,50}" + p + ".{0,50}", s)
                    if mm:
                        cnt[k] += 1
                        if len(ex[k]) < 2:
                            ex[k].append("%s:%s %s" % (os.path.basename(f), c["source"]["lines"][0], mm.group(0)))
                cond = re.sub(r'"(?:[^"\\]|\\.)*"', '""', c["condition"])
                bound = {k.split("@")[0] for k in c["bindings"]} | set(x for t in DEFRE.findall(cond) for x in t if x)
                miss = sorted({x for x in VARRE.findall(cond) if x not in bound})
                if miss:
                    free += 1
                    if len(freeex) < 12:
                        freeex.append("%s:%s %s" % (os.path.basename(f), c["source"]["lines"][0], miss))
                for b, v in c["bindings"].items():
                    if "@L" in b and NOINIT.search(v) and "init:" not in v:
                        noinit += 1
    print("tổng check", tot)
    for k in PATS:
        print("%-22s %d" % (k, cnt[k]))
    print("binding ++/+= thiếu khởi tạo (không có 'init:')", noinit)
    print("check có biến num/text/flag/value/obj tự do trong condition (không binding, không FOR/lambda)", free)
    if "-v" in sys.argv:
        for x in freeex:
            print("tự do |", x)
        for k, e in ex.items():
            for x in e:
                print(k, "|", x)


main()
