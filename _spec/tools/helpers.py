#!/usr/bin/env python3
"""Gom lời gọi helper opaque từ T2 và T3, chọn tập phủ >= 95%, sinh đặc tả.

Chạy: python3 -I _spec/tools/helpers.py
Không thực thi mã C#; chỉ đọc văn bản.
"""
import collections
import glob
import json
import os
import re
import sys

WT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ROOT = "/Users/tuanle/codes/HTKK"
T2 = ROOT + "/.claude/worktrees/agent-a638a7f65966b6d67/_spec/rules/opaque_calls.json"
T3DIR = ROOT + "/.claude/worktrees/agent-a51ffc579b757e7c6/_spec/validations"
IDX = ROOT + "/.claude/worktrees/agent-a9857e03398898f03/_spec/csharp/index.json"
DEC = ROOT + "/_decompiled/"
OUT = WT + "/_spec/helpers/"
CLASSES = {"GeneralFunctions", "ConvertHelpers", "HTKKFile", "ValidateBase", "HTKKFormater"}
THRESH = 0.95


def norm(name):
    return re.sub(r"/\d+$", "", name)


def load_index():
    idx = json.load(open(IDX))
    methods = collections.defaultdict(list)
    for f in idx["files"]:
        for m in f["methods"]:
            if m["class"] in CLASSES:
                methods[m["class"] + "." + m["name"]].append(
                    {"file": f["path"], "params": m["params"], "return_type": m["return_type"],
                     "start": m["start_line"], "end": m["end_line"]})
    byname = collections.defaultdict(set)
    for k in methods:
        byname[k.split(".", 1)[1]].add(k)
    return methods, byname


def t2_counts():
    c = collections.Counter()
    excluded = collections.Counter()
    for x in json.load(open(T2)):
        if x["category"] == "htkk-helper":
            c[norm(x["call"])] += x["sites"]
        else:
            excluded[x["category"]] += x["sites"]
    return c, excluded


def t3_counts(byname):
    sites = set()
    excl = collections.Counter()
    ambiguous = {}
    for f in sorted(glob.glob(T3DIR + "/*.json")):
        b = os.path.basename(f)
        if b in ("index.json", "messages_unresolved.json"):
            continue
        for me in json.load(open(f))["methods"]:
            for k in me.get("checks", []):
                if k["status"] != "unresolved":
                    continue
                src = k["source"]["file"]
                ln = k["source"]["lines"][0]
                api = norm(k["mark_api"] or "")
                if api in ("frmMessage",):
                    excl[api] += 1
                elif api:
                    sites.add((src, ln, api))
                for p in (k["message"] or {}).get("parts", []):
                    a = p.get("api")
                    if a == "HTKKFile.GetMessage":
                        sites.add((src, p.get("gm_line"), a))
                    elif a == "direct_cell_read":
                        excl[a] += 1
                    elif a == "frmMessage":
                        excl[a] += 1
                for nm in set(re.findall(r"\b([A-Za-z_]\w*)\(", k["condition"])):
                    cand = byname.get(nm)
                    if cand:
                        if len(cand) > 1:
                            # không đoán class: loại khỏi mẫu số, báo lại
                            ambiguous[nm] = ambiguous.get(nm, 0) + 1
                            continue
                        sites.add((src, ln, next(iter(cand))))
    # messages_unresolved: mỗi mục là một vị trí GetMessage chưa giải
    mu = json.load(open(T3DIR + "/messages_unresolved.json"))["items"]
    for it in mu:
        sites.add((it["file"], it["line"], "HTKKFile.GetMessage"))
    c = collections.Counter(s[2] for s in sites)
    return c, excl, ambiguous


def cover(total):
    ranked = sorted(total.items(), key=lambda kv: (-kv[1], kv[0]))
    grand = sum(total.values())
    cum = 0
    chosen = []
    for k, v in ranked:
        chosen.append(k)
        cum += v
        if cum / grand >= THRESH:
            break
    return ranked, chosen, grand, cum


def read_lines(path, s, e):
    with open(DEC + path, encoding="utf-8", errors="replace") as fh:
        L = fh.read().split("\n")
    return "\n".join(L[s - 1:e])


def features(body, selfkey, methods):
    tags = []
    chk = [("BackColor", "đánh dấu màu nền ô"), ("SetNodeBackGround", "đánh dấu nút XML (hồng/vàng)"),
           ("Locked", "khóa/mở khóa ô"), ("CellType", "đặt kiểu ô"), ("Formula", "đặt/đọc công thức ô"),
           ("InsertRows", "chèn dòng"), ("DeleteRows", "xóa dòng"), ("RemoveRows", "xóa dòng"),
           ("AddRows", "thêm dòng"),
           ("RowHeight", "đổi chiều cao dòng"), ("File.", "IO tệp"), ("StreamWriter", "IO tệp"),
           ("StreamReader", "IO tệp"), ("XmlDocument", "đọc/ghi XML"), ("MessageBox", "hộp thoại"),
           ("catch", "bắt ngoại lệ"), ("throw", "ném ngoại lệ"), ("Visible", "đổi hiển thị"),
           ("Enabled", "đổi trạng thái cho phép"), ("SetValue", "ghi giá trị ô"), ("Value =", "ghi giá trị ô"),
           ("Text =", "ghi giá trị ô"), ("Items", "thao tác danh sách combobox"),
           ("DataTable", "thao tác bảng dữ liệu")]
    for kw, desc in chk:
        hit = (re.search(r"\bFile\.", body) if kw == "File." else kw in body)
        if hit and desc not in tags:
            tags.append(desc)
    calls = []
    same = {k.split(".", 1)[1]: k for k in methods if k.startswith(selfkey.split(".")[0] + ".")}
    for k in methods:
        if k == selfkey:
            continue
        cls, nm = k.split(".", 1)
        if re.search(r"\b" + cls + r"\." + re.escape(nm) + r"\(", body):
            calls.append(k)
    for nm, k in same.items():
        if k != selfkey and k not in calls and re.search(r"(?<![\w.])" + re.escape(nm) + r"\(", body):
            calls.append(k)
    return tags, sorted(set(calls))


def ws(s):
    return re.sub(r"\s+", " ", s).strip()


def verbatim_guard(spec, bodies):
    bad = []
    for key, sp in spec.items():
        body = ws(bodies[key])

        def strs(o):
            if isinstance(o, str):
                yield o
            elif isinstance(o, dict):
                for v in o.values():
                    yield from strs(v)
            elif isinstance(o, list):
                for v in o:
                    yield from strs(v)
        skip = {"file"}
        for s in strs({k: v for k, v in sp.items() if k not in ("source", "overloads")}):
            t = ws(s).replace(key.split(".", 1)[1], "NAME")
            if ";" in t or "{" in t:
                bad.append((key, "ký tự ; hoặc {", t[:60]))
                continue
            for i in range(0, max(1, len(t) - 39)):
                if len(t) >= 40 and t[i:i + 40] in body:
                    bad.append((key, "trùng 40 ký tự với mã", t[i:i + 40]))
                    break
    return bad


def main():
    os.makedirs(OUT, exist_ok=True)
    methods, byname = load_index()
    t2, ex2 = t2_counts()
    t3, ex3, amb = t3_counts(byname)
    total = collections.Counter()
    for k, v in t2.items():
        total[k] += v
    for k, v in t3.items():
        total[k] += v
    ranked, chosen, grand, cum = cover(total)
    # kiểm chéo: tổng theo set-sum trên dict dựng riêng
    d2 = {}
    for k in set(t2) | set(t3):
        d2[k] = t2.get(k, 0) + t3.get(k, 0)
    cum2 = sum(d2[k] for k in chosen)
    unmatched = sorted(k for k in total if k not in methods)

    print("T2 loại trừ (không phải htkk-helper):", dict(ex2))
    print("T3 loại trừ:", dict(ex3))
    print("T3 tên không qualified khớp nhiều class (loại khỏi mẫu số, số check):", amb)
    print("Tổng lượt gọi helper T2 =", sum(t2.values()), "| T3 =", sum(t3.values()), "| gộp =", grand)
    print("Số hàm khác nhau: T2 =", len(t2), "| T3 =", len(t3), "| gộp =", len(total))
    print("Số hàm trong tập phủ =", len(chosen))
    print("Tỷ lệ phủ (vòng lặp tích lũy) = %.4f%% (%d/%d)" % (100 * cum / grand, cum, grand))
    print("Tỷ lệ phủ (set-sum dict riêng)  = %.4f%% (%d/%d)" % (100 * cum2 / grand, cum2, grand),
          "| khớp:", cum == cum2)
    print("Tên chưa khớp index:", unmatched)

    overlay = {}
    op = OUT + "manual_overlay.json"
    if os.path.exists(op):
        overlay = json.load(open(op))

    spec = {}
    bodies = {}
    for key in chosen:
        if key not in methods:
            continue
        ovs = methods[key]
        body = "\n".join(read_lines(o["file"], o["start"], o["end"]) for o in ovs)
        bodies[key] = body
        tags, calls = features(body, key, methods)
        ov = overlay.get(key)
        if ov:
            purpose, sem, conf = ov["purpose"], ov["semantics"], ov["confidence"]
            tags = ov.get("side_effects", tags)
            calls = ov.get("calls", calls)
        else:
            nm = key.split(".", 1)[1]
            purpose = "Hàm %s (suy từ tên và dấu hiệu trong thân hàm, chưa đọc kỹ)" % nm
            sem = {"precondition": "chưa xác định", "effects": "dấu hiệu: " + (", ".join(tags) or "không có"),
                   "returns": ("không trả giá trị" if all(o["return_type"] in ("void", None) for o in ovs)
                               else "kiểu " + "/".join(sorted({str(o['return_type']) for o in ovs}))),
                   "exceptions": "có bắt/ném ngoại lệ" if ("catch" in body or "throw" in body) else "không thấy"}
            conf = "thấp"
        spec[key] = {
            "name": key,
            "overloads": [{"params": o["params"], "return_type": o["return_type"],
                           "lines": [o["start"], o["end"]]} for o in ovs],
            "purpose": purpose, "semantics": sem, "side_effects": tags, "calls": calls,
            "source": {"file": "_decompiled/" + ovs[0]["file"], "lines": [[o["start"], o["end"]] for o in ovs]},
            "confidence": conf,
            "call_count": {"t2": t2.get(key, 0), "t3": t3.get(key, 0), "total": total[key]},
        }
    bad = verbatim_guard(spec, bodies)
    print("Kiểm tra không chép nguyên văn:", "ĐẠT" if not bad else "KHÔNG ĐẠT %s" % bad[:5])
    print("Số đặc tả có overlay tay:", sum(1 for k in spec if k in overlay),
          "| confidence:", dict(collections.Counter(v["confidence"] for v in spec.values())))

    json.dump({"generated_from": {"T2": T2.split("/.claude/")[1], "T3": "agent-a51ffc579b757e7c6/_spec/validations",
                                  "index": "agent-a9857e03398898f03/_spec/csharp/index.json"},
               "count": len(spec), "helpers": list(spec.values())},
              open(OUT + "helpers.json", "w"), ensure_ascii=False, indent=1)
    cov_rows = []
    run = 0
    for i, (k, v) in enumerate(ranked, 1):
        run += v
        cov_rows.append({"rank": i, "name": k, "t2": t2.get(k, 0), "t3": t3.get(k, 0), "total": v,
                         "cumulative": run, "cumulative_pct": round(100 * run / grand, 4),
                         "in_cover_set": k in chosen, "in_index": k in methods})
    json.dump({"threshold_pct": 100 * THRESH,
               "denominator": {"t2_helper_sites": sum(t2.values()), "t3_helper_sites_dedup": sum(t3.values()),
                               "combined": grand, "distinct_functions": len(total)},
               "excluded": {"t2_by_category": dict(ex2), "t3": dict(ex3)},
               "cover_set_size": len(chosen), "covered_sites": cum, "coverage_pct": round(100 * cum / grand, 4),
               "crosscheck_set_sum_equal": cum == cum2, "unmatched_in_index": unmatched,
               "ambiguous_condition_names": amb,
               "assumptions": ["T2: chỉ category htkk-helper; bỏ bcl, object-member, class-local",
                               "T3: mark_api + message.parts.api + tên hàm trong condition của check unresolved + messages_unresolved; dedupe theo (file, dòng, hàm)",
                               "Chuẩn hóa: bỏ hậu tố /N, giữ tiền tố class; overload gộp"],
               "rows": cov_rows}, open(OUT + "coverage.json", "w"), ensure_ascii=False, indent=1)
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
