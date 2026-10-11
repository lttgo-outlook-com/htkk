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
T2 = ROOT + "/.claude/worktrees/agent-a69710347074048d8/_spec/rules/opaque_calls.json"
T3DIR = ROOT + "/.claude/worktrees/agent-ae09dfc0b7c50411a/_spec/validations"
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


def resolve_by_source(nm, cand, chk):
    """Gán class cho tên không qualified bằng cách đọc dòng mã nguồn của check: gọi có tiền tố class thì lấy class đó."""
    lo, hi = chk["source"]["lines"][0], chk["source"]["lines"][-1]
    try:
        with open(ROOT + "/" + chk["source"]["file"], encoding="utf-8", errors="replace") as fh:
            text = "\n".join(fh.read().split("\n")[lo - 1:hi])
    except OSError:
        return None
    hit = {c for c in cand if re.search(r"\b" + c.split(".")[0] + r"\." + nm + r"\(", text)}
    if len(hit) == 1:
        return next(iter(hit))
    return None


def t3_counts(byname):
    sites = set()
    excl = collections.Counter()
    ambiguous = {}
    files = sorted(glob.glob(T3DIR + "/*.json")) + sorted(glob.glob(T3DIR + "/_app/*.json"))
    for f in files:
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
                            cls = resolve_by_source(nm, cand, k)
                            ambiguous.setdefault(nm, collections.Counter())[cls or "chưa gán được"] += 1
                            if cls:
                                sites.add((src, ln, cls))
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


def worktree_head(name):
    """Đọc commit HEAD của một worktree bằng cách đọc tệp (không chạy git)."""
    try:
        head = open("%s/.git/worktrees/%s/HEAD" % (ROOT, name)).read().strip()
        if head.startswith("ref: "):
            ref = head[5:]
            rp = "%s/.git/%s" % (ROOT, ref)
            if os.path.exists(rp):
                return open(rp).read().strip()
            for ln in open(ROOT + "/.git/packed-refs"):
                if ln.strip().endswith(" " + ref):
                    return ln.split()[0]
            return "không tra được ref " + ref
        return head
    except OSError as e:
        return "không đọc được: %s" % e


def sha(path):
    import hashlib
    return hashlib.sha256(open(path, "rb").read()).hexdigest()[:16]


def input_versions():
    t3files = sorted(glob.glob(T3DIR + "/*.json")) + sorted(glob.glob(T3DIR + "/_app/*.json"))
    h = __import__("hashlib").sha256()
    for f in t3files:
        h.update(open(f, "rb").read())
    return {
        "T2": {"worktree": "agent-a69710347074048d8", "head_commit": worktree_head("agent-a69710347074048d8"),
               "file": "_spec/rules/opaque_calls.json", "sha256_16": sha(T2)},
        "T3": {"worktree": "agent-ae09dfc0b7c50411a", "head_commit": worktree_head("agent-ae09dfc0b7c50411a"),
               "files": len(t3files), "sha256_16_all_files": h.hexdigest()[:16],
               "messages_unresolved_sha256_16": sha(T3DIR + "/messages_unresolved.json")},
        "T0_index": {"worktree": "agent-a9857e03398898f03", "head_commit": worktree_head("agent-a9857e03398898f03"),
                     "sha256_16": sha(IDX)},
        "note": "Commit là HEAD của nhánh worktree lúc chạy. Thay đổi chưa commit trong worktree đó (T2, T3 đang được sửa) không hiện ở commit nên sha256 của tệp là định danh chính xác của dữ liệu đọc.",
    }


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
    print("T3 tên không qualified khớp nhiều class, kết quả gán theo mã nguồn:", {k: dict(v) for k, v in amb.items()})
    print("Tổng lượt gọi helper T2 =", sum(t2.values()), "| T3 =", sum(t3.values()), "| gộp =", grand)
    print("Số hàm khác nhau: T2 =", len(t2), "| T3 =", len(t3), "| gộp =", len(total))
    print("Số hàm trong tập phủ =", len(chosen))
    print("Tỷ lệ phủ (vòng lặp tích lũy) = %.4f%% (%d/%d)" % (100 * cum / grand, cum, grand))
    print("Tỷ lệ phủ (set-sum dict riêng)  = %.4f%% (%d/%d)" % (100 * cum2 / grand, cum2, grand),
          "| khớp:", cum == cum2)
    print("Tên chưa khớp index:", unmatched)

    overlay = {}
    for opn in ("manual_overlay.json", "manual_overlay2.json", "manual_overlay3.json", "manual_overlay4.json"):
        op = OUT + opn
        if os.path.exists(op):
            overlay.update(json.load(open(op)))

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

    json.dump({"generated_from": {"T2": T2.split("/.claude/")[1], "T3": "agent-ae09dfc0b7c50411a/_spec/validations",
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
    iv = input_versions()
    print("Phiên bản đầu vào:", json.dumps(iv, ensure_ascii=False))
    print("ConvertToDateTime không qualified:", {k: dict(v) for k, v in amb.items()})
    json.dump({"threshold_pct": 100 * THRESH, "input_versions": iv,
               "t3_denominator_definition": "Mỗi lượt T3 là một vị trí mã (file nguồn, dòng đầu của check hoặc dòng GetMessage, hàm helper) duy nhất, lấy từ: (a) mark_api của check status unresolved, (b) message.parts.api là HTKKFile.GetMessage của check unresolved, (c) tên hàm trong condition của check unresolved nếu khớp đúng một method của 5 class GeneralFunctions, ConvertHelpers, HTKKFile, ValidateBase, HTKKFormater (tên trùng nhiều class được gán theo tiền tố class ở dòng mã nguồn của check, không gán được thì loại), (d) mỗi mục của messages_unresolved.json là một vị trí HTKKFile.GetMessage. Quét _spec/validations/*.json và _spec/validations/_app/*.json, bỏ index.json. Loại direct_cell_read và frmMessage. Không tính method.calls[]. Hai nguồn trùng vị trí chỉ tính một lần. T3 có thể chồng lên T2 ở cùng dòng mã, tổng gộp là tổng lượt không khử chồng giữa T2 và T3.",
               "ambiguous_name_resolution": {k: dict(v) for k, v in amb.items()},
               "denominator": {"t2_helper_sites": sum(t2.values()), "t3_helper_sites_dedup": sum(t3.values()),
                               "combined": grand, "distinct_functions": len(total)},
               "excluded": {"t2_by_category": dict(ex2), "t3": dict(ex3)},
               "cover_set_size": len(chosen), "covered_sites": cum, "coverage_pct": round(100 * cum / grand, 4),
               "crosscheck_set_sum_equal": cum == cum2, "unmatched_in_index": unmatched,
               
               "assumptions": ["T2: chỉ category htkk-helper; bỏ bcl, object-member, class-local",
                               "T3: mark_api + message.parts.api + tên hàm trong condition của check unresolved + messages_unresolved; dedupe theo (file, dòng, hàm)",
                               "Chuẩn hóa: bỏ hậu tố /N, giữ tiền tố class; overload gộp"],
               "rows": cov_rows}, open(OUT + "coverage.json", "w"), ensure_ascii=False, indent=1)
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
