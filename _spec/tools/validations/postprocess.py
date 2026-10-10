# Hậu xử lý _spec/validations/*.json: thay định danh C# còn sót bằng ký hiệu trung tính, bổ sung khởi tạo cho binding "++",
# bỏ binding thừa, ghi bảng ánh xạ vào legend của index.json. Chỉ dùng thư viện chuẩn; không chạy dotnet.
# Dùng: python -I postprocess.py <validations_dir>      (ghi đè tại chỗ; chạy lại nhiều lần cho cùng kết quả)
# build.py sinh dữ liệu từ raw JSON của roslyn_validations; chạy postprocess.py sau build.py (xem run.sh).
import glob
import json
import os
import re
import sys

# ---- bảng ánh xạ (cũng được ghi vào legend["ánh xạ ký hiệu trung tính"]) ----
MAPPING = {
    "x.TotalDays": "days(x)",
    "workbook.ActiveSheetIndex": "active_sheet()",
    "sheet(...).ActiveRowIndex; sheet(...).ActiveCell.Row.Index": "active_row()",
    "GetType(x)": "type_of(x)",
    "x.Text": "text(x)",
    "IsNulOrEmpty(x)": "is_empty(x)",
    "ToStringOrDefault(x)": "str(x)",
    "tham số CultureInfo.* / DateTimeStyles.* trong parse/format": "bỏ khỏi lời gọi; ghi vào trường phụ culture / datetime_styles của check",
    "Where(src, (a => e)); FindAll(src, (a => e))": "where(src, a: e)",
    "Any/All/Find/FindIndex/Count/FirstOrDefault(src, (a => e))": "giữ tên hàm, đổi lambda thành a: e (Any, All thành any, all)",
}
LAMBDA_RENAME = {"Where": "where", "FindAll": "where", "Any": "any", "All": "all"}
LEGEND_KEY = "ánh xạ ký hiệu trung tính"
LEGEND_TEXT = ("định danh C# được thay bằng ký hiệu trung tính trong condition/bindings/cells_marked: " +
               "; ".join("%s thành %s" % (k, v) for k, v in MAPPING.items()) +
               "; binding biến vòng '++' ghi kèm '(init: <giá trị>)' lấy từ FOR ... FROM, '(init: ?)' khi không xác định; "
               "binding không còn được điều kiện (trực tiếp hoặc qua binding khác) tham chiếu thì bị bỏ; "
               "trường culture/datetime_styles của check ghi tham số văn hóa đã bỏ khỏi parse/format")


def mask_of(s):
    """Thay nội dung chuỗi "..." bằng ký tự \\x00 (giữ độ dài) để tìm cấu trúc ngoài chuỗi."""
    out = []
    ins = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ins:
            if ch == "\\" and i + 1 < len(s):
                out.append("\x00\x00")
                i += 2
                continue
            if ch == '"':
                ins = False
                out.append('"')
            else:
                out.append("\x00")
        else:
            if ch == '"':
                ins = True
            out.append(ch)
        i += 1
    return "".join(out)


def pairs_of(m):
    st = []
    pr = {}
    for i, ch in enumerate(m):
        if ch in "([{":
            st.append(i)
        elif ch in ")]}" and st:
            j = st.pop()
            pr[i] = j
            pr[j] = i
    return pr


def sub(s, pat, fn):
    """Thay theo regex trên phần ngoài chuỗi; fn(match) trả chuỗi thay."""
    m = mask_of(s)
    out = []
    last = 0
    for mo in re.finditer(pat, m):
        out.append(s[last:mo.start()])
        out.append(fn(mo, s))
        last = mo.end()
    out.append(s[last:])
    return "".join(out)


def receiver_start(m, pr, dot):
    """Vị trí bắt đầu biểu thức nhận (trước dấu chấm ở dot)."""
    p = dot - 1
    while p >= 0:
        ch = m[p]
        if ch in ")]":
            if p not in pr:
                return None
            p = pr[p] - 1
        elif ch.isalnum() or ch in "_.":
            p -= 1
        else:
            break
    return p + 1


def postfix_to_call(s, member, fname):
    """recv.member -> fname(recv) cho mọi lần xuất hiện ngoài chuỗi."""
    pat = re.compile(r"\." + member + r"\b")
    while True:
        m = mask_of(s)
        mo = pat.search(m)
        if not mo:
            return s
        pr = pairs_of(m)
        st = receiver_start(m, pr, mo.start())
        if st is None or st == mo.start():
            return s[:mo.start()] + s[mo.end():]  # không tìm được: bỏ hậu tố
        recv = s[st:mo.start()]
        if recv.startswith("(") and pr.get(st) == mo.start() - 1:
            recv = recv[1:-1]
        s = s[:st] + fname + "(" + recv + ")" + s[mo.end():]


def lambdas(s):
    pat = re.compile(r"\((\w+) =>\s*")
    while True:
        m = mask_of(s)
        mo = pat.search(m)
        if not mo:
            return s
        pr = pairs_of(m)
        close = pr.get(mo.start())
        if close is None:
            return s[:mo.start()] + mo.group(1) + ": " + s[mo.end():]
        body = s[mo.end():close]
        s = s[:mo.start()] + mo.group(1) + ": " + body + s[close + 1:]


def norm(s, info):
    if not isinstance(s, str):
        return s
    # định danh đơn giản
    s = sub(s, r"\bIsNulOrEmpty\(", lambda mo, t: "is_empty(")
    s = sub(s, r"\bToStringOrDefault\(", lambda mo, t: "str(")
    s = sub(s, r"\bGetType\(", lambda mo, t: "type_of(")
    # phiên âm vị trí hoạt động
    s = sub(s, r"(?:sheet\([^()]*\)|workbook)\.ActiveSheetIndex\b", lambda mo, t: "active_sheet()")
    s = sub(s, r"sheet\([^()]*\)\.ActiveRowIndex\b", lambda mo, t: "active_row()")
    s = sub(s, r"sheet\([^()]*\)\.ActiveCell\.Row\.Index\b", lambda mo, t: "active_row()")
    s = sub(s, r"\w+(?:\.\w+)*\.ActiveSheetIndex\b", lambda mo, t: "active_sheet()")
    s = sub(s, r"\w+(?:\.\w+)*\.ActiveRowIndex\b", lambda mo, t: "active_row()")
    s = sub(s, r"\w+(?:\.\w+)*\.ActiveCell\.Row\.Index\b", lambda mo, t: "active_row()")
    # culture / styles
    def cul(mo, t):
        info.setdefault("culture", set()).add(mo.group(1))
        return ""
    def sty(mo, t):
        info.setdefault("datetime_styles", set()).add(mo.group(1))
        return ""
    s = sub(s, r",\s*(?:null|CultureInfo\.\w+)(?=,\s*DateTimeStyles\.\w+)", lambda mo, t: "")
    s = sub(s, r",\s*DateTimeStyles\.(\w+)", sty)
    s = sub(s, r",\s*CultureInfo\.(\w+)", cul)
    s = sub(s, r"\(\s*CultureInfo\.(\w+)\s*\)", lambda mo, t: (info.setdefault("culture", set()).add(mo.group(1)), "()")[1])
    # hậu tố
    s = postfix_to_call(s, "TotalDays", "days")
    s = postfix_to_call(s, "Text", "text")
    # lambda
    s = lambdas(s)
    s = sub(s, r"\b(Where|FindAll|Any|All)\(", lambda mo, t: LAMBDA_RENAME[mo.group(1)] + "(")
    return s


def ident_set(text):
    return set(re.findall(r"[A-Za-z_]\w*", mask_of(text)))


def prune(rec):
    """Bỏ binding (và binding_info) không được condition/cells_marked/message tham chiếu, kể cả qua binding khác."""
    b = rec["bindings"]
    if not b:
        return 0
    base = {k: k.split("@")[0] for k in b}
    msg = rec.get("message", {})
    roots = rec["condition"] + " " + " ".join(rec.get("cells_marked", []))
    keep = set()
    frontier = ident_set(roots)
    for key in ("template", "expr"):
        frontier |= set(re.findall(r"[A-Za-z_]\w*", msg.get(key) or ""))
    seen = set()
    while frontier:
        names = frontier - seen
        seen |= names
        frontier = set()
        for k, v in b.items():
            if base[k] in names and k not in keep:
                keep.add(k)
                frontier |= ident_set(v)
    removed = [k for k in b if k not in keep]
    for k in removed:
        del b[k]
        rec.get("binding_info", {}).pop(k, None)
    return len(removed)


def fix_init(rec):
    n = 0
    for k, v in rec["bindings"].items():
        if "@L" in k and isinstance(v, str) and re.search(r": (\+\+|[+-]=)", v) and "init:" not in v:
            mo = re.search(r"FOR " + re.escape(k) + r" FROM (.+?) WHILE ", rec["condition"])
            init = mo.group(1) if mo else None
            if init is None:
                sib = rec["bindings"].get(k.split("@")[0], "")
                mo2 = re.search(r"L%s: ([^|]+?)(?: \||$)" % re.escape(k.split("@L")[1]), sib)
                if mo2 and not re.match(r"(\+\+|[+-]=)", mo2.group(1)):
                    init = mo2.group(1)
            rec["bindings"][k] = v + " (init: %s)" % (init if init else "?")
            n += 1
    return n


def dump(path, obj):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(obj, ensure_ascii=False, indent=1, sort_keys=True) + "\n")


def main():
    d = sys.argv[1]
    files = sorted(glob.glob(os.path.join(d, "*.json")) + glob.glob(os.path.join(d, "_app", "*.json")))
    pruned = inits = 0
    for f in files:
        if os.path.basename(f) in ("index.json", "messages_unresolved.json"):
            continue
        j = json.load(open(f, encoding="utf-8"))
        for m in j.get("methods", []):
            for c in m["checks"]:
                info = {}
                c["condition"] = norm(c["condition"], info)
                c["bindings"] = {k: norm(v, info) for k, v in c["bindings"].items()}
                c["cells_marked"] = [norm(x, info) for x in c["cells_marked"]]
                if c.get("message", {}).get("expr"):
                    c["message"]["expr"] = norm(c["message"]["expr"], info)
                for key, val in info.items():
                    c[key] = sorted(set(c.get(key, [])) | val)
                inits += fix_init(c)
                pruned += prune(c)
        dump(f, j)
    ip = os.path.join(d, "index.json")
    idx = json.load(open(ip, encoding="utf-8"))
    idx["legend"][LEGEND_KEY] = LEGEND_TEXT
    dump(ip, idx)
    print("postprocess: thêm init %d, bỏ binding thừa %d" % (inits, pruned))


main()
