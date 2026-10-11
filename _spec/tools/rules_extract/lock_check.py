#!/usr/bin/env python3
"""Đối chiếu ngẫu nhiên N effect lock (via checkLockUnLockCells) với mẫu .xls: ô đích phải nằm trong danh sách ô của công thức AM
cùng dòng cấu hình, sheet đích = sheet của ô AK; kiểm thêm giá trị cache của AM (chuỗi nối &) = nối giá trị các ô đích đọc từ sheet đích,
và giá trị cache của AK = giá trị ô điều khiển. Giá trị ô đọc bằng bản ghi giá trị BIFF (không qua bộ giải mã token).
Dùng: python3 -I lock_check.py <rulesDir> <lock_config.json> <formsRoot(thư mục chứa _forms)> <cfb.py dir> [seed] [N]"""
import glob, json, os, random, re, sys

rules, cfg_path, root, cfb_dir = sys.argv[1:5]
seed = int(sys.argv[5]) if len(sys.argv) > 5 else 1123
N = int(sys.argv[6]) if len(sys.argv) > 6 else 10
here = os.path.dirname(os.path.abspath(__file__))
sys.argv = [sys.argv[0], "", "", "", "", cfb_dir, ""]
src = open(os.path.join(here, "lock_config.py"), encoding="utf-8").read().replace("\nmain()\n", "\n")
exec(compile(src, "lock_config.py", "exec"))  # nạp lớp Book (đọc bản ghi BIFF) của công cụ
cfg = json.load(open(cfg_path))
pool = []
for f in sorted(glob.glob(os.path.join(rules, "*.json"))):
    b = os.path.basename(f)
    if b.startswith("_") or b in ("index.json", "formula_points.json", "opaque_calls.json"):
        continue
    d = json.load(open(f))
    for r in d["records"]:
        for e in r["effects"]:
            p = e.get("params", {})
            if e["op"] == "lock" and p.get("via") == "checkLockUnLockCells" and e.get("target") and re.fullmatch(r"[A-Z]+\d+", e["target"]["cell"]):
                pool.append((d["key"]["Class"], r["hook"], e["at"], e["target"]["sheet"], e["target"]["cell"], p.get("configRow"), p.get("template")))
pool = sorted(set(pool))
pick = random.Random(seed).sample(pool, min(N, len(pool)))
ok_all = True; books = {}
def val(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return "" if v is None else str(v)
for cls, hook, at, sh, cell, crow, tpl in pick:
    c = cfg[cls]
    rown = int(re.search(r"AK(\d+)", crow).group(1))
    row = next(x for x in c["rows"] if x["row"] == rown)
    tg = [(t["refSheet"], "%s%d" % (t["col"], t["row"])) for t in row.get("targets", [])]
    in_am = any(cell == x for _, x in tg) and sh == row["sheet"]
    if tpl not in books:
        books[tpl] = Book(os.path.join(root, tpl))
    bk = books[tpl]
    last = c["lastSheet"]
    am_cache = val(bk.cells[last].get((rown - 1, 38), {}).get("v"))
    def cv(s, a):
        m = re.fullmatch(r"([A-Z]+)(\d+)", a); col = 0
        for ch in m.group(1):
            col = col * 26 + ord(ch) - 64
        return val(bk.cells[s].get((int(m.group(2)) - 1, col - 1), {}).get("v"))
    concat = "".join(cv(s, a) for s, a in tg)
    ak_cache = val(bk.cells[last].get((rown - 1, 36), {}).get("v"))
    ak_ref = cv(row["refSheet"], "%s%d" % (row["controlCol"], row["controlRow"]))
    # tham chiếu tới ô trống: Excel tính ra 0 (giá trị cache của AK là 0)
    ak_eq = ak_cache == ak_ref or (ak_ref == "" and ak_cache == "0")
    good = in_am and am_cache == concat and ak_eq
    ok_all &= good
    print("%s %s %s:%d lock S%d!%s | cấu hình S%d!AK%d AM=%s | ô trong AM: %s | cache AM %r = nối ô %r: %s | cache AK %r = ô %s%d %r: %s | %s" % (
        "ĐÚNG" if good else "SAI", cls, hook, at, sh, cell, last, rown, row.get("AM"), in_am, am_cache[:30], concat[:30], am_cache == concat,
        ak_cache[:20], row["controlCol"], row["controlRow"], ak_ref[:20], ak_eq, tpl))
print("KẾT LUẬN:", "ĐẠT %d/%d" % (len(pick), len(pick)) if ok_all else "CÓ SAI")
