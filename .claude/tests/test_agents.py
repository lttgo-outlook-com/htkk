#!/usr/bin/env python3
"""Kiểm khớp maxTurns giữa frontmatter và quy tắc "Ngân sách lượt" trong thân prompt.
Chạy: python3 -I /đường/dẫn/.claude/tests/test_agents.py [thư_mục_agents]"""
import re
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    yaml = None

DEFAULT_DIR = Path(__file__).resolve().parent.parent / "agents"


def parse_manual(text):
    d = {}
    for line in text.split("\n"):
        m = re.match(r"^([A-Za-z][A-Za-z0-9_]*):\s*(.*)$", line)
        if m:
            v = m.group(2).strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
                v = v[1:-1]
            elif re.fullmatch(r"-?\d+", v):
                v = int(v)
            d[m.group(1)] = v
    return d


def split(text):
    lines = text.split("\n")
    if not lines or lines[0] != "---":
        return None, text
    for i in range(1, len(lines)):
        if lines[i] == "---":
            return "\n".join(lines[1:i]), "\n".join(lines[i + 1:])
    return None, text


def parse_fm(fm):
    if yaml is not None:
        try:
            d = yaml.safe_load(fm)
            if isinstance(d, dict):
                return d
        except yaml.YAMLError:
            return None
        return None
    return parse_manual(fm)


def check(path):
    """Trả về danh sách (tên ca, lỗi hoặc None)."""
    res = []
    stem = path.stem
    fm_text, body = split(path.read_text(encoding="utf-8"))
    fm = parse_fm(fm_text) if fm_text is not None else None
    if fm is None:
        return [(f"{stem}: frontmatter parse được", "không tìm thấy hoặc không parse được frontmatter")]
    res.append((f"{stem}: frontmatter parse được", None))
    if stem == "coordinator":
        res.append((f"{stem}: không có maxTurns", "có khóa maxTurns" if "maxTurns" in fm else None))
        return res
    res.append((f"{stem}: name khớp tên tệp", None if fm.get("name") == stem else f"name={fm.get('name')!r}"))
    mt = fm.get("maxTurns")
    ok = isinstance(mt, int) and not isinstance(mt, bool) and mt > 0
    res.append((f"{stem}: maxTurns là số nguyên dương", None if ok else f"maxTurns={mt!r}"))
    hits = [l for l in body.split("\n") if "Ngân sách lượt" in l]
    if len(hits) != 1:
        res.append((f"{stem}: đúng một dòng Ngân sách lượt", f"tìm thấy {len(hits)} dòng"))
        return res
    res.append((f"{stem}: đúng một dòng Ngân sách lượt", None))
    m = re.search(r"tối đa (\d+) lượt", hits[0])
    if not m:
        err = "không tìm thấy mẫu 'tối đa N lượt'"
    elif not ok:
        err = "maxTurns không hợp lệ để so"
    elif int(m.group(1)) != mt:
        err = f"prompt N={m.group(1)}, frontmatter maxTurns={mt}"
    else:
        err = None
    res.append((f"{stem}: N khớp maxTurns", err))
    return res


def main():
    d = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else DEFAULT_DIR
    files = sorted(d.glob("*.md"))
    if not files:
        print(f"FAIL: không có tệp .md trong {d}")
        return 1
    total = fails = 0
    for f in files:
        for name, err in check(f):
            total += 1
            if err:
                fails += 1
                print(f"FAIL {name}: {err}")
            else:
                print(f"PASS {name}")
    print(f"Tổng: {total} ca, PASS {total - fails}, FAIL {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
