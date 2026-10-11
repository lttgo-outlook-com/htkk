#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_cong_thuc.py: kiểm độc lập các _forms/**/cong_thuc.txt với _spec/rules (không dùng mã của render_cong_thuc.py).

Chạy: python3 -I _spec/tools/verify_cong_thuc.py [--repo .] [--excel 417348]
In: số dòng ô Excel (so với --excel), số dòng C#, số effect setFormula không unreachable của class đã map,
số nhánh cases thêm, class không map được; thoát mã 1 nếu lệch.
Dòng C# = dòng có cột nguồn bắt đầu 'C#:'; dòng Excel = cột nguồn là xls | xls1 | xlsx
(cột nguồn là cột thứ 4 khi tách theo ' | ' trong phần sheet, không tính ROWREPLACE/SUPBOOK).
Dòng C# kỳ vọng của một class = số effect setFormula không unreachable, cộng thêm (số nhánh - 1) cho effect có formula.cases."""
import os, sys, re, json, argparse, collections

CS_RE = re.compile(r' \| C#:([A-Za-z0-9_]+)\.[A-Za-z0-9_]+:L\S+(?: \| khi: .*)?$')
MAP_RE = re.compile(r'^=> \d+ dòng công thức C# trong file này \(\d+ effect setFormula, class (\S+)\)$')

SKIP = {'index.json', '_app.json', '_shared.json', 'opaque_calls.json', 'formula_points.json'}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', default=os.getcwd())
    ap.add_argument('--excel', type=int, default=417348)
    a = ap.parse_args()
    F = os.path.join(a.repo, '_forms')
    excel = cs = 0
    cs_by_class = collections.Counter()
    mapped = set()
    for dp, dn, fn in os.walk(F):
        if 'cong_thuc.txt' not in fn:
            continue
        mode = None
        for ln in open(os.path.join(dp, 'cong_thuc.txt'), encoding='utf-8').read().split('\n'):
            if ln.startswith('## sheet:'):
                mode = 'sheet'
            elif ln.startswith('## ROWREPLACE') or ln.startswith('## SUPBOOK'):
                mode = 'other'
            elif ln.startswith('=== FILE:'):
                mode = None
            m = MAP_RE.match(ln)
            if m:
                mapped.add(m.group(1))
            if mode != 'sheet' or ln.startswith(('#', '=>')) or not ln:
                continue
            m = CS_RE.search(ln)
            if m:
                cs += 1
                cs_by_class[m.group(1)] += 1
            elif re.search(r' \| (?:xls|xls1|xlsx)(?: \| (?:mảng|mất tham chiếu): .*)?$', ln):
                excel += 1
    eff = cases_extra = tot_eff = 0
    exp_lines = collections.Counter()
    unm = []
    rd = os.path.join(a.repo, '_spec', 'rules')
    for fn in sorted(os.listdir(rd)):
        if not fn.endswith('.json') or fn in SKIP:
            continue
        j = json.load(open(os.path.join(rd, fn), encoding='utf-8'))
        if not isinstance(j, dict) or 'records' not in j:
            continue
        cls = j['key']['Class']
        ne = nl = 0
        for r in j['records']:
            if r.get('unreachable'):
                continue
            for e in r['effects']:
                if e['op'] == 'setFormula':
                    ne += 1
                    f = e.get('formula')
                    nl += len(f['cases']) if isinstance(f, dict) and isinstance(f.get('cases'), list) else 1
        tot_eff += ne
        if cls in mapped:
            eff += ne
            exp_lines[cls] = nl
            cases_extra += nl - ne
        else:
            unm.append((cls, ne))
    bad = [c for c in set(exp_lines) | set(cs_by_class) | mapped if exp_lines[c] != cs_by_class[c]]
    print('dòng ô Excel: %d (kỳ vọng %d, lệch %d)' % (excel, a.excel, excel - a.excel))
    print('dòng C#: %d; effect setFormula không unreachable của class map được: %d; dòng thêm do nhánh cases: %d; '
          'class lệch số dòng: %d' % (cs, eff, cases_extra, len(bad)))
    print('class map được: %d; không map được: %d (effect setFormula %d); tổng effect setFormula không unreachable: %d'
          % (len(mapped), len(unm), sum(n for _, n in unm), tot_eff))
    for c, n in unm:
        print('  không map: %s (%d effect)' % (c, n))
    ok = excel == a.excel and cs == eff + cases_extra and not bad
    print('KẾT QUẢ: %s' % ('ĐẠT' if ok else 'LỆCH'))
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
