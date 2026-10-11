#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_counts.py: kiểm số đếm của T5 (xsd_merge) và T9 (biff_dv) từ output đã sinh.

Chạy:
  python3 -I _spec/tools/check_counts.py xsd <log xsd_merge>
  python3 -I _spec/tools/check_counts.py dv <thư mục dv có index.json>
In một dòng tóm tắt; thoát mã 0 khi khớp, 1 khi lệch.
"""
import glob
import json
import re
import sys


def check_xsd(log):
    t = open(log, encoding='utf-8').read()
    ok = 'loi parse: 0' in t and 'Tham chieu khong giai duoc: 0\n' in t
    rows = re.findall(r'^\s+(\w+)\s+(\d+) \|\s+(\d+) \|\s+(\d+) \|\s+(\d+) \(\d+/\d+\) \| json ngoai nguon (\d+)', t, re.M)
    # nguồn (parse) = JSON (vị trí duy nhất) + không đến được; JSON ngoài nguồn = 0
    bad = [r[0] for r in rows if int(r[3]) + int(r[4]) != int(r[1]) or r[5] != '0']
    print('xsd: %d khối đếm (kỳ vọng 4), parse/tham chiếu 0 lỗi: %s, khối lệch: %s' % (len(rows), ok, bad))
    return ok and len(rows) == 4 and not bad


def check_dv(d):
    idx = json.load(open(d + '/index.json', encoding='utf-8'))
    s_dv = sum(f['dv_count'] for f in idx['files'])
    s_dval = sum(f['dval_count'] for f in idx['files'])
    n = len([p for p in glob.glob(d + '/*.json') if not p.endswith('/index.json')])
    print('dv: DV %d, DVAL %d, tệp có DV %d / %d quét, tệp JSON %d, lỗi đọc %d' % (
        idx['total_dv'], idx['total_dval'], idx['files_with_dv'], idx['files_scanned'], n, len(idx['read_failures'])))
    return (not idx['read_failures'] and s_dv == idx['total_dv'] and s_dval == idx['total_dval']
            and n == idx['files_with_dv'] == len(idx['files']))


if __name__ == '__main__':
    ok = check_xsd(sys.argv[2]) if sys.argv[1] == 'xsd' else check_dv(sys.argv[2])
    sys.exit(0 if ok else 1)
