"""Đo giới hạn còn lại của T2 từ _spec/rules (chỉ đọc). Dùng: python3 -I measure_limits.py <rulesDir>"""
import json, glob, re, sys, collections
R = sys.argv[1]
TAG = re.compile(r'cell\((?:[^()]|\([^()]*\))*\)#\d+')
K = ('too-many-cases', 'index-out-of-range', 'expr:SwitchExpression')


def leaves(o, out):
    if isinstance(o, dict):
        for v in o.values():
            leaves(v, out)
    elif isinstance(o, list):
        for v in o:
            leaves(v, out)
    elif isinstance(o, str):
        out.append(o)


pairs = 0
recs = set()
ex = []
pi = 0
pirec = []
op = collections.Counter()
op_with_reason = collections.Counter()
unr = 0
nrec = 0
for f in sorted(glob.glob(R + '/*.json')):
    if f.rsplit('/', 1)[1] in ('index.json', 'opaque_calls.json', 'formula_points.json'):
        continue
    d = json.load(open(f))
    if not isinstance(d, dict) or 'records' not in d:
        continue
    for i, r in enumerate(d['records']):
        nrec += 1
        if r.get('unreachable'):
            unr += 1
        have = {v['ref'] for v in r['cellVersions']}
        out = []
        leaves([r.get('effects'), r.get('guards')], out)
        tags = set()
        for s in out:
            tags.update(TAG.findall(s))
        miss = tags - have
        if miss:
            pairs += len(miss)
            recs.add((f, i))
            if len(ex) < 3:
                ex.append((f.rsplit('/',1)[1], r['hook'], sorted(miss)[:1]))
        for s in out:
            for k in K:
                op[k] += s.count('opaque(' + k)
        out2 = []
        leaves([r.get('reason')], out2)
        for s in out2:
            for k in K:
                op_with_reason[k] += s.count('opaque(' + k)
        if any(re.search(r'(?<![.\w])PeriodInfo\b', s) for s in out + out2):
            pi += 1
            pirec.append((f.rsplit('/',1)[1], r['hook']))
print('records', nrec, 'unreachable', unr)
print('opaque effects+guards', dict(op), 'them reason', dict(op_with_reason))
print('pairs_missing', pairs, 'records', len(recs), ex)
print('PeriodInfo records', pi, pirec[:3])
