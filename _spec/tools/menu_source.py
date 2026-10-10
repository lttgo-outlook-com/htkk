#!/usr/bin/env python3
"""Doi chieu menuID chi co trong ma voi cac ban menu*.xml; ghi menu_source vao by_menu.json va header vao app_branches.json.
Chay sau app_branches.py:  python3 -I menu_source.py
"""
import glob, json, os, re
import xml.etree.ElementTree as ET

REPO = '/Users/tuanle/codes/HTKK'
WT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
OUT = os.path.join(WT, '_spec', 'app')

srcs = {}
for f in sorted(glob.glob(REPO + '/_extracted/menu*.xml')):
    name = os.path.basename(f)
    raw = open(f, encoding='utf-8-sig', errors='replace').read()
    ids = {}
    try:
        root = ET.fromstring(raw.encode('utf-8'))
        for e in root.iter('Menu'):
            ids.setdefault(e.get('ID'), []).append({'caption': e.get('Caption'), 'ten_to_khai': e.get('TenToKhai'),
                                                    'MaFile': e.get('MaFile'), 'VersionXML': e.get('VersionXML'), 'in_comment': False})
        parse = 'xml'
    except ET.ParseError as ex:
        parse = 'loi parse: %s' % ex
    # ID nam trong chu thich <!-- -->
    cm = {}
    for c in re.findall(r'<!--(.*?)-->', raw, re.S):
        for m in re.finditer(r'<Menu\b[^>]*?\bID="([^"]+)"[^>]*>', c):
            a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(0)))
            cm.setdefault(m.group(1), []).append({'caption': a.get('Caption'), 'ten_to_khai': a.get('TenToKhai'),
                                                  'MaFile': a.get('MaFile'), 'VersionXML': a.get('VersionXML'), 'in_comment': True})
    srcs[name] = {'live': ids, 'comment': cm, 'parse': parse}

by = json.load(open(OUT + '/by_menu.json', encoding='utf-8'))
if by and isinstance(next(iter(by.values())), dict):
    by = {k: v['points'] for k, v in by.items()}
cur = srcs['menu.xml']['live']
not_in = sorted(k for k in by if k not in cur)
print('So ID chi co trong ma (khong co trong menu.xml):', len(not_in))
for n, s in srcs.items():
    hit = [k for k in not_in if k in s['live'] or k in s['comment']]
    print('%s: parse=%s; menu live=%d, trong chu thich=%d; tim thay %d/%d ID' % (
        n, s['parse'], len(s['live']), len(s['comment']), len(hit), len(not_in)))
res = {}
for k, pts in by.items():
    found = []
    if k not in cur:
        for n, s in srcs.items():
            if n == 'menu.xml':
                continue
            for kind in ('live', 'comment'):
                for e in s[kind].get(k, []):
                    found.append(dict(e, file=n))
    ms = {'in_menu_xml': k in cur, 'found_in_backups': found}
    if k not in cur and not found:
        ms['ghi_chu'] = 'không rõ nguồn'
    res[k] = {'menu_source': ms, 'points': pts}
json.dump(res, open(OUT + '/by_menu.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
unk = [k for k in not_in if not res[k]['menu_source']['found_in_backups']]
print('Khong ro nguon (%d):' % len(unk), unk)
for k in not_in:
    f = res[k]['menu_source']['found_in_backups']
    if f:
        print(k, '|', '; '.join('%s%s %s MaFile=%s Ver=%s' % (x['file'], '(cmt)' if x['in_comment'] else '', x['ten_to_khai'] or x['caption'], x['MaFile'], x['VersionXML']) for x in f[:3]))
ab = json.load(open(OUT + '/app_branches.json', encoding='utf-8'))
ab['_header'] = {'luu_y_impact': 'Truong impact (loai tac dong) la GOI Y theo tu khoa, khong phai ket luan. Do dung tren mau doc tay doc lap (seed 43): 13/20 (65%). Mau seed 42: 17/20 nhung luat da chinh theo mau do.',
                 'menu_ids_ngoai_menu_xml': 'Xem by_menu.json[menuID].menu_source'}
json.dump(ab, open(OUT + '/app_branches.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
