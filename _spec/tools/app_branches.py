#!/usr/bin/env python3
"""Kiem ke cac diem tang ung dung HTKK re nhanh theo menuID / MaFile / MaToKhai.

Chay:  python3 -I app_branches.py            -> sinh app_branches.json, by_menu.json, in thong ke
       python3 -I app_branches.py --sample N [--from K] -> in mau ngau nhien (seed 42) de doc tay
       python3 -I app_branches.py --accuracy  -> doc _spec/app/manual_check.json, in do dung
Chi doc _decompiled (khong sua). Khong chay ma nguon cua doi tuong phan tich.
"""
import glob
import json
import os
import random
import re
import sys
import collections
import xml.etree.ElementTree as ET

REPO = '/Users/tuanle/codes/HTKK'
HERE = os.path.dirname(os.path.abspath(__file__))
WT = os.path.dirname(os.path.dirname(HERE))  # goc worktree
OUT = os.path.join(WT, '_spec', 'app')
SCOPE = ['htkk', 'tax.htkk.utilities', 'tax.htkk.computehash']  # khong gom tax.htkk.business

ID_LAST = {'menuid': 'menu', 'strmenuid': 'menu', 'mafile': 'file', 'matokhai': 'tk'}
OPD = r'((?:[A-Za-z_]\w*(?:\(\))?\??\.)*[A-Za-z_]\w*)(?:\.ToString\(\))?'
LIT = r'"((?:[^"\\]|\\.)*)"'

IMPACTS = ['in', 'xuất XML', 'chuyển phiên bản', 'kiểm tra', 'nạp dữ liệu', 'hiển thị/UI', 'khác']
KW = [
    ('in', r'SpreadPrint|PrintOption|PrintSheet|PrintPreview|PrintCertainPage|CoolPrint|PrintDocument|PrintInfo|\bPrint\w*'),
    ('xuất XML', r'ExportXML|ExportHeaderXML|ExportTableF7XML|XmlDocument|CreateElement|SaveXml|SaveXML|XmlTextWriter|XmlWriter|WriteXml|\.Save\(|InnerXml|AppendChild'),
    ('chuyển phiên bản', r'VersionXML|VersionKHBS|ChuyenDoi|ChuyenPhienBan|ConvertVersion|UpdateVersion|NangCap|\bVersion\w*'),
    ('kiểm tra', r'Validate\w*|\bCheck\w+|MessageBox|ShowMessage|frmMessage|return false|IsValid|\bLoi\w*|Error'),
    ('nạp dữ liệu', r'\bLoad\w*|ReadXml|DataSet|DataTable|\.Fill\(|OpenFile|OpenXml|GetData\w*|\bSelectNodes?\b|SelectSingleNode|\bDocFile\w*|\bGetXml\w*|Catalogue\w*|\.Open\(|Dictionary|DanhMuc'),
    ('hiển thị/UI', r'\.Visible|\.Enabled|\.Text\b|\.Checked|Caption|\.Show\(|\.Hide\(|ReadOnly|\.Width|\.Height|\.Items|\)\.Location\b|\.Focus\(|\.Select\(|\.Rows|\.Columns|\.Cells|SetLocation|ForeColor|BackColor|Enable\w*|Visible\w*'),
]
KWC = [(n, re.compile(p)) for n, p in KW]
METHOD_FB = [
    ('in', re.compile(r'Print', re.I)),
    ('xuất XML', re.compile(r'Export|XML|Xml', re.I)),
    ('chuyển phiên bản', re.compile(r'Version|ChuyenDoi', re.I)),
    ('kiểm tra', re.compile(r'Validate|Check|Kiem', re.I)),
    ('nạp dữ liệu', re.compile(r'Load|Open|Read|Get', re.I)),
    ('hiển thị/UI', re.compile(r'Visible|Enable|Show|Hide|Display|Set|Init|Format|Click|Change', re.I)),
]
FILE_FB = [
    ('in', re.compile(r'Print', re.I)),
    ('kiểm tra', re.compile(r'Validate', re.I)),
    ('nạp dữ liệu', re.compile(r'OpenFile|Catalogue', re.I)),
    ('hiển thị/UI', re.compile(r'^frm|^Uc|Progress', re.I)),
]


def strip_str(s):
    s = re.sub(r'"(?:[^"\\]|\\.)*"', '""', s)
    s = re.sub(r"'(?:[^'\\]|\\.)*'", "''", s)
    return s


def paren_bal(s):
    t = strip_str(s)
    return t.count('(') - t.count(')')


def last_seg(expr):
    seg = expr.split('.')[-1]
    seg = re.sub(r'\(\)$', '', seg).rstrip('?')
    return seg


class MenuIndex:
    def __init__(self, path):
        r = ET.parse(path).getroot()
        self.ids = []
        self.by_file = collections.defaultdict(set)
        self.by_tk = collections.defaultdict(set)
        for e in r.iter('Menu'):
            mid = e.get('ID')
            self.ids.append(mid)
            if e.get('MaFile'):
                self.by_file[e.get('MaFile')].add(mid)
            if e.get('MaToKhai'):
                self.by_tk[e.get('MaToKhai')].add(mid)
            for v in e.iter('Validity'):
                if v.get('FolderName'):
                    self.by_file[v.get('FolderName')].add(mid)
        self.idset = set(self.ids)

    def resolve(self, kind, tok):
        if kind == 'menu':
            return ({tok} if tok in self.idset else set())
        if kind == 'file':
            return set(self.by_file.get(tok, ()))
        return set(self.by_tk.get(tok, ()))


def lit_tokens(s):
    return [t for t in re.split(r'[~,;#\s|]+', s) if t]


def read(f):
    return open(f, encoding='utf-8-sig', errors='replace').read()


def list_files():
    out = []
    for d in SCOPE:
        out += glob.glob(os.path.join(REPO, '_decompiled', d, '**', '*.cs'), recursive=True)
    return sorted(out)


def brace_match(lines, k):
    """lines[k] la '{' ; tra ve chi so dong '}' tuong ung."""
    depth = 0
    for i in range(k, len(lines)):
        t = strip_str(lines[i])
        depth += t.count('{') - t.count('}')
        if depth <= 0:
            return i
    return len(lines) - 1


STR_ASSIGN = r'\b{n}\s*(?:\[\s*\])?\s*=\s*(?:new\s*(?:string\s*\[\s*\d*\s*\]|List<string>|\[\])?\s*(?:\(\))?\s*\{{([^}}]*)\}}|{lit})'


def resolve_var(name, text_scope, depth=0):
    """Tra ve list token tu khai bao/gan cua bien name trong text_scope, hoac None."""
    toks = []
    found = False
    pat = re.compile(STR_ASSIGN.format(n=re.escape(name), lit=LIT))
    for m in pat.finditer(text_scope):
        found = True
        if m.group(1) is not None:
            for l in re.findall(LIT, m.group(1)):
                toks += lit_tokens(l)
        elif m.group(2) is not None:
            toks += lit_tokens(m.group(2))
    for m in re.finditer(r'\b' + re.escape(name) + r'\.Add\(\s*' + LIT, text_scope):
        found = True
        toks += lit_tokens(m.group(1))
    if not found and depth < 1:
        m = re.search(r'\b' + re.escape(name) + r'\s*=\s*([A-Za-z_]\w*)\s*;', text_scope)
        if m and m.group(1) != name:
            return resolve_var(m.group(1), text_scope, depth + 1)
    return toks if found else None


METHOD_STRONG = [
    ('in', re.compile(r'Print')),
    ('xuất XML', re.compile(r'Export|XML|Xml|FileName|SaveFile')),
    ('chuyển phiên bản', re.compile(r'Version|ChuyenDoi|NangCap')),
    ('kiểm tra', re.compile(r'Validat|Check(?!Box)|KiemTra|Valid')),
    ('nạp dữ liệu', re.compile(r'Import|Load|OpenFile|ReadFile|ReadXml|BindData|GetData|DocFile|Open|^DM[A-Z]')),
    ('hiển thị/UI', re.compile(r'Resize|Layout|SetVisible|ShowDetail|Hide|Display|InitControl|Init')),
]
FILE_STRONG = [
    ('khác', re.compile(r'^frmHTKKSystems')),  # dieu huong menu he thong
    ('in', re.compile(r'Print')),
    ('kiểm tra', re.compile(r'Validate')),
    ('nạp dữ liệu', re.compile(r'OpenFile')),
]


def classify(body_text, method, fname):
    for name, rx in METHOD_STRONG:
        m = rx.search(method or '')
        if m:
            return name, 'method:' + m.group(0)
    base = os.path.basename(fname)
    for name, rx in FILE_STRONG:
        m = rx.search(base)
        if m:
            return name, 'file:' + m.group(0)
    for name, rx in KWC:
        m = rx.search(body_text)
        if m:
            return name, 'body:' + m.group(0)
    for name, rx in METHOD_FB:
        m = rx.search(method or '')
        if m:
            return name, 'method:' + m.group(0)
    base = os.path.basename(fname)
    for name, rx in FILE_FB:
        m = rx.search(base)
        if m:
            return name, 'file:' + m.group(0)
    return 'khác', 'none'


def get_args(text, pos):
    """text[pos] == '(' ; tra ve (args_text, end_pos)."""
    depth = 0
    for i in range(pos, len(text)):
        c = text[i]
        if c == '(':
            depth += 1
        elif c == ')':
            depth -= 1
            if depth == 0:
                return text[pos + 1:i], i
    return text[pos + 1:], len(text)


class Scanner:
    def __init__(self, idx):
        self.idx = idx
        self.points = []
        self.dyn = []
        self.atom_counter = collections.Counter()
        self.direct_eq_atoms = collections.Counter()  # theo tep: so atom ==,!=,Equals voi toan hang ten truc tiep
        self.global_text = {}

    def scan_all(self, files):
        texts = {f: read(f) for f in files}
        self.global_text = texts
        self.all_text = '\n'.join(texts.values())
        for f in files:
            self.scan_file(f, texts[f])

    # ---------- phan tich 1 tep ----------
    def scan_file(self, f, text):
        rel = os.path.relpath(f, REPO + '/_decompiled')
        lines = text.split('\n')
        n = len(lines)
        # phat hien method
        methods = []  # (line_idx, name)
        for i, ln in enumerate(lines):
            m = re.match(r'^\t{1,2}(?:public|private|internal|protected)[^;=]*?(\w+)\s*\(', ln)
            if m:
                methods.append((i, m.group(1)))
        mstarts = [m[0] for m in methods]

        def method_at(i):
            lo, name, idx0 = 0, None, 0
            for k, (s, nm) in enumerate(methods):
                if s <= i:
                    name, idx0 = nm, k
                else:
                    break
            return name, idx0

        # alias theo method
        aliases = {}
        for k, (s, nm) in enumerate(methods):
            e = methods[k + 1][0] if k + 1 < len(methods) else n
            seg = '\n'.join(lines[s:e])
            al = {}
            for m in re.finditer(r'\b(?:string|var|object)\s+(\w+)\s*=\s*' + OPD + r'\s*;', seg):
                ls = last_seg(m.group(2)).lower()
                if ls in ID_LAST:
                    al[m.group(1)] = ID_LAST[ls]
            for m in re.finditer(r'\b(?:string|var|object)\s+(\w+)\s*=\s*' + OPD + r'\.ToString\(\)\s*;', seg):
                ls = last_seg(m.group(2)).lower()
                if ls in ID_LAST:
                    al[m.group(1)] = ID_LAST[ls]
            aliases[k] = al

        def kind_of(expr, al):
            ls = last_seg(expr)
            if '.' not in expr and ls in al:
                return al[ls]
            return ID_LAST.get(ls.lower())

        def mscope(i):
            name, k = method_at(i)
            if name is None:
                return '', {}, None
            s = methods[k][0]
            return '\n'.join(lines[s:i + 1]), aliases.get(k, {}), name

        # tracking switch
        switch_stack = []  # (kind, depth_open)
        depth = 0
        i = 0
        prev_case_group = None
        while i < n:
            ln = lines[i]
            st = ln.strip()
            tstr = strip_str(ln)
            # dynamic dispatch qua reflection
            mm = re.search(r'GetMethod\(\s*"([^"]+)"\s*\+', ln)
            if mm:
                meth, _k = method_at(i)[0], None
                self.dyn.append({'file': rel, 'line': i + 1, 'method': method_at(i)[0],
                                 'expr': st[:200], 'prefix': mm.group(1)})
            # switch tren toan hang id
            sm = re.match(r'switch\s*\((.*)\)\s*$', st)
            if sm:
                scope, al, mname = mscope(i)
                om = re.fullmatch(OPD, sm.group(1).strip())
                kd = kind_of(om.group(1), al) if om else None
                # '{' o dong ke
                switch_stack.append([kd, depth + 1, i])
            # case group
            cm = re.match(r'case\s+"((?:[^"\\]|\\.)*)"\s*:', st)
            if cm and switch_stack and switch_stack[-1][0] and depth == switch_stack[-1][1]:
                j = i
                labels = []
                while j < n:
                    c2 = re.match(r'case\s+"((?:[^"\\]|\\.)*)"\s*:', lines[j].strip())
                    if not c2:
                        break
                    labels.append(c2.group(1))
                    j += 1
                # than nhom: tu j den break/return
                body = []
                k2 = j
                while k2 < n and len(body) < 60:
                    s2 = lines[k2].strip()
                    if re.match(r'(case\s|default\s*:)', s2):
                        break
                    body.append(lines[k2])
                    if s2.startswith('break;') or s2.startswith('return') or s2.startswith('continue;') or s2.startswith('goto '):
                        break
                    if s2 == '}' and strip_str(lines[k2]).count('}') and depth and k2 > j and False:
                        break
                    k2 += 1
                kind = switch_stack[-1][0]
                atoms = [{'type': 'case', 'op': 'case', 'kind': kind, 'tokens': [l for l in labels], 'negated': False, 'unres': None}]
                scope, al, mname = mscope(i)
                self.add_point(rel, i + 1, mname, 'case ' + ', '.join('"%s"' % l for l in labels)[:300], atoms,
                               '\n'.join(body), 'case', lines)
                i = j
                for ll in lines[i - len(labels):i]:
                    pass
                # cap nhat depth cho cac dong da bo qua (khong co { })
                continue
            # dong logic (noi dong neu ngoac chua can bang)
            has_atom_hint = re.search(r'(?i)menuid|mafile|matokhai', ln)
            ln_end = i
            logical = ln
            if has_atom_hint or paren_bal(ln) > 0:
                bal = paren_bal(ln)
                while bal > 0 and ln_end + 1 < n and ln_end - i < 8:
                    ln_end += 1
                    logical += ' ' + lines[ln_end].strip()
                    bal = paren_bal(logical)
            # cap nhat depth + switch_stack sau khi xu ly dong
            if has_atom_hint or ln_end > i:
                self.process_logical(rel, lines, i, ln_end, logical, mscope, kind_of)
            for q in range(i, ln_end + 1):
                tq = strip_str(lines[q])
                depth += tq.count('{') - tq.count('}')
                while switch_stack and depth < switch_stack[-1][1] and q > switch_stack[-1][2]:
                    switch_stack.pop()
            i = ln_end + 1

    # ---------- xu ly 1 dong logic ----------
    def process_logical(self, rel, lines, i, j, logical, mscope, kind_of):
        st = logical.strip()
        if not re.search(r'(?i)menuid|mafile|matokhai', st):
            return
        scope, al, mname = mscope(i)
        atoms = []
        taken = []  # khoang da dung

        def free(a, b):
            return all(b <= x or a >= y for x, y in taken)

        def neg_before(pos):
            return bool(re.search(r'!\s*\(*\s*$', st[:pos]))

        # 1. so sanh == / !=
        for m in re.finditer(r'(?<![\w.])' + OPD + r'\s*(==|!=)\s*(?:' + LIT + r'|([A-Za-z_][\w.]*(?:\(\))?))', st):
            kd = kind_of(m.group(1), al)
            if not kd:
                continue
            op = m.group(2)
            if m.group(3) is not None:
                lit = m.group(3)
                if lit == '':
                    continue
                atoms.append({'type': 'cmp', 'op': op, 'kind': kd, 'tokens': [lit], 'negated': op == '!=', 'unres': None,
                              'direct': kd and '.' in m.group(1) or last_seg(m.group(1)).lower() in ID_LAST})
            else:
                rhs = m.group(4)
                if rhs in ('null', 'string.Empty', 'String.Empty', 'true', 'false'):
                    continue
                # bien hang cuc bo
                vals = resolve_var(rhs, scope) if '.' not in rhs else None
                if vals:
                    atoms.append({'type': 'cmp', 'op': op, 'kind': kd, 'tokens': vals, 'negated': op == '!=', 'unres': None, 'direct': False})
                else:
                    atoms.append({'type': 'cmp', 'op': op, 'kind': kd, 'tokens': [], 'negated': op == '!=', 'unres': 'toan hang khong phai hang: ' + rhs, 'direct': False})
            taken.append((m.start(), m.end()))
        for m in re.finditer(LIT + r'\s*(==|!=)\s*' + OPD, st):
            kd = kind_of(m.group(3), al)
            if not kd or m.group(1) == '':
                continue
            if not free(m.start(), m.end()):
                continue
            atoms.append({'type': 'cmp', 'op': m.group(2), 'kind': kd, 'tokens': [m.group(1)], 'negated': m.group(2) == '!=', 'unres': None,
                          'direct': last_seg(m.group(3)).lower() in ID_LAST})
            taken.append((m.start(), m.end()))
        # 2. Equals
        for m in re.finditer(r'(?<![\w.])' + OPD + r'\.Equals\(\s*(?:' + LIT + r'|([A-Za-z_][\w.]*(?:\(\))?))', st):
            kd = kind_of(m.group(1), al)
            if not kd:
                continue
            neg = neg_before(m.start())
            if m.group(2) is not None:
                if m.group(2) == '':
                    continue
                atoms.append({'type': 'cmp', 'op': 'Equals', 'kind': kd, 'tokens': [m.group(2)], 'negated': neg, 'unres': None,
                              'direct': last_seg(m.group(1)).lower() in ID_LAST})
            else:
                rhs = m.group(3)
                if rhs in ('null', 'string.Empty'):
                    continue
                vals = resolve_var(rhs, scope) if '.' not in rhs else None
                atoms.append({'type': 'cmp', 'op': 'Equals', 'kind': kd, 'tokens': vals or [], 'negated': neg,
                              'unres': None if vals else 'toan hang khong phai hang: ' + rhs, 'direct': False})
            taken.append((m.start(), m.end()))
        for m in re.finditer(LIT + r'\.Equals\(\s*' + OPD + r'\s*[,)]', st):
            kd = kind_of(m.group(2), al)
            if not kd or m.group(1) == '':
                continue
            atoms.append({'type': 'cmp', 'op': 'Equals', 'kind': kd, 'tokens': [m.group(1)], 'negated': neg_before(m.start()), 'unres': None,
                          'direct': last_seg(m.group(2)).lower() in ID_LAST})
            taken.append((m.start(), m.end()))
        for m in re.finditer(r'string\.Equals\(\s*' + OPD + r'\s*,\s*' + LIT, st):
            kd = kind_of(m.group(1), al)
            if kd:
                atoms.append({'type': 'cmp', 'op': 'Equals', 'kind': kd, 'tokens': [m.group(2)], 'negated': neg_before(m.start()), 'unres': None, 'direct': False})
        # 3. ID.StartsWith/EndsWith/Contains/IndexOf("lit")
        list_atoms = []
        for m in re.finditer(r'(?<![\w.])' + OPD + r'\.(StartsWith|EndsWith|Contains|IndexOf)\(\s*' + LIT, st):
            kd = kind_of(m.group(1), al)
            if not kd:
                continue
            atoms.append({'type': 'substr', 'op': m.group(2), 'kind': kd, 'tokens': [m.group(3)], 'negated': neg_before(m.start()), 'unres': None, 'direct': False})
        # 4. thanh vien danh sach: "lit".Contains(ARG co id) / "lit".IndexOf(...)
        for m in re.finditer(LIT + r'\.(Contains|IndexOf)\(', st):
            args, endp = get_args(st, m.end() - 1)
            kd = self.arg_kind(args, al, kind_of)
            if not kd:
                continue
            neg = neg_before(m.start())
            tail = st[endp + 1:endp + 12]
            if m.group(2) == 'IndexOf':
                if re.match(r'\s*(<\s*0|==\s*-1|<=\s*-1)', tail):
                    neg = True
                elif re.match(r'\s*(>=\s*0|>\s*-1|!=\s*-1)', tail):
                    neg = False
            atoms.append({'type': 'list', 'op': m.group(2), 'kind': kd, 'tokens': lit_tokens(m.group(1)), 'negated': neg, 'unres': None, 'direct': False,
                          'source': 'literal'})
        # 5. bien.Contains(id) / Array.IndexOf(bien, id) / new[]{..}.Contains(id)
        for m in re.finditer(r'(?<![\w."])((?:[A-Za-z_]\w*\.)*[A-Za-z_]\w*)\.(Contains|IndexOf)\(', st):
            recv = m.group(1)
            if kind_of(recv, al):
                continue
            if recv in ('string', 'String'):
                continue
            args, endp = get_args(st, m.end() - 1)
            kd = self.arg_kind(args, al, kind_of)
            if not kd:
                continue
            neg = neg_before(m.start())
            tail = st[endp + 1:endp + 12]
            if m.group(2) == 'IndexOf':
                if re.match(r'\s*(<\s*0|==\s*-1|<=\s*-1)', tail):
                    neg = True
                elif re.match(r'\s*(>=\s*0|>\s*-1|!=\s*-1)', tail):
                    neg = False
            name = recv.split('.')[-1]
            toks = resolve_var(name, scope)
            src = 'method'
            if toks is None:
                ft = self.cur_file_text(rel)
                if re.search(r'(?:private|public|protected|internal|static)[^;=\n(]*\b' + re.escape(name) + r'\b\s*=', ft):
                    toks = resolve_var(name, ft)
                    src = 'file'
            if toks is None:
                if re.search(r'(?:private|public|protected|internal|static)[^;=\n(]*\b' + re.escape(name) + r'\b\s*=', self.all_text):
                    toks = resolve_var(name, self.all_text)
                    src = 'global-field'
            atoms.append({'type': 'list', 'op': m.group(2), 'kind': kd, 'tokens': toks or [], 'negated': neg,
                          'unres': None if toks else 'danh sach %s khong giai duoc' % recv, 'direct': False, 'source': src})
        for m in re.finditer(r'Array\.IndexOf\(\s*([A-Za-z_][\w.]*)\s*,\s*([^)]*)\)', st):
            kd = self.arg_kind(m.group(2), al, kind_of)
            if not kd:
                continue
            name = m.group(1).split('.')[-1]
            toks = resolve_var(name, scope) or resolve_var(name, self.cur_file_text(rel)) or resolve_var(name, self.all_text)
            atoms.append({'type': 'list', 'op': 'Array.IndexOf', 'kind': kd, 'tokens': toks or [], 'negated': False,
                          'unres': None if toks else 'danh sach %s khong giai duoc' % m.group(1), 'direct': False, 'source': 'var'})
        # inline new[] { ... }.Contains(id)
        for m in re.finditer(r'new\s*(?:string\s*\[\s*\d*\s*\]|List<string>|\[\])?\s*(?:\(\))?\s*\{([^}]*)\}\s*\.Contains\(', st):
            args, endp = get_args(st, m.end() - 1)
            kd = self.arg_kind(args, al, kind_of)
            if kd:
                toks = []
                for l in re.findall(LIT, m.group(1)):
                    toks += lit_tokens(l)
                atoms.append({'type': 'list', 'op': 'Contains', 'kind': kd, 'tokens': toks, 'negated': neg_before(m.start()), 'unres': None, 'direct': False, 'source': 'inline'})
        if not atoms:
            return
        # than nhanh
        head = re.match(r'(else\s+if|if|while|foreach|for|return|switch)\b', st)
        site_kind = head.group(1).replace('  ', ' ') if head else ('expr')
        body = st
        if head and head.group(1) in ('if', 'else if', 'while', 'else  if') or (head and head.group(1).startswith('else')):
            k = j + 1
            while k < len(lines) and not lines[k].strip():
                k += 1
            if k < len(lines) and lines[k].strip() == '{':
                e = brace_match(lines, k)
                body = '\n'.join(lines[k + 1:min(e, k + 80)])
            elif k < len(lines):
                body = lines[k]
                kk = k
                while not body.rstrip().endswith((';', '}')) and kk + 1 < len(lines) and kk - k < 6:
                    kk += 1
                    body += '\n' + lines[kk]
        self.add_point(rel, i + 1, mname, st[:300], atoms, body, site_kind, lines)

    def arg_kind(self, args, al, kind_of):
        for m in re.finditer(r'(?<![\w.])((?:[A-Za-z_]\w*(?:\(\))?\??\.)*[A-Za-z_]\w*)', strip_str(args)):
            kd = kind_of(m.group(1), al)
            if kd:
                return kd
        return None

    def cur_file_text(self, rel):
        return self.global_text[os.path.join(REPO, '_decompiled', rel)]

    def add_point(self, rel, line, method, expr, atoms, body, site_kind, lines):
        ids_pos, ids_neg = set(), set()
        unres_operand = []
        lit_missing = []
        not_in_xml = []
        n_cmp = 0
        for a in atoms:
            if a['unres']:
                unres_operand.append(a['unres'])
            if a['type'] == 'cmp' or a['type'] == 'substr' or a['type'] == 'case':
                n_cmp += 1 if a['type'] != 'case' else 0
            for t in a['tokens']:
                if a['type'] == 'substr':
                    continue
                r = self.idx.resolve(a['kind'], t)
                if a['kind'] == 'menu':
                    # literal menuID giu nguyen du khong co trong menu.xml (menu cu/da go); danh dau rieng
                    r = {t}
                    if t not in self.idx.idset:
                        not_in_xml.append(t)
                if r:
                    (ids_neg if a['negated'] else ids_pos).update(r)
                else:
                    lit_missing.append('%s:%s' % (a['kind'], t))
        impact, why = classify(body, method, rel)
        pt = {
            'file': rel, 'line': line, 'method': method, 'expr': expr, 'site': site_kind,
            'atoms': [{k: v for k, v in a.items() if k != 'direct'} for a in atoms],
            'n_atomic': len(atoms),
            'menu_ids': sorted(ids_pos | ids_neg),
            'menu_ids_negated': sorted(ids_neg - ids_pos),
            'unresolved_operand': unres_operand,
            'literal_not_in_menu_xml': sorted(set(lit_missing)),
            'menu_ids_not_in_menu_xml': sorted(set(not_in_xml)),
            'only_substr': all(a['type'] == 'substr' for a in atoms),
            'impact': impact, 'impact_why': why,
            'body_head': body.strip()[:200],
        }
        pt['direct_eq_atoms'] = sum(1 for a in atoms if a['type'] == 'cmp' and a.get('direct'))
        self.points.append(pt)


def run():
    idx = MenuIndex(REPO + '/_extracted/menu.xml')
    sc = Scanner(idx)
    files = list_files()
    sc.scan_all(files)
    pts = sorted(sc.points, key=lambda p: (p['file'], p['line']))
    for k, p in enumerate(pts):
        p['id'] = k + 1
    os.makedirs(OUT, exist_ok=True)
    json.dump({'points': pts, 'dynamic_dispatch': sc.dyn}, open(os.path.join(OUT, 'app_branches.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    by = collections.defaultdict(list)
    for p in pts:
        for mid in p['menu_ids']:
            by[mid].append({'id': p['id'], 'file': p['file'], 'line': p['line'], 'method': p['method'],
                            'negated': mid in p['menu_ids_negated'], 'impact': p['impact']})
    json.dump({k: by[k] for k in sorted(by)}, open(os.path.join(OUT, 'by_menu.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    return idx, pts, sc, files, by


def report(idx, pts, sc, files, by):
    ops = collections.Counter()
    for p in pts:
        for a in p['atoms']:
            ops[a['type'] + ':' + a['op']] += 1
    eq_atoms = sum(v for k, v in ops.items() if k.startswith('cmp:'))
    list_sites = sum(1 for p in pts if any(a['type'] == 'list' for a in p['atoms']))
    case_groups = sum(1 for p in pts if p['site'] == 'case')
    case_labels = sum(len(a['tokens']) for p in pts for a in p['atoms'] if a['type'] == 'case')
    print('Tong diem re nhanh (site):', len(pts))
    print('Atom theo loai:', dict(ops))
    print('Atom so sanh ==,!=,Equals (so voi ~1105):', eq_atoms)
    print('Atom substr (StartsWith/Contains/IndexOf tren id):', sum(v for k, v in ops.items() if k.startswith('substr:')))
    print('Diem co danh sach thanh vien .Contains/IndexOf (so voi ~125):', list_sites)
    src = collections.Counter(a.get('source') for p in pts for a in p['atoms'] if a['type'] == 'list')
    print('Danh sach thanh vien theo nguon (atom):', dict(src))
    print('Nhom case:', case_groups, '; nhan case:', case_labels)
    print('Dynamic dispatch (GetMethod ten+chi so):', len(sc.dyn))
    c = collections.Counter(p['impact'] for p in pts)
    print('Theo tac dong:', {k: c.get(k, 0) for k in IMPACTS})
    inx = [m for m in by if m in idx.idset]
    print('So menuID co >=1 diem: %d (trong do %d co trong menu.xml / %d menu; %d ID chi co trong ma, khong co trong menu.xml)' % (len(by), len(inx), len(idx.ids), len(by) - len(inx)))
    print('ID chi co trong ma:', sorted(set(by) - idx.idset))
    nores = [p for p in pts if not p['menu_ids']]
    print('Diem khong co menuID nao giai duoc:', len(nores), '(trong do chi la so khop chuoi con:', sum(1 for p in nores if p['only_substr']), ')')
    print('Diem co ID literal khong nam trong menu.xml:', sum(1 for p in pts if p['menu_ids_not_in_menu_xml']))
    print('Diem co toan hang/danh sach khong giai duoc:', sum(1 for p in pts if p['unresolved_operand']))
    print('Diem co hang khong nam trong menu.xml:', sum(1 for p in pts if p['literal_not_in_menu_xml']))
    print('Diem co (khong giai duoc: khong menuID hoac toan hang khong giai duoc):', sum(1 for p in pts if (not p['menu_ids']) or p['unresolved_operand']))
    print('Diem phu dinh (!= / !):', sum(1 for p in pts if p['menu_ids_negated']))
    pf = collections.Counter(p['file'].split('/')[-1] for p in pts)
    print('Theo tep:', dict(pf.most_common()))
    # kiem cheo
    cross(sc, files, pts)


def cross(sc, files, pts):
    """Dem lai doc lap: moi tep, regex thuan tren van ban (khong tach dong logic, khong alias),
    chi dem toan hang ten truc tiep so voi hang chuoi."""
    names = r'(?:\bmenuID|\bMenuID|\bmenuId|\bstrMenuID|\bMaFile|\bmaFile|\bMaToKhai|\bmaToKhai)'
    rx1 = re.compile(names + r'(?:\.ToString\(\))?\s*(?:==|!=)\s*"[^"]+"')
    rx2 = re.compile(r'"[^"]+"\s*(?:==|!=)\s*(?:[\w.]*\.)?' + names + r'\b')
    rx3 = re.compile(names + r'(?:\.ToString\(\))?\.Equals\(\s*"[^"]+"')
    rx4 = re.compile(r'"[^"]+"\.Equals\(\s*(?:[\w.]*\.)?' + names + r'\s*[,)]')
    tot = 0
    perfile = collections.Counter()
    for f in files:
        t = read(f)
        # bo dong trong chuoi? giu nguyen
        c = len(rx1.findall(t)) + len(rx2.findall(t)) + len(rx3.findall(t)) + len(rx4.findall(t))
        if c:
            perfile[os.path.basename(f)] = c
        tot += c
    mine = collections.Counter()
    for p in pts:
        mine[os.path.basename(p['file'])] += p['direct_eq_atoms']
    mt = sum(mine.values())
    print('KIEM CHEO (dem regex doc lap tren van ban tho, toan hang ten truc tiep): %d ; script (atom direct): %d ; %s' % (tot, mt, 'KHOP' if tot == mt else 'LECH'))
    diff = {k: (perfile.get(k, 0), mine.get(k, 0)) for k in set(perfile) | set(mine) if perfile.get(k, 0) != mine.get(k, 0)}
    print('Tep lech (doc lap, script):', diff)


def sample(n, start=0, seed=42):
    d = json.load(open(os.path.join(OUT, 'app_branches.json'), encoding='utf-8'))
    pts = d['points']
    rnd = random.Random(seed)
    chosen = rnd.sample(range(len(pts)), 20)
    for ord_, k in enumerate(chosen[start:start + n], start + 1):
        p = pts[k]
        f = os.path.join(REPO, '_decompiled', p['file'])
        lines = read(f).split('\n')
        a = max(0, p['line'] - 1 - 6)
        b = min(len(lines), p['line'] - 1 + 22)
        print('=' * 70)
        print('#%d id=%d %s:%d method=%s' % (ord_, p['id'], p['file'], p['line'], p['method']))
        print('SCRIPT: impact=%s (%s) menu_ids=%s neg=%s unres=%s' % (p['impact'], p['impact_why'], p['menu_ids'][:12], p['menu_ids_negated'][:5], p['unresolved_operand']))
        for q in range(a, b):
            print('%5d| %s' % (q + 1, lines[q][:170]))


def accuracy():
    d = json.load(open(os.path.join(OUT, 'manual_check.json'), encoding='utf-8'))
    pts = {p['id']: p for p in json.load(open(os.path.join(OUT, 'app_branches.json'), encoding='utf-8'))['points']}
    for name, smp in d['samples'].items():
        items = smp['items']
        ok = 0
        bad = []
        for v in items:
            lab = pts[v['id']]['impact']
            if lab == v['tay']:
                ok += 1
            else:
                bad.append('#%s id=%s script=%s tay=%s' % (v['stt'], v['id'], lab, v['tay']))
        print('%s (%s): doc tay %d diem, dung %d, sai %d (%.0f%%)' % (name, smp['ghi'], len(items), ok, len(items) - ok, 100 * ok / len(items)))
        for b in bad:
            print('   SAI', b)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == '--sample':
        sample(int(a[1]), int(a[3]) if len(a) > 3 and a[2] == '--from' else 0, int(a[5]) if len(a) > 5 and a[4] == '--seed' else 42)
    elif a and a[0] == '--accuracy':
        accuracy()
    else:
        report(*run())
