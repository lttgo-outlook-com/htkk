#!/usr/bin/env python3
"""xsd_merge.py: gop va chuan hoa XSD HTKK 5.7.9 thanh JSON theo to khai + phien ban.

Chay:  python3 -I _spec/tools/xsd_merge.py [--src DIR] [--out DIR] [--forms DIR] [--menu FILE]
Mac dinh: src=/Users/tuanle/codes/HTKK/_extracted, out=<repo>/_spec/xsd, forms=<repo>/_forms.

Dac diem:
- Nhan dang tep schema theo noi dung (root la xs:schema), khong theo duoi tep (ten bi cat duoi).
- Giai xs:redefine / xs:include / xs:import theo ten tep (khong phan biet hoa thuong, '\\' = '/',
  '-' = '_'). Khong truy cap mang; chi doc tep trong src.
- Parse bang expat (khong tai DTD/entity ngoai), khong exec du lieu.
- Dau ra khong co thoi gian, dung thu tu co dinh: hai lan chay cho JSON giong tung byte.
"""
import hashlib
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from xml.parsers import expat

GENERATOR = "_spec/tools/xsd_merge.py"
VERSION = "1.0.0"
XS = "http://www.w3.org/2001/XMLSchema"
DS = "http://www.w3.org/2000/09/xmldsig#"
MAX_NAME = 28  # ten tep trong cabinet bi cat o ~28 ky tu

SIMPLE_FACETS = {
    "enumeration", "length", "minLength", "maxLength", "minInclusive", "maxInclusive",
    "minExclusive", "maxExclusive", "totalDigits", "fractionDigits", "whiteSpace",
}
TOP_KIND = {
    "simpleType": "type", "complexType": "type", "element": "element", "group": "group",
    "attributeGroup": "attributeGroup", "attribute": "attribute",
}


# ------------------------------------------------------------------ parse
class Node:
    __slots__ = ("tag", "attrs", "children", "line", "col", "file", "nsmap", "parent", "top")

    def __init__(self, tag, attrs, line, file, nsmap, parent):
        self.tag = tag
        self.attrs = attrs
        self.children = []
        self.line = line
        self.col = 0
        self.file = file
        self.nsmap = nsmap
        self.parent = parent
        self.top = None

    def loc(self):
        return "%s:%d" % (self.file, self.line)

    def uid(self):
        return "%s:%d:%d" % (self.file, self.line, self.col)


class Schema:
    def __init__(self, name, sha, root, raw_len):
        self.name = name
        self.sha = sha
        self.root = root
        self.raw_len = raw_len
        self.tns = root.attrs.get("targetNamespace", "")


def parse_schema(path, name):
    """Tra ve Schema neu root la xs:schema; None neu khong; nem Exception neu loi parse."""
    with open(path, "rb") as fh:
        data = fh.read()
    sha = hashlib.sha256(data).hexdigest()
    p = expat.ParserCreate()
    p.buffer_text = False
    stack = []
    holder = {"root": None, "skip": 0, "not_schema": False}

    def start(qname, attrs):
        if holder["skip"]:
            holder["skip"] += 1
            return
        parent = stack[-1] if stack else None
        nsmap = parent.nsmap if parent else {"xml": "http://www.w3.org/XML/1998/namespace"}
        changed = False
        for k, v in attrs.items():
            if k == "xmlns" or k.startswith("xmlns:"):
                if not changed:
                    nsmap = dict(nsmap)
                    changed = True
                nsmap["" if k == "xmlns" else k[6:]] = v
        if ":" in qname:
            pre, local = qname.split(":", 1)
        else:
            pre, local = "", qname
        ns = nsmap.get(pre, "")
        if parent is None and not (ns == XS and local == "schema"):
            holder["not_schema"] = True
        tag = local if ns == XS else "{foreign}" + local
        node = Node(tag, {k: v for k, v in attrs.items() if not (k == "xmlns" or k.startswith("xmlns:"))},
                    p.CurrentLineNumber, name, nsmap, parent)
        node.col = p.CurrentColumnNumber
        if parent is not None:
            parent.children.append(node)
            node.top = node if parent.tag in ("schema", "redefine") else parent.top
        else:
            holder["root"] = node
        if tag in ("annotation", "{foreign}appinfo", "{foreign}documentation"):
            holder["skip"] = 1
        stack.append(node)

    def end(qname):
        if holder["skip"]:
            holder["skip"] -= 1
            if holder["skip"] > 0:
                return
        stack.pop()

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.Parse(data, True)
    if holder["not_schema"] or holder["root"] is None:
        return None
    return Schema(name, sha, holder["root"], len(data))


def iter_nodes(node):
    stk = [node]
    while stk:
        n = stk.pop()
        yield n
        stk.extend(reversed(n.children))


# ------------------------------------------------------------------ scope
class Comp:
    __slots__ = ("kind", "ns", "name", "node", "file", "prev", "redefined")

    def __init__(self, kind, ns, name, node, file, prev=None, redefined=False):
        self.kind = kind
        self.ns = ns
        self.name = name
        self.node = node
        self.file = file
        self.prev = prev
        self.redefined = redefined


class Scope:
    def __init__(self):
        self.comps = {}
        self.sources = []      # [(file, role)]
        self.applied = []      # redefine da ap dung
        self.unres = []        # [(kind, value, loc)]

    def add_source(self, name, role):
        if all(n != name for n, _ in self.sources):
            self.sources.append((name, role))

    def merge(self, other, role):
        self.comps.update(other.comps)
        for n, r in other.sources:
            self.add_source(n, role if r == "root" else r)
        self.applied.extend(other.applied)
        self.unres.extend(other.unres)


def norm_loc(loc):
    b = loc.replace("\\", "/").split("/")[-1].lower().replace("-", "_")
    return b


class Repo:
    def __init__(self):
        self.schemas = {}   # name -> Schema
        self.by_norm = {}   # norm basename -> name
        self.cache = {}

    def add(self, sc):
        self.schemas[sc.name] = sc
        self.by_norm.setdefault(norm_loc(sc.name), sc.name)

    def find(self, loc):
        return self.by_norm.get(norm_loc(loc))

    def build_scope(self, name, stack=()):
        if name in self.cache:
            return self.cache[name]
        scope = Scope()
        scope.add_source(name, "root")
        if name in stack:
            scope.unres.append(("cycle", name, name))
            return scope
        sc = self.schemas[name]
        tns = sc.tns
        for child in sc.root.children:
            t = child.tag
            if t in ("include", "redefine", "import"):
                loc = child.attrs.get("schemaLocation")
                if not loc:
                    continue
                tgt = self.find(loc)
                if tgt is None:
                    scope.unres.append(("schemaLocation", loc, child.loc()))
                    continue
                scope.merge(self.build_scope(tgt, stack + (name,)), t)
                if t == "redefine":
                    for comp in child.children:
                        kind = TOP_KIND.get(comp.tag)
                        if kind is None or "name" not in comp.attrs:
                            continue
                        key = (kind, tns, comp.attrs["name"])
                        old = scope.comps.get(key)
                        if old is None:
                            scope.unres.append(("redefine_target", comp.attrs["name"], comp.loc()))
                        scope.comps[key] = Comp(kind, tns, comp.attrs["name"], comp, name, old, True)
                        scope.applied.append({
                            "block": child.loc(),
                            "block_uid": child.uid(),
                            "schemaLocation": loc,
                            "resolved_to": tgt,
                            "component": comp.attrs["name"],
                            "kind": comp.tag,
                            "derivation": derivation_of(comp),
                            "source": comp.loc(),
                            "original": old.node.loc() if old is not None else None,
                        })
            elif t in TOP_KIND and "name" in child.attrs:
                key = (TOP_KIND[t], tns, child.attrs["name"])
                scope.comps[key] = Comp(TOP_KIND[t], tns, child.attrs["name"], child, name)
        self.cache[name] = scope
        return scope


def derivation_of(comp):
    for n in iter_nodes(comp):
        if n.tag in ("extension", "restriction"):
            return "%s base=%s" % (n.tag, n.attrs.get("base", ""))
    return ""


# ------------------------------------------------------------------ resolve
def split_qname(node, q):
    """Tra ve (ns, local) hoac None neu tien to khong khai bao."""
    if ":" in q:
        pre, local = q.split(":", 1)
        if pre not in node.nsmap:
            return None
        return node.nsmap[pre], local
    return node.nsmap.get("", ""), q


class Expander:
    def __init__(self, scope, root_sc):
        self.scope = scope
        self.tns = root_sc.tns
        self.out = []
        self.unres = []       # (kind, value, loc)
        self.stack = []
        self.roots = []
        self.reach = {"pattern": set(), "use_required": set(), "minOccurs0": set()}

    # tra cuu thanh phan; kind = type/element/group/attributeGroup/attribute
    def lookup(self, kind, node, q, ctx=None, base=False):
        r = split_qname(node, q)
        if r is None:
            self.unres.append((kind, q, node.loc()))
            return None
        ns, local = r
        if ns == XS and kind == "type":
            return ("builtin", local)
        key = (kind, ns, local)
        if base and ctx is not None and ctx.prev is not None and key == (ctx.kind, ctx.ns, ctx.name):
            return ctx.prev
        comp = self.scope.comps.get(key)
        if comp is None:
            self.unres.append((kind, q, node.loc()))
        return comp

    # ---------------- kieu don gian
    def simple_of_node(self, snode, ctx, depth=0):
        """snode: node simpleType. Tra ve dict(builtin, patterns[(val,loc)], facets)."""
        res = {"builtin": "xs:anySimpleType", "patterns": [], "facets": {}}
        if depth > 30:
            return res
        for c in snode.children:
            if c.tag == "restriction":
                b = c.attrs.get("base")
                inner = next((x for x in c.children if x.tag == "simpleType"), None)
                if b is not None:
                    sub = self.simple_of_ref(c, b, ctx, depth + 1)
                elif inner is not None:
                    sub = self.simple_of_node(inner, ctx, depth + 1)
                else:
                    sub = {"builtin": "xs:anySimpleType", "patterns": [], "facets": {}}
                res["builtin"] = sub["builtin"]
                res["patterns"] = list(sub["patterns"])
                res["facets"] = {k: list(v) for k, v in sub["facets"].items()}
                for f in c.children:
                    if f.tag == "pattern":
                        res["patterns"].append((f.attrs.get("value", ""), f.loc(), f.uid()))
                    elif f.tag in SIMPLE_FACETS:
                        res["facets"].setdefault(f.tag, []).append(f.attrs.get("value", ""))
                return res
            if c.tag == "list":
                res["builtin"] = "list"
                return res
            if c.tag == "union":
                res["builtin"] = "union"
                return res
        return res

    def simple_of_ref(self, nsnode, q, ctx, depth=0):
        comp = self.lookup("type", nsnode, q, ctx, base=True)
        if comp is None:
            return {"builtin": "unresolved:" + q, "patterns": [], "facets": {}}
        if isinstance(comp, tuple):
            return {"builtin": "xs:" + comp[1], "patterns": [], "facets": {}}
        if comp.node.tag == "simpleType":
            return self.simple_of_node(comp.node, comp, depth)
        # complexType voi simpleContent
        return self.simple_of_complex(comp, depth)

    def simple_of_complex(self, comp, depth=0):
        if depth > 30:
            return {"builtin": "xs:anySimpleType", "patterns": [], "facets": {}}
        for c in comp.node.children:
            if c.tag == "simpleContent":
                for d in c.children:
                    if d.tag in ("extension", "restriction") and "base" in d.attrs:
                        sub = self.simple_of_ref(d, d.attrs["base"], comp, depth + 1)
                        sub = {"builtin": sub["builtin"], "patterns": list(sub["patterns"]),
                               "facets": {k: list(v) for k, v in sub["facets"].items()}}
                        for f in d.children:
                            if f.tag == "pattern":
                                sub["patterns"].append((f.attrs.get("value", ""), f.loc(), f.uid()))
                        return sub
        return {"builtin": "xs:anyType", "patterns": [], "facets": {}}

    # ---------------- phan tu
    def emit(self, e):
        self.out.append(e)
        return e

    def expand_element(self, node, path, ctx, req, minchain, rep):
        occ = node
        decl = node
        if "ref" in node.attrs:
            comp = self.lookup("element", node, node.attrs["ref"])
            if comp is None:
                r = split_qname(node, node.attrs["ref"])
                nm = r[1] if r else node.attrs["ref"]
                e = self.new_entry(path + "/" + nm, "element", occ)
                e["type"] = "unresolved"
                e["source"] = occ.loc()
                self.emit(e)
                return
            decl = comp.node
            ctx = comp
            ns = comp.ns
            nm = comp.name
            if ns == DS:
                nm = "ds:" + nm
        else:
            nm = node.attrs.get("name", "?")
        p = path + "/" + nm
        e = self.new_entry(p, "element", occ)
        e["source"] = decl.loc()
        if decl is not occ:
            e["ref_at"] = occ.loc()
        if decl.attrs.get("nillable") == "true":
            e["nillable"] = True
        for k in ("default", "fixed"):
            if k in decl.attrs:
                e[k] = decl.attrs[k]
        if decl.attrs.get("abstract") == "true":
            e["abstract"] = True
        mn = int(occ.attrs.get("minOccurs", "1")) if occ.attrs.get("minOccurs", "1").isdigit() else 1
        if occ.attrs.get("minOccurs") == "0":
            self.reach["minOccurs0"].add(occ.uid())
        e["required"] = bool(req and mn >= 1)
        if minchain:
            e["container_min0"] = list(minchain)
        mx = occ.attrs.get("maxOccurs", "1")
        r2 = rep or mx == "unbounded" or (mx.isdigit() and int(mx) > 1)
        if r2:
            e["repeatable"] = True
        self.emit(e)
        tq = decl.attrs.get("type")
        anon_c = next((c for c in decl.children if c.tag == "complexType"), None)
        anon_s = next((c for c in decl.children if c.tag == "simpleType"), None)
        if tq is not None:
            e["base"] = tq
            comp = self.lookup("type", decl, tq, ctx)
            if comp is None:
                e["type"] = "unresolved"
            elif isinstance(comp, tuple):
                e["type"] = "xs:" + comp[1]
            elif comp.node.tag == "simpleType":
                self.fill_simple(e, self.simple_of_node(comp.node, comp))
            else:
                self.fill_complex(e, comp.node, comp, p, req and mn >= 1, minchain, r2, comp)
        elif anon_c is not None:
            e["base"] = "anonymous"
            self.fill_complex(e, anon_c, ctx, p, req and mn >= 1, minchain, r2, None)
        elif anon_s is not None:
            e["base"] = "anonymous"
            r = next((c for c in anon_s.children if c.tag == "restriction"), None)
            if r is not None and "base" in r.attrs:
                e["base"] = r.attrs["base"]
            self.fill_simple(e, self.simple_of_node(anon_s, ctx))
        else:
            e["type"] = "xs:anyType"
            e["base"] = "xs:anyType"

    def fill_simple(self, e, info):
        e["type"] = info["builtin"]
        if info["patterns"]:
            e["pattern"] = [v for v, _, _ in info["patterns"]]
            e["pattern_src"] = [l for _, l, _ in info["patterns"]]
            for _, _, u in info["patterns"]:
                self.reach["pattern"].add(u)
        if info["facets"]:
            e["facets"] = info["facets"]

    def fill_complex(self, e, cnode, cctx, path, req, minchain, rep, comp):
        key = id(cnode)
        if key in self.stack:
            e["type"] = "complex"
            e["recursive"] = True
            return
        self.stack.append(key)
        try:
            e["type"] = "complex"
            sc = None
            for c in cnode.children:
                if c.tag == "simpleContent":
                    sc = c
            if sc is not None:
                comp_for = comp if comp is not None else Comp("type", "", "", cnode, cnode.file)
                info = self.simple_of_complex(comp_for)
                self.fill_simple(e, info)
                e["type"] = info["builtin"]
            self.walk_complex(cnode, cctx, path, req, minchain, rep)
        finally:
            self.stack.pop()

    def new_entry(self, path, kind, occ):
        e = {"path": path, "kind": kind}
        e["type"] = None
        e["base"] = None
        if kind == "element":
            e["minOccurs"] = occ.attrs.get("minOccurs", "1")
            e["maxOccurs"] = occ.attrs.get("maxOccurs", "1")
        return e

    # ---------------- noi dung complexType
    def walk_complex(self, cnode, ctx, path, req, minchain, rep):
        parts = []
        for c in cnode.children:
            if c.tag in ("complexContent", "simpleContent"):
                for d in c.children:
                    if d.tag in ("extension", "restriction"):
                        if c.tag == "complexContent" and d.tag == "extension" and "base" in d.attrs:
                            bc = self.lookup("type", d, d.attrs["base"], ctx, base=True)
                            if bc is not None and not isinstance(bc, tuple) and bc.node.tag == "complexType":
                                key = id(bc.node)
                                if key in self.stack:
                                    pass
                                else:
                                    self.stack.append(key)
                                    try:
                                        self.walk_complex(bc.node, bc, path, req, minchain, rep)
                                    finally:
                                        self.stack.pop()
                        elif c.tag == "simpleContent" and "base" in d.attrs:
                            bc = self.lookup("type", d, d.attrs["base"], ctx, base=True)
                            if bc is not None and not isinstance(bc, tuple) and bc.node.tag == "complexType":
                                key = id(bc.node)
                                if key not in self.stack:
                                    self.stack.append(key)
                                    try:
                                        self.walk_complex_attrs_only(bc.node, bc, path)
                                    finally:
                                        self.stack.pop()
                        self.particles(d.children, ctx, path, req, minchain, rep, False)
            else:
                parts.append(c)
        if parts:
            self.particles(parts, ctx, path, req, minchain, rep, False)

    def walk_complex_attrs_only(self, cnode, ctx, path):
        for c in cnode.children:
            if c.tag == "simpleContent":
                for d in c.children:
                    if d.tag in ("extension", "restriction"):
                        self.particles([x for x in d.children if x.tag in ("attribute", "attributeGroup")],
                                       ctx, path, True, [], False, False)

    def particles(self, nodes, ctx, path, req, minchain, rep, in_choice):
        for c in nodes:
            t = c.tag
            if t in ("sequence", "choice", "all"):
                mnv = c.attrs.get("minOccurs", "1")
                mn = int(mnv) if mnv.isdigit() else 1
                mxv = c.attrs.get("maxOccurs", "1")
                nrep = rep or mxv == "unbounded" or (mxv.isdigit() and int(mxv) > 1)
                nmc = minchain + [c.loc()] if mnv == "0" else minchain
                if mnv == "0":
                    self.reach["minOccurs0"].add(c.uid())
                kids = [x for x in c.children if x.tag in ("element", "group", "sequence", "choice", "any", "all")]
                nreq = req and mn >= 1 and not (t == "choice" and len(kids) > 1)
                self.particles(c.children, ctx, path, nreq, nmc, nrep, t == "choice")
            elif t == "element":
                self.expand_element(c, path, ctx, req, minchain, rep)
            elif t == "group":
                if "ref" in c.attrs:
                    g = self.lookup("group", c, c.attrs["ref"], ctx, base=False)
                    if g is not None:
                        mnv = c.attrs.get("minOccurs", "1")
                        mn = int(mnv) if mnv.isdigit() else 1
                        nmc = minchain + [c.loc()] if mnv == "0" else minchain
                        if mnv == "0":
                            self.reach["minOccurs0"].add(c.uid())
                        key = id(g.node)
                        if key not in self.stack:
                            self.stack.append(key)
                            try:
                                self.particles(g.node.children, g, path, req and mn >= 1, nmc, rep, False)
                            finally:
                                self.stack.pop()
            elif t == "any":
                e = self.new_entry(path + "/*", "any", c)
                e["type"] = "xs:anyType"
                e["base"] = "any"
                e["source"] = c.loc()
                e["namespace"] = c.attrs.get("namespace", "##any")
                e["required"] = False
                if c.attrs.get("minOccurs") == "0":
                    self.reach["minOccurs0"].add(c.uid())
                self.emit(e)
            elif t == "attribute":
                self.expand_attribute(c, path, ctx, req)
            elif t == "attributeGroup":
                if "ref" in c.attrs:
                    g = self.lookup("attributeGroup", c, c.attrs["ref"], ctx, base=False)
                    if g is not None:
                        key = id(g.node)
                        if key not in self.stack:
                            self.stack.append(key)
                            try:
                                self.particles(g.node.children, g, path, req, minchain, rep, False)
                            finally:
                                self.stack.pop()
            # anyAttribute, annotation: bo qua

    def expand_attribute(self, node, path, ctx, req=False):
        decl = node
        if "ref" in node.attrs:
            comp = self.lookup("attribute", node, node.attrs["ref"])
            if comp is None:
                r = split_qname(node, node.attrs["ref"])
                e = {"path": path + "/@" + (r[1] if r else node.attrs["ref"]), "kind": "attribute",
                     "type": "unresolved", "base": None, "source": node.loc()}
                self.emit(e)
                return
            decl = comp.node
            nm = comp.name
        else:
            nm = node.attrs.get("name", "?")
        e = {"path": path + "/@" + nm, "kind": "attribute", "type": None, "base": None}
        e["source"] = decl.loc()
        if decl is not node:
            e["ref_at"] = node.loc()
        use = node.attrs.get("use", decl.attrs.get("use", "optional"))
        e["use"] = use
        e["required"] = bool(req and use == "required")
        if use == "required":
            self.reach["use_required"].add(node.uid())
        for k in ("default", "fixed"):
            if k in node.attrs:
                e[k] = node.attrs[k]
            elif k in decl.attrs:
                e[k] = decl.attrs[k]
        self.emit(e)
        tq = decl.attrs.get("type")
        anon_s = next((c for c in decl.children if c.tag == "simpleType"), None)
        if tq is not None:
            e["base"] = tq
            self.fill_simple(e, self.simple_of_ref(decl, tq, None))
        elif anon_s is not None:
            e["base"] = "anonymous"
            r = next((c for c in anon_s.children if c.tag == "restriction"), None)
            if r is not None and "base" in r.attrs:
                e["base"] = r.attrs["base"]
            self.fill_simple(e, self.simple_of_node(anon_s, None))
        else:
            e["type"] = "xs:anySimpleType"
            e["base"] = "xs:anySimpleType"

    def run(self, root_names):
        for nm in root_names:
            comp = self.scope.comps[("element", self.tns, nm)]
            self.expand_element(comp.node, "", comp, True, [], False)


# ------------------------------------------------------------------ kiem tra tham chieu tinh
REF_ATTRS = (("type", "type"), ("base", "type"), ("itemType", "type"), ("substitutionGroup", "element"))


def static_check(repo, scope, out_unres):
    ex = Expander(scope, repo.schemas[scope.sources[0][0]])
    for fname, _ in scope.sources:
        sc = repo.schemas[fname]
        for n in iter_nodes(sc.root):
            tag = n.tag
            if tag in ("{foreign}x",):
                continue
            for a, kind in REF_ATTRS:
                if a in n.attrs:
                    top = n.top
                    if a == "base" and top is not None and top.parent is not None and top.parent.tag == "redefine" \
                            and n.attrs[a].split(":")[-1] == top.attrs.get("name"):
                        key = (TOP_KIND.get(top.tag, "type"), sc.tns, top.attrs.get("name"))
                        c = scope.comps.get(key)
                        if c is not None and c.prev is not None:
                            continue
                        out_unres.add(("redefine_base", n.attrs[a], n.loc()))
                        continue
                    if a == "memberTypes":
                        continue
                    r = ex.lookup(kind, n, n.attrs[a])
                    if r is None:
                        out_unres.add((kind, n.attrs[a], n.loc()))
            if "memberTypes" in n.attrs:
                for q in n.attrs["memberTypes"].split():
                    if ex.lookup("type", n, q) is None:
                        out_unres.add(("type", q, n.loc()))
            if "ref" in n.attrs and tag in ("element", "attribute", "group", "attributeGroup"):
                kind = tag
                if ex.lookup(kind, n, n.attrs["ref"]) is None:
                    out_unres.add((kind, n.attrs["ref"], n.loc()))
    for k, v, l in ex.unres:
        pass
    # ex.unres da ghi nhan o lookup() trong vong tren: dua vao out_unres
    for k, v, l in ex.unres:
        out_unres.add((k, v, l))


# ------------------------------------------------------------------ main
def sha_of(b):
    return hashlib.sha256(b).hexdigest()


def main(argv):
    here = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(os.path.dirname(here))
    opts = {"--src": "/Users/tuanle/codes/HTKK/_extracted", "--out": os.path.join(repo_root, "_spec", "xsd"),
            "--forms": os.path.join(repo_root, "_forms"), "--menu": None}
    i = 1
    while i < len(argv):
        if argv[i] in opts and i + 1 < len(argv):
            opts[argv[i]] = argv[i + 1]
            i += 2
        else:
            print("Tham so khong hop le:", argv[i])
            return 2
    src, out, forms = opts["--src"], opts["--out"], opts["--forms"]
    menu = opts["--menu"] or os.path.join(src, "menu.xml")

    # 1. doc tep
    names = sorted(f for f in os.listdir(src) if os.path.isfile(os.path.join(src, f)))
    repo = Repo()
    candidates = 0
    parse_fail = []
    n_xsd_ext = 0
    for f in names:
        path = os.path.join(src, f)
        if os.path.getsize(path) > 8 * 1024 * 1024:
            continue
        with open(path, "rb") as fh:
            head = fh.read(4096)
        if b"schema" not in head or b"XMLSchema" not in head:
            continue
        candidates += 1
        try:
            sc = parse_schema(path, f)
        except Exception as ex:  # noqa
            parse_fail.append((f, str(ex)))
            continue
        if sc is None:
            continue
        repo.add(sc)
        if f.lower().endswith(".xsd"):
            n_xsd_ext += 1
    all_names = sorted(repo.schemas)

    # 2. tep goc = tep khong bi tep khac tham chieu
    referenced = set()
    scopes = {}
    for nm in all_names:
        for n in iter_nodes(repo.schemas[nm].root):
            if n.tag in ("include", "redefine", "import") and n.attrs.get("schemaLocation"):
                t = repo.find(n.attrs["schemaLocation"])
                if t is not None:
                    referenced.add(t)
    roots = [n for n in all_names if n not in referenced]
    libs = [n for n in all_names if n in referenced]

    # 3. dem tren nguon (parse)
    def src_counts(names_):
        c = {"pattern": set(), "use_required": set(), "minOccurs0": set(), "redefine": set()}
        for nm in names_:
            for n in iter_nodes(repo.schemas[nm].root):
                if n.tag == "pattern":
                    c["pattern"].add(n.uid())
                if n.tag == "attribute" and n.attrs.get("use") == "required":
                    c["use_required"].add(n.uid())
                if n.attrs.get("minOccurs") == "0":
                    c["minOccurs0"].add(n.uid())
                if n.tag == "redefine":
                    c["redefine"].add(n.uid())
        return c

    src_all = src_counts(all_names)
    src_ext = src_counts([n for n in all_names if n.lower().endswith(".xsd")])
    raw_re = {"pattern": 0, "use_required": 0, "minOccurs0": 0, "redefine_tags": 0}
    for nm in all_names:
        if not nm.lower().endswith(".xsd"):
            continue
        with open(os.path.join(src, nm), "rb") as fh:
            t = fh.read().decode("utf-8", "replace")
        raw_re["pattern"] += len(re.findall(r"<xs:pattern\b", t))
        raw_re["use_required"] += len(re.findall(r'use="required"', t))
        raw_re["minOccurs0"] += len(re.findall(r'minOccurs="0"', t))
        raw_re["redefine_tags"] += len(re.findall(r"</?xs:redefine\b", t))

    # 4. menu
    mtree = ET.parse(menu)
    menus = []
    for m in mtree.getroot().iter("Menu"):
        if m.get("MaFile") and m.get("VersionXML"):
            menus.append({"MaFile": m.get("MaFile"), "VersionXML": m.get("VersionXML"),
                          "XsdSchema": m.get("XsdSchema") or m.get("MaFile"), "ID": m.get("ID"),
                          "VersionKHBS": m.get("VersionKHBS", "")})

    def match_file(full):
        if full in repo.schemas:
            return full
        for nm in all_names:
            if len(nm) >= MAX_NAME - 2 and nm not in referenced and full.startswith(nm) and \
                    not nm.lower().endswith(".xsd"):
                return nm
        return None

    file_keys = {}   # file -> list of menu dicts
    menu_missing = []
    for m in sorted(menus, key=lambda x: (x["MaFile"], x["VersionXML"], x["ID"])):
        s = m["XsdSchema"].lower()
        stems = [s]
        if s[0].isdigit():
            stems.append("_" + s[1:])
        stems.append("_" + s)
        ver = m["VersionXML"].replace(".", "")
        hit = None
        for st in stems:
            hit = match_file("%s_%s.xsd" % (st, ver))
            if hit:
                break
        if hit is None:
            menu_missing.append(m)
        else:
            file_keys.setdefault(hit, []).append(m)

    # _forms: sha -> (MaFile, ver, duong dan) de bo sung
    forms_map = {}
    if os.path.isdir(forms):
        for g in sorted(os.listdir(forms)):
            gp = os.path.join(forms, g)
            if not os.path.isdir(gp):
                continue
            for d in sorted(os.listdir(gp)):
                dp = os.path.join(gp, d)
                if not os.path.isdir(dp):
                    continue
                for f in sorted(os.listdir(dp)):
                    mo = re.match(r"^(.+)_v([0-9.]+)_xsd\.xsd$", f)
                    if mo:
                        with open(os.path.join(dp, f), "rb") as fh:
                            forms_map.setdefault(sha_of(fh.read()), []).append(
                                {"MaFile": mo.group(1), "VersionXML": mo.group(2), "path": "%s/%s/%s" % (g, d, f)})
    forms_any = {}
    if os.path.isdir(forms):
        for g in sorted(os.listdir(forms)):
            gp = os.path.join(forms, g)
            if not os.path.isdir(gp):
                continue
            for d in sorted(os.listdir(gp)):
                dp = os.path.join(gp, d)
                if not os.path.isdir(dp):
                    continue
                for f in sorted(os.listdir(dp)):
                    if f.lower().endswith(".xsd") or f.lower().endswith(".xs"):
                        with open(os.path.join(dp, f), "rb") as fh:
                            forms_any.setdefault(sha_of(fh.read()), []).append("%s/%s/%s" % (g, d, f))

    # khoa map cho tung tep goc: tier 1 = menu.xml, 2 = _forms (sha), 3 = suy ra tu ten tep, 4 = ten goc
    stem_table = {}
    for m in menus:
        for st in {m["XsdSchema"].lower(), m["MaFile"].lower()}:
            for variant in {st, "_" + st[1:] if st[0].isdigit() else st, "_" + st}:
                stem_table.setdefault(variant, set()).add((m["MaFile"], m["XsdSchema"]))
    mapping = {}
    for nm in roots:
        sc0 = repo.schemas[nm]
        keys = file_keys.get(nm, [])
        if keys:
            prim = sorted(keys, key=lambda x: (x["MaFile"] != x["XsdSchema"], x["MaFile"], x["VersionXML"], x["ID"]))[0]
            key = {"ma_file": prim["MaFile"], "version_xml": prim["VersionXML"], "xsd_schema": prim["XsdSchema"],
                   "menu_ids": sorted(set(k["ID"] for k in keys)),
                   "also": sorted(set("%s_%s" % (k["MaFile"], k["VersionXML"]) for k in keys
                                      if (k["MaFile"], k["VersionXML"]) != (prim["MaFile"], prim["VersionXML"]))),
                   "origin": "menu.xml"}
            mapping[nm] = (key, "%s_%s" % (prim["MaFile"], prim["VersionXML"]), 1)
            continue
        fm = forms_map.get(sc0.sha)
        if fm:
            prim = sorted(fm, key=lambda x: (x["MaFile"], x["VersionXML"]))[0]
            key = {"ma_file": prim["MaFile"], "version_xml": prim["VersionXML"], "xsd_schema": None,
                   "menu_ids": [], "also": [], "origin": "_forms"}
            mapping[nm] = (key, "%s_%s" % (prim["MaFile"], prim["VersionXML"]), 2)
            continue
        mo = re.match(r"^(.+)_(\d{3})\.xsd$", nm)
        if mo and not nm.startswith("khbs_") and mo.group(1).lower() in stem_table:
            cands = sorted(stem_table[mo.group(1).lower()], key=lambda x: (x[0] != x[1], x[0]))
            ver = ".".join(mo.group(2))
            key = {"ma_file": cands[0][0], "version_xml": ver, "xsd_schema": cands[0][1], "menu_ids": [],
                   "also": [], "origin": "suy_ra_tu_ten_tep (phien ban khong co trong menu.xml)"}
            mapping[nm] = (key, "%s_%s" % (cands[0][0], ver), 3)
            continue
        bn = re.sub(r"[^A-Za-z0-9_.-]", "_", nm)
        if bn.lower().endswith(".xsd"):
            bn = bn[:-4]
        mapping[nm] = (None, bn.rstrip("."), 4)
    roots = sorted(roots, key=lambda n: (mapping[n][2], n))

    # 5. sinh JSON
    os.makedirs(out, exist_ok=True)
    for old in os.listdir(out):
        if old.endswith(".json"):
            os.remove(os.path.join(out, old))
    used_names = {}
    index = []
    unres_all = set()
    reach = {"pattern": set(), "use_required": set(), "minOccurs0": set(), "redefine": set()}
    total_entries = 0
    fail_expand = []
    mapped_menu = 0
    mapped_forms = 0
    mapped_infer = 0
    for nm in roots:
        sc = repo.schemas[nm]
        scope = repo.build_scope(nm)
        tns = sc.tns
        # goc: element toan cuc trong tns khong bi ref
        refd = set()
        for fn, _ in scope.sources:
            for n in iter_nodes(repo.schemas[fn].root):
                if n.tag == "element" and "ref" in n.attrs:
                    r = split_qname(n, n.attrs["ref"])
                    if r:
                        refd.add(r)
        root_names = []
        for (kind, ns, name_), comp in scope.comps.items():
            if kind == "element" and ns == tns and (ns, name_) not in refd:
                root_names.append(name_)
        root_names.sort()
        ex = Expander(scope, sc)
        try:
            ex.run(root_names)
        except RecursionError:
            fail_expand.append(nm)
        static_unres = set()
        static_check(repo, scope, static_unres)
        for k, v, l in ex.unres:
            static_unres.add((k, v, l))
        for k, v, l in scope.unres:
            static_unres.add((k, v, l))
        unres_all |= static_unres

        key, base_name, tier = mapping[nm]
        if tier == 1:
            mapped_menu += 1
        elif tier == 2:
            mapped_forms += 1
        elif tier == 3:
            mapped_infer += 1
        jname = base_name
        if jname.lower() in used_names:
            jname = base_name + "__" + re.sub(r"[^A-Za-z0-9_]", "_", nm.rstrip("."))
        used_names[jname.lower()] = nm
        entries = ex.out
        total_entries += len(entries)
        for k in ("pattern", "use_required", "minOccurs0"):
            reach[k] |= ex.reach[k]
        for a in scope.applied:
            reach["redefine"].add(a["block_uid"])
        src_list = []
        for fn, role in scope.sources:
            src_list.append({"file": fn, "role": role, "sha256": repo.schemas[fn].sha})
        header = {"generator": GENERATOR, "version": VERSION, "sha256": sc.sha,
                  "source_files": src_list}
        stats = {
            "root_elements": root_names,
            "entries": len(entries),
            "elements": sum(1 for e in entries if e["kind"] == "element"),
            "attributes": sum(1 for e in entries if e["kind"] == "attribute"),
            "patterns": sum(len(e.get("pattern", [])) for e in entries),
        }
        doc_applied = [dict(a) for a in scope.applied]
        unres_list = sorted(["%s|%s|%s" % t for t in static_unres])
        fa = forms_any.get(sc.sha, [])
        parts = ['{\n"header":', json.dumps(header, ensure_ascii=False, separators=(",", ":")),
                 ',\n"key":', json.dumps(key, ensure_ascii=False, separators=(",", ":")),
                 ',\n"forms_copies":', json.dumps(fa, ensure_ascii=False, separators=(",", ":")),
                 ',\n"stats":', json.dumps(stats, ensure_ascii=False, separators=(",", ":")),
                 ',\n"redefines_applied":[\n',
                 ",\n".join(json.dumps(a, ensure_ascii=False, separators=(",", ":")) for a in doc_applied),
                 '\n],\n"unresolved_refs":[\n',
                 ",\n".join(json.dumps(u, ensure_ascii=False) for u in unres_list),
                 '\n],\n"elements":[\n',
                 ",\n".join(json.dumps(e, ensure_ascii=False, separators=(",", ":")) for e in entries),
                 "\n]\n}\n"]
        with open(os.path.join(out, jname + ".json"), "w", encoding="utf-8", newline="\n") as fh:
            fh.write("".join(parts))
        index.append({"json": jname + ".json", "source_file": nm, "sha256": sc.sha,
                      "key": ({"ma_file": key["ma_file"], "version_xml": key["version_xml"],
                               "menu_ids": key["menu_ids"], "origin": key["origin"]} if key else None),
                      "entries": len(entries), "redefines": len(doc_applied), "unresolved": len(unres_list)})

    # 6. doi chieu
    def cnt(s):
        return len(s)

    unreached = {k: sorted(src_all[k] - reach[k]) for k in src_all}
    over = {k: sorted(reach[k] - src_all[k]) for k in src_all}
    lib_set = set(libs)

    def by_file(locs):
        d = {}
        for l in locs:
            f = l.rsplit(":", 2)[0]
            d[f] = d.get(f, 0) + 1
        return d

    unres_sorted = sorted(unres_all)
    recon = {}
    for k in ("pattern", "use_required", "minOccurs0", "redefine"):
        un = by_file(unreached[k])
        un_lib = sum(v for f, v in un.items() if f in lib_set)
        recon[k] = {"nguon_tong_hop_noi_dung": cnt(src_all[k]), "nguon_chi_duoi_xsd": cnt(src_ext[k]),
                    "trong_json_vi_tri_duy_nhat": cnt(reach[k]), "khong_den_duoc": cnt(unreached[k]),
                    "khong_den_duoc_trong_tep_dung_chung": un_lib,
                    "khong_den_duoc_trong_tep_goc": cnt(unreached[k]) - un_lib,
                    "json_ngoai_nguon": cnt(over[k])}
    summary = {
        "tep_ung_vien": candidates, "tep_schema_doc_duoc": len(repo.schemas), "tep_duoi_xsd": n_xsd_ext,
        "tep_loi_parse": parse_fail, "tep_dung_chung": libs, "tep_goc": len(roots),
        "json_sinh_ra": len(index), "json_map_menu": mapped_menu, "json_map_forms": mapped_forms,
        "json_map_suy_ra_ten_tep": mapped_infer, "json_map_tong": mapped_menu + mapped_forms + mapped_infer, "tong_entry": total_entries,
        "menu_khong_tim_thay_xsd": ["%s_%s" % (m["MaFile"], m["VersionXML"]) for m in menu_missing],
        "menu_co_mafile_versionxml": len(menus),
        "nguon_regex_tren_tep_xsd": raw_re,
        "doi_chieu": recon,
        "tham_chieu_khong_giai_duoc": len(unres_sorted),
        "tham_chieu_khong_giai_duoc_mau": ["%s|%s|%s" % t for t in unres_sorted[:10]],
        "loi_expand": fail_expand,
    }
    idx = {"header": {"generator": GENERATOR, "version": VERSION}, "summary": {
        k: v for k, v in summary.items()}, "files": index,
        "unresolved_refs": ["%s|%s|%s" % t for t in unres_sorted],
        "unreached_by_kind": {k: by_file(unreached[k]) for k in unreached},
        "unreached_locations": {k: unreached[k] for k in unreached}}
    with open(os.path.join(out, "index.json"), "w", encoding="utf-8", newline="\n") as fh:
        json.dump(idx, fh, ensure_ascii=False, indent=1)
        fh.write("\n")

    print("== XSD MERGE ==")
    print("Tep ung vien (chua 'schema'+'XMLSchema'): %d; schema doc duoc: %d (duoi .xsd: %d); loi parse: %d"
          % (candidates, len(repo.schemas), n_xsd_ext, len(parse_fail)))
    print("Tep dung chung (bi redefine/include/import): %d; tep goc sinh JSON: %d" % (len(libs), len(roots)))
    print("Tong entry (phan tu+thuoc tinh) trong JSON: %d" % total_entries)
    print("Khoi nguon (parse, tat ca %d tep) | nguon (.xsd) | JSON (vi tri duy nhat) | khong den duoc (dung chung/goc)" % len(repo.schemas))
    for k in ("pattern", "use_required", "minOccurs0", "redefine"):
        r = recon[k]
        print("  %-13s %6d | %6d | %6d | %5d (%d/%d) | json ngoai nguon %d" % (
            k, r["nguon_tong_hop_noi_dung"], r["nguon_chi_duoi_xsd"], r["trong_json_vi_tri_duy_nhat"],
            r["khong_den_duoc"], r["khong_den_duoc_trong_tep_dung_chung"], r["khong_den_duoc_trong_tep_goc"],
            r["json_ngoai_nguon"]))
    print("Regex thuan tren tep .xsd: %s (redefine_tags = the mo + the dong)" % json.dumps(raw_re))
    print("Tham chieu khong giai duoc: %d" % len(unres_sorted))
    for u in summary["tham_chieu_khong_giai_duoc_mau"]:
        print("   ", u)
    print("JSON map menu.xml: %d; map them qua _forms (sha): %d; suy ra tu ten tep: %d; ten goc (khong map): %d"
          % (mapped_menu, mapped_forms, mapped_infer, len(index) - mapped_menu - mapped_forms - mapped_infer))
    print("Menu co MaFile+VersionXML: %d; khong tim thay XSD: %d" % (len(menus), len(menu_missing)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
