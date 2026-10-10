// RulesExtract: trích logic hook HTKK (đặt công thức, đặt giá trị, khóa/mở ô, đổi định dạng) thành JSON.
// Chỉ đọc text .cs và gọi CSharpSyntaxTree.ParseText. Không biên dịch, không workspace, không chạy mã decompile.
// Dùng: RulesExtract <decompiledDir> <_spec/csharp> <outDir>
using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using RulesExtract;

if (args.Length < 3) { Console.Error.WriteLine("Dùng: RulesExtract <decompiledDir> <_spec/csharp> <outDir>"); return 2; }
string decDir = Path.GetFullPath(args[0]);
string idxDir = args[1];
string outDir = args[2];
Directory.CreateDirectory(outDir);
foreach (var old in Directory.EnumerateFiles(outDir, "*.json")) File.Delete(old);

var jsonOpt = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
string[] AppFiles = { "htkk/HTKK/WinDoWorkEvent.cs", "htkk/HTKK/WinSpreadProcess.cs", "htkk/HTKK/frmWinToKhai.cs" };
var HookKinds = new HashSet<string> { "CellChange", "AutoBindData", "Refresh", "UpdateFormula", "ClearData", "CellEditchange", "ButtonClick" };

// ---------- chỉ mục T0 ----------
using var classesDoc = JsonDocument.Parse(File.ReadAllText(Path.Combine(idxDir, "classes.json")));
using var indexDoc = JsonDocument.Parse(File.ReadAllText(Path.Combine(idxDir, "index.json")));
var classRows = classesDoc.RootElement.GetProperty("classes").EnumerateArray().OrderBy(c => c.GetProperty("class").GetString(), StringComparer.Ordinal).ToList();
var hooksByPath = new Dictionary<string, List<(string name, string kind, int sheet, int start)>>(StringComparer.Ordinal);
var t0Counts = new SortedDictionary<string, int>(StringComparer.Ordinal);
foreach (var fe in indexDoc.RootElement.GetProperty("files").EnumerateArray())
{
    string p = fe.GetProperty("path").GetString();
    var l = new List<(string, string, int, int)>();
    foreach (var m in fe.GetProperty("methods").EnumerateArray())
    {
        if (m.GetProperty("hook_kind").ValueKind != JsonValueKind.String) continue;
        string k = m.GetProperty("hook_kind").GetString();
        if (!HookKinds.Contains(k)) continue;
        int sh = m.GetProperty("sheet_index").ValueKind == JsonValueKind.Number ? m.GetProperty("sheet_index").GetInt32() : -1;
        l.Add((m.GetProperty("name").GetString(), k, sh, m.GetProperty("start_line").GetInt32()));
    }
    hooksByPath[p] = l;
}

// ---------- parse ----------
var trees = new SortedDictionary<string, SyntaxTree>(StringComparer.Ordinal);
var bizFiles = Directory.EnumerateFiles(Path.Combine(decDir, "tax.htkk.business"), "*.cs", SearchOption.AllDirectories)
    .Select(p => Path.GetRelativePath(decDir, p).Replace('\\', '/')).ToList();
foreach (var rel in bizFiles.Concat(AppFiles).Distinct().OrderBy(x => x, StringComparer.Ordinal))
    trees[rel] = CSharpSyntaxTree.ParseText(File.ReadAllText(Path.Combine(decDir, rel)), new CSharpParseOptions(LanguageVersion.Latest));

int LineOf(SyntaxNode n) => n.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
string PointId(string file, SyntaxNode n) { var p = n.GetLocation().GetLineSpan().StartLinePosition; return file + ":" + (p.Line + 1) + ":" + (p.Character + 1); }

ClassCtx BuildClass(TypeDeclarationSyntax td, string file)
{
    var c = new ClassCtx { Name = td.Identifier.Text, File = file };
    foreach (var mem in td.Members)
    {
        switch (mem)
        {
            case MethodDeclarationSyntax md:
                if (!c.Methods.TryGetValue(md.Identifier.Text, out var l)) c.Methods[md.Identifier.Text] = l = new();
                l.Add(md); break;
            case FieldDeclarationSyntax fd:
                foreach (var v in fd.Declaration.Variables) { c.FieldInit[v.Identifier.Text] = v.Initializer?.Value; c.Members.Add(v.Identifier.Text); }
                break;
            case PropertyDeclarationSyntax pd:
                c.Members.Add(pd.Identifier.Text);
                if (pd.ExpressionBody != null) c.PropBody[pd.Identifier.Text] = pd.ExpressionBody.Expression;
                else
                {
                    var get = pd.AccessorList?.Accessors.FirstOrDefault(a => a.IsKind(SyntaxKind.GetAccessorDeclaration));
                    var set = pd.AccessorList?.Accessors.FirstOrDefault(a => a.IsKind(SyntaxKind.SetAccessorDeclaration));
                    if (get != null && get.Body == null && get.ExpressionBody == null) { c.FieldInit[pd.Identifier.Text] = pd.Initializer?.Value; if (set != null) c.Mutable.Add(pd.Identifier.Text); }
                    else if (get?.ExpressionBody != null) c.PropBody[pd.Identifier.Text] = get.ExpressionBody.Expression;
                    else if (get?.Body != null) c.PropBody[pd.Identifier.Text] = get.Body;
                }
                break;
        }
    }
    // trường bị gán trong thân method/ctor/accessor: coi là giá trị lúc chạy
    foreach (var body in td.Members.Where(m => m is BaseMethodDeclarationSyntax || m is PropertyDeclarationSyntax))
        foreach (var n in body.DescendantNodes())
        {
            ExpressionSyntax left = n switch { AssignmentExpressionSyntax a => a.Left, PostfixUnaryExpressionSyntax pu => pu.Operand, PrefixUnaryExpressionSyntax pr when pr.IsKind(SyntaxKind.PreIncrementExpression) || pr.IsKind(SyntaxKind.PreDecrementExpression) => pr.Operand, _ => null };
            if (left == null) continue;
            string nm = left is IdentifierNameSyntax id ? id.Identifier.Text : left is MemberAccessExpressionSyntax ma && ma.Expression is ThisExpressionSyntax ? ma.Name.Identifier.Text : null;
            if (nm == null || !c.Members.Contains(nm)) continue;
            if (left is IdentifierNameSyntax && body is BaseMethodDeclarationSyntax bm && bm.DescendantNodes().OfType<VariableDeclaratorSyntax>().Any(v => v.Identifier.Text == nm)) continue;
            if (left is IdentifierNameSyntax && body is BaseMethodDeclarationSyntax bm2 && bm2.ParameterList.Parameters.Any(p => p.Identifier.Text == nm)) continue;
            c.Mutable.Add(nm);
        }
    return c;
}

// ---------- điểm .Formula = ----------
var points = new SortedDictionary<string, (string file, int line, int col, SyntaxNode node)>(StringComparer.Ordinal);
foreach (var (rel, tree) in trees)
    foreach (var a in tree.GetRoot().DescendantNodes().OfType<AssignmentExpressionSyntax>())
        if (a.IsKind(SyntaxKind.SimpleAssignmentExpression) && a.Left is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "Formula")
        {
            var p = a.GetLocation().GetLineSpan().StartLinePosition;
            points[PointId(rel, a)] = (rel, p.Line + 1, p.Character + 1, a);
        }
var pointStatus = new Dictionary<string, (int rank, string via, SortedSet<string> ctx)>(StringComparer.Ordinal);
string[] Rank = { "resolved", "parametric", "unresolved" };

var opaqueSites = new HashSet<string>(StringComparer.Ordinal);
var hookCounts = new SortedDictionary<string, int>(StringComparer.Ordinal);
foreach (var k in HookKinds.OrderBy(x => x, StringComparer.Ordinal)) hookCounts[k] = 0;
var recStatus = new SortedDictionary<string, int>(StringComparer.Ordinal) { ["resolved"] = 0, ["parametric"] = 0, ["unresolved"] = 0 };
var effOps = new SortedDictionary<string, int>(StringComparer.Ordinal);
var effStatus = new SortedDictionary<string, int>(StringComparer.Ordinal);
var truncated = new List<string>();
var classSummaries = new List<object>();
int totalRecords = 0;

// ---------- hiển thị ----------
string ColText(V c1) { var fl = new Flags(); return c1 is VConst k && k.Val is long n && n >= 1 ? Fmt.Letters(n) : "{col=" + R.Expr(c1, fl) + "}"; }
string RowText(V r1) { var fl = new Flags(); return r1 is VConst k && k.Val is long n ? n.ToString() : "{" + R.Expr(r1, fl) + "}"; }

// Tỉa nhánh cases mâu thuẫn với dữ kiện đường đi (atom == hằng, khoảng, tập giá trị)
List<Cond> Conj(Cond c) => c is CAnd a ? a.Xs : new List<Cond> { c };
bool Infeasible(Cond when, List<Cond> facts, HashSet<string> factKeys)
{
    foreach (var w in Conj(when))
    {
        if (factKeys.Contains(C.Not(w).Key)) return true;
        if (w is CCmp m && m.Op == "==" && m.A is VLin la && la.IsSingleAtom(out var at) && m.B is VConst kb && kb.Val is long kv)
            foreach (var f in facts)
            {
                if (f is CCmp fm && fm.A is VLin fa && fa.IsSingleAtom(out var fat) && fat == at && fm.B is VConst fk && fk.Val is long fv)
                {
                    if (fm.Op == "==" && fv != kv) return true;
                    if (fm.Op == "!=" && fv == kv) return true;
                    if ((fm.Op == ">=" && kv < fv) || (fm.Op == ">" && kv <= fv) || (fm.Op == "<=" && kv > fv) || (fm.Op == "<" && kv >= fv)) return true;
                }
                if (f is COr fo && fo.Xs.All(x => x is CCmp xm && xm.Op == "==" && xm.A is VLin xa && xa.IsSingleAtom(out var xat) && xat == at && xm.B is VConst xk && xk.Val is long))
                    if (!fo.Xs.Any(x => (long)((VConst)((CCmp)x).B).Val == kv)) return true;
            }
    }
    return false;
}
V Prune(V v, List<Cond> facts)
{
    if (v is not VCases cs) return v;
    var keys = new HashSet<string>(facts.Select(f => f.Key), StringComparer.Ordinal);
    var alts = new List<(Cond, V)>();
    foreach (var (w, val) in cs.Alts)
    {
        if (w != null && Infeasible(w, facts, keys)) continue;
        var w2 = w == null ? null : C.And(Conj(w).Where(x => !keys.Contains(x.Key)));
        alts.Add((w2, Prune(val, facts)));
    }
    // gộp nhánh liền kề cùng giá trị
    var res = new List<(Cond, V)>();
    foreach (var a in alts)
    {
        if (res.Count > 0 && res[^1].Item2.Key == a.Item2.Key) { res[^1] = (res[^1].Item1 == null || a.Item1 == null ? null : C.Or(res[^1].Item1, a.Item1), a.Item2); continue; }
        res.Add(a);
    }
    if (res.Count == 0) return v;
    if (res.Count == 1) return res[0].Item2;
    res[^1] = (null, res[^1].Item2);
    return new VCases(res);
}

object ValueJson(V v, Flags fl, bool typed)
{
    if (v is VCases cs)
        return new Dictionary<string, object> { ["cases"] = cs.Alts.Select(a => new Dictionary<string, object> { ["when"] = a.When == null ? "else" : R.Cond(a.When, new Flags()), ["value"] = ValueJson(a.Val, fl, typed) }).ToList() };
    if (typed && v is VConst k && (k.Val is bool || k.Val is long || k.Val is double || k.Val == null)) return k.Val;
    return R.Template(v, fl);
}

bool TriggerOnly(Cond c, out bool hasTrig)
{
    hasTrig = false;
    switch (c)
    {
        case CAnd a: { bool any = false; foreach (var x in a.Xs) { if (!TriggerOnly(x, out var h)) return false; any |= h; } hasTrig = any; return true; }
        case COr o: { bool any = false; foreach (var x in o.Xs) { if (!TriggerOnly(x, out var h)) return false; any |= h; } hasTrig = any; return true; }
        case CCmp m:
            {
                bool ta = m.A is VLin la && la.IsSingleAtom(out var at) && (at == R.TRow || at == R.TCol);
                var fl = new Flags(); R.Expr(m.B, fl);
                bool tb = fl.Notes.Contains(R.TRow) || fl.Notes.Contains(R.TCol);
                hasTrig = ta; return ta && !tb && !fl.Opaque;
            }
        default: return false;
    }
}
bool IsAtomCmp(Cond c, string atom, out string op, out V other)
{
    op = null; other = null;
    if (c is CCmp m && m.A is VLin la && la.IsSingleAtom(out var at) && at == atom) { op = m.Op; other = m.B; return true; }
    return false;
}

Dictionary<string, object> BuildTrigger(List<Cond> trig, string kind)
{
    var cells = new List<string>(); var seen = new HashSet<string>();
    void AddCell(string s) { if (seen.Add(s)) cells.Add(s); }
    List<V> rowsEq = null, colsEq = null; V lo = null, hi = null;
    foreach (var t in trig)
    {
        if (IsAtomCmp(t, R.TRow, out var op, out var o))
        {
            if (op == "==") rowsEq = new List<V> { o };
            else if (op == ">=") lo = o; else if (op == ">") lo = Ar.Add(o, new VConst(1L));
            else if (op == "<=") hi = o; else if (op == "<") hi = Ar.Add(o, new VConst(-1L));
        }
        else if (IsAtomCmp(t, R.TCol, out var op2, out var o2)) { if (op2 == "==") colsEq = new List<V> { o2 }; }
        else if (t is COr or)
        {
            if (or.Xs.All(x => IsAtomCmp(x, R.TRow, out var xo, out _) && xo == "==")) rowsEq = or.Xs.Select(x => ((CCmp)x).B).ToList();
            else if (or.Xs.All(x => IsAtomCmp(x, R.TCol, out var xo, out _) && xo == "==")) colsEq = or.Xs.Select(x => ((CCmp)x).B).ToList();
            else
                foreach (var x in or.Xs)
                    if (x is CAnd xa && xa.Xs.Count == 2)
                    {
                        V r = null, cc = null;
                        foreach (var y in xa.Xs) { if (IsAtomCmp(y, R.TRow, out var yo, out var yv) && yo == "==") r = yv; else if (IsAtomCmp(y, R.TCol, out var yo2, out var yv2) && yo2 == "==") cc = yv2; }
                        if (r != null && cc != null) AddCell(ColText(cc) + RowText(r));
                    }
        }
    }
    object rowRange = null;
    if (rowsEq != null && colsEq != null) foreach (var cc in colsEq) foreach (var r in rowsEq) AddCell(ColText(cc) + RowText(r));
    else if (rowsEq != null) foreach (var r in rowsEq) AddCell("*" + RowText(r));
    if (rowsEq == null && (lo != null || hi != null))
        rowRange = new Dictionary<string, object> { ["from"] = lo == null ? null : RowText(lo), ["to"] = hi == null ? null : RowText(hi), ["cols"] = colsEq?.Select(ColText).ToList() };
    else if (rowsEq == null && colsEq != null && cells.Count == 0)
        rowRange = new Dictionary<string, object> { ["from"] = null, ["to"] = null, ["cols"] = colsEq.Select(ColText).ToList() };
    var fl = new Flags();
    return new Dictionary<string, object>
    {
        ["cells"] = cells,
        ["rowRange"] = rowRange,
        ["base"] = kind is "CellChange" or "CellEditchange" ? "iRow/iCol 1-based (dòng/cột Excel)" : kind == "ButtonClick" ? "idxRow/idxCol 0-based, đã quy về dòng/cột Excel 1-based" : null,
        ["when"] = trig.Select(t => R.Cond(t, fl)).ToList(),
    };
}

List<Dictionary<string, object>> BuildRecords(List<Effect> effects, Dictionary<string, object> key, string hook, string kind, int sheet, string file, string ctxLabel)
{
    var groups = new List<(string gk, List<Effect> effs, List<Cond> trig, List<string> guards)>();
    var gidx = new Dictionary<string, int>(StringComparer.Ordinal);
    foreach (var ef in effects)
    {
        var conj = new List<Cond>();
        foreach (var p in ef.Path) { if (p.C is CConst) continue; conj.AddRange(p.C is CAnd a ? a.Xs : new List<Cond> { p.C }); }
        var trig = new List<Cond>(); var guards = new List<string>(); var fl = new Flags();
        var seenG = new HashSet<string>(StringComparer.Ordinal);
        foreach (var c in conj)
        {
            if (TriggerOnly(c, out var h) && h) { if (!trig.Any(t => t.Key == c.Key)) trig.Add(c); }
            else { var g = R.Cond(c, fl); if (seenG.Add(g)) guards.Add(g); }
        }
        foreach (var lp in ef.Loops) guards.Add("loop: " + lp.Desc);
        string gk = string.Join("\u0001", trig.Select(t => t.Key)) + "\u0002" + string.Join("\u0001", guards);
        if (!gidx.TryGetValue(gk, out var gi)) { gi = groups.Count; gidx[gk] = gi; groups.Add((gk, new List<Effect>(), trig, guards)); }
        groups[gi].effs.Add(ef);
    }
    var recs = new List<Dictionary<string, object>>();
    foreach (var g in groups)
    {
        int worst = -1; bool anyOpaque = false;
        var notes = new SortedSet<string>(StringComparer.Ordinal);
        var effJson = new List<object>();
        int lmin = int.MaxValue, lmax = 0;
        foreach (var ef in g.effs)
        {
            var fl = new Flags();
            var ej = new Dictionary<string, object> { ["op"] = ef.Op };
            if (ef.Target is VCell vc) ej["target"] = new Dictionary<string, object> { ["sheet"] = vc.S.Index is VConst sk ? sk.Val : (object)R.Expr(vc.S.Index, fl), ["cell"] = R.CellRef(vc, fl) };
            else if (ef.Target != null) { ej["target"] = null; fl.Opaque = true; fl.Notes.Add("đích không phải ô bảng tính (" + (ef.Target is VOpaque vo ? vo.What : "?") + ")"); }
            else ej["target"] = null;
            var facts = new List<Cond>();
            foreach (var p in ef.Path) if (p.C is not CConst) facts.AddRange(Conj(p.C));
            if (ef.Op == "setFormula") ej["formula"] = ValueJson(Prune(ef.Val, facts), fl, false);
            else if (ef.Op == "setValue") ej["value"] = ValueJson(Prune(ef.Val, facts), fl, true);
            else if (ef.Val != null && ef.Op is "lock") R.Expr(ef.Val, fl);
            if (ef.Params.Count > 0) ej["params"] = new SortedDictionary<string, object>(ef.Params, StringComparer.Ordinal);
            ej["at"] = ef.Line;
            if (ef.Via != null) ej["via"] = ef.Via;
            string stt;
            if (ef.Op == "opaque") { stt = "opaque"; anyOpaque = true; notes.Add("gọi chưa giải: " + ef.Params["call"]); }
            else { stt = fl.Opaque ? "unresolved" : fl.Sym ? "parametric" : "resolved"; worst = Math.Max(worst, Array.IndexOf(Rank, stt)); foreach (var n in fl.Notes) notes.Add(n); }
            ej["status"] = stt;
            effJson.Add(ej);
            effOps[ef.Op] = effOps.GetValueOrDefault(ef.Op) + 1;
            effStatus[stt] = effStatus.GetValueOrDefault(stt) + 1;
            if (ef.Point != null)
            {
                int rk = Array.IndexOf(Rank, stt);
                if (!pointStatus.TryGetValue(ef.Point, out var ps)) pointStatus[ef.Point] = ps = (rk, ctxLabel, new SortedSet<string>(StringComparer.Ordinal));
                if (rk < ps.rank) ps = (rk, ctxLabel, ps.ctx);
                ps.ctx.Add(key["Class"] + "." + hook);
                pointStatus[ef.Point] = ps;
            }
            foreach (var l in ef.Lines) { if (l < lmin) lmin = l; if (l > lmax) lmax = l; }
        }
        string status = worst < 0 ? "unresolved" : Rank[worst];
        string reason = worst < 0 ? "chỉ có lời gọi chưa giải" : status == "resolved" ? (anyOpaque ? "có lời gọi chưa giải kèm theo" : "") : "";
        if (status != "resolved" || anyOpaque)
        {
            var ns = notes.Where(n => status == "unresolved" ? true : !n.StartsWith("gọi chưa giải")).Take(8).ToList();
            reason = (reason.Length > 0 ? reason + " | " : "") + (status == "parametric" ? "placeholder: " : status == "unresolved" ? "chưa giải: " : "") + string.Join(", ", ns);
        }
        recStatus[status]++;
        recs.Add(new Dictionary<string, object>
        {
            ["key"] = key,
            ["hook"] = hook,
            ["kind"] = kind,
            ["sheetIndex"] = sheet >= 0 ? sheet : null,
            ["trigger"] = BuildTrigger(g.trig, kind),
            ["guards"] = g.guards,
            ["effects"] = effJson,
            ["source"] = new Dictionary<string, object> { ["file"] = file, ["lines"] = new[] { lmin, lmax } },
            ["status"] = status,
            ["reason"] = reason,
        });
    }
    return recs;
}

Dictionary<string, V> BindParams(BaseMethodDeclarationSyntax m, int sheet, string kind)
{
    var d = new Dictionary<string, V>(StringComparer.Ordinal);
    var ps = m.ParameterList.Parameters;
    for (int i = 0; i < ps.Count; i++) d[ps[i].Identifier.Text] = Interp.ParamValue(ps[i].Type?.ToString() ?? "", ps[i].Identifier.Text, sheet, kind, i, ps.Count);
    return d;
}

var reachedMethods = new HashSet<string>(StringComparer.Ordinal); // file:spanStart của method chứa điểm đã tới
List<Dictionary<string, object>> RunStandalone(string rel, ClassCtx cls, BaseMethodDeclarationSyntax m, Dictionary<string, object> key)
{
    var it = new Interp(rel, opaqueSites);
    it.RunMethod(m, cls, BindParams(m, -1, null), null, null, null);
    string name = Interp.MethodName(m);
    if (it.Truncated) truncated.Add(rel + ":" + name);
    return BuildRecords(it.Effects, key, name, "helper", -1, rel, "standalone");
}

// ---------- 394 class ----------
var classByFile = new Dictionary<string, Dictionary<string, ClassCtx>>(StringComparer.Ordinal);
Dictionary<string, ClassCtx> ClassesIn(string rel)
{
    if (classByFile.TryGetValue(rel, out var d)) return d;
    d = new Dictionary<string, ClassCtx>(StringComparer.Ordinal);
    foreach (var td in trees[rel].GetRoot().DescendantNodes().OfType<TypeDeclarationSyntax>())
        if (td.Parent is not TypeDeclarationSyntax && !d.ContainsKey(td.Identifier.Text)) d[td.Identifier.Text] = BuildClass(td, rel);
    classByFile[rel] = d;
    return d;
}
ClassCtx OwnerOf(string rel, SyntaxNode n)
{
    var td = n.Ancestors().OfType<TypeDeclarationSyntax>().LastOrDefault();
    return td == null ? null : ClassesIn(rel).GetValueOrDefault(td.Identifier.Text);
}

var perClassOut = new Dictionary<string, (Dictionary<string, object> key, string rel, List<object> hooks, List<Dictionary<string, object>> recs)>(StringComparer.Ordinal);
foreach (var row in classRows)
{
    string cname = row.GetProperty("class").GetString();
    string rel = row.GetProperty("path").GetString();
    var key = new Dictionary<string, object>
    {
        ["MenuID"] = row.GetProperty("MenuID").ValueKind == JsonValueKind.String ? row.GetProperty("MenuID").GetString() : null,
        ["MaFile"] = row.GetProperty("MaFile").ValueKind == JsonValueKind.String ? row.GetProperty("MaFile").GetString() : null,
        ["VersionXML"] = row.GetProperty("VersionXML").ValueKind == JsonValueKind.String ? row.GetProperty("VersionXML").GetString() : null,
        ["Class"] = cname,
        ["in_menu"] = row.GetProperty("in_menu").GetBoolean(),
    };
    var cls = ClassesIn(rel).GetValueOrDefault(cname) ?? ClassesIn(rel).Values.First();
    var hookList = new List<object>(); var recs = new List<Dictionary<string, object>>();
    foreach (var (name, kind, sheet, start) in hooksByPath.GetValueOrDefault(rel) ?? new())
    {
        var m = cls.Methods.GetValueOrDefault(name)?.FirstOrDefault(x => LineOf(x) <= start && x.GetLocation().GetLineSpan().EndLinePosition.Line + 1 >= start)
             ?? cls.Methods.GetValueOrDefault(name)?.FirstOrDefault();
        if (m == null) continue;
        var it = new Interp(rel, opaqueSites);
        it.RunMethod(m, cls, BindParams(m, sheet, kind), null, null, null);
        hookCounts[kind]++;
        var r = BuildRecords(it.Effects, key, name, kind, sheet, rel, "hook");
        if (it.Truncated) truncated.Add(rel + ":" + name);
        var span = m.GetLocation().GetLineSpan();
        hookList.Add(new Dictionary<string, object> { ["hook"] = name, ["kind"] = kind, ["sheetIndex"] = sheet >= 0 ? sheet : null, ["lines"] = new[] { span.StartLinePosition.Line + 1, span.EndLinePosition.Line + 1 }, ["records"] = r.Count, ["effects"] = it.Effects.Count, ["truncated"] = it.Truncated });
        recs.AddRange(r);
    }
    perClassOut[cname] = (key, rel, hookList, recs);
}

// ---------- điểm chưa tới: chạy độc lập method chứa nó ----------
var sharedRecs = new List<Dictionary<string, object>>(); var appRecs = new List<Dictionary<string, object>>();
var doneMethods = new HashSet<string>(StringComparer.Ordinal);
foreach (var (pid, pt) in points)
{
    if (pointStatus.ContainsKey(pid)) continue;
    var m = pt.node.Ancestors().OfType<BaseMethodDeclarationSyntax>().FirstOrDefault();
    if (m == null) continue;
    string mk = pt.file + ":" + m.SpanStart;
    if (!doneMethods.Add(mk)) continue;
    var cls = OwnerOf(pt.file, m);
    string cname = cls?.Name ?? "?";
    if (perClassOut.TryGetValue(cname, out var pc) && pc.rel == pt.file)
        pc.recs.AddRange(RunStandalone(pt.file, cls, m, pc.key));
    else
    {
        var key = new Dictionary<string, object> { ["MenuID"] = null, ["MaFile"] = null, ["VersionXML"] = null, ["Class"] = cname, ["in_menu"] = false };
        var r = RunStandalone(pt.file, cls, m, key);
        (pt.file.StartsWith("htkk/", StringComparison.Ordinal) ? appRecs : sharedRecs).AddRange(r);
    }
}
foreach (var (pid, pt) in points)
    if (!pointStatus.ContainsKey(pid)) pointStatus[pid] = (2, "none", new SortedSet<string>(StringComparer.Ordinal));

// ---------- ghi tệp ----------
void Write(string name, object o) => File.WriteAllText(Path.Combine(outDir, name), JsonSerializer.Serialize(o, jsonOpt) + "\n");
foreach (var (cname, pc) in perClassOut)
{
    totalRecords += pc.recs.Count;
    var sc = new SortedDictionary<string, int>(StringComparer.Ordinal) { ["resolved"] = 0, ["parametric"] = 0, ["unresolved"] = 0 };
    foreach (var r in pc.recs) sc[(string)r["status"]]++;
    Write(cname + ".json", new Dictionary<string, object> { ["key"] = pc.key, ["file"] = pc.rel, ["hooks"] = pc.hooks, ["records"] = pc.recs });
    classSummaries.Add(new Dictionary<string, object> { ["Class"] = cname, ["file"] = pc.rel, ["in_menu"] = pc.key["in_menu"], ["hooks"] = pc.hooks.Count, ["records"] = pc.recs.Count, ["status"] = sc });
}
totalRecords += sharedRecs.Count + appRecs.Count;
Write("_shared.json", new Dictionary<string, object> { ["note"] = "Method ngoài 394 class nghiệp vụ trong tax.htkk.business có điểm .Formula =, phân tích độc lập", ["records"] = sharedRecs });
Write("_app.json", new Dictionary<string, object> { ["note"] = "Method tầng app (htkk/HTKK) có điểm .Formula =, phân tích độc lập", ["records"] = appRecs });

var ptList = points.Select(kv => new Dictionary<string, object>
{
    ["file"] = kv.Value.file, ["line"] = kv.Value.line, ["col"] = kv.Value.col,
    ["status"] = Rank[pointStatus[kv.Key].rank], ["via"] = pointStatus[kv.Key].via,
    ["contexts"] = pointStatus[kv.Key].ctx.Take(5).ToList(), ["context_count"] = pointStatus[kv.Key].ctx.Count,
}).OrderBy(d => (string)d["file"], StringComparer.Ordinal).ThenBy(d => (int)d["line"]).ThenBy(d => (int)d["col"]).ToList();
Write("formula_points.json", ptList);

var opq = opaqueSites.Select(s => s.Split('\t')).GroupBy(p => p[0], StringComparer.Ordinal)
    .Select(g => new Dictionary<string, object>
    {
        ["call"] = g.Key,
        ["category"] = g.Key.StartsWith("GeneralFunctions.") || g.Key.StartsWith("HTKKFile.") || g.Key.StartsWith("ConvertHelpers.") || g.Key.StartsWith("HTKKFormater.") || g.Key.StartsWith("WinSettings.") ? "htkk-helper"
            : !g.Key.Contains('.') ? "class-local" : g.Key.StartsWith("<") ? "object-member" : "bcl",
        ["sites"] = g.Select(p => p[1]).Distinct().Count(),
        ["statement_sites"] = g.Where(p => p[2] == "statement").Select(p => p[1]).Distinct().Count(),
        ["value_sites"] = g.Where(p => p[2] == "value").Select(p => p[1]).Distinct().Count(),
    })
    .OrderByDescending(d => (int)d["sites"]).ThenBy(d => (string)d["call"], StringComparer.Ordinal).ToList();
Write("opaque_calls.json", opq);

var t0Biz = new SortedDictionary<string, int>(StringComparer.Ordinal);
foreach (var k in HookKinds.OrderBy(x => x, StringComparer.Ordinal)) t0Biz[k] = 0;
foreach (var row in classRows) foreach (var h in hooksByPath.GetValueOrDefault(row.GetProperty("path").GetString()) ?? new()) t0Biz[h.kind]++;
var ptCounts = new SortedDictionary<string, int>(StringComparer.Ordinal) { ["resolved"] = 0, ["parametric"] = 0, ["unresolved"] = 0 };
foreach (var p in ptList) ptCounts[(string)p["status"]]++;
truncated.Sort(StringComparer.Ordinal);
Write("index.json", new Dictionary<string, object>
{
    ["tool"] = "_spec/tools/rules_extract (Roslyn ParseText, không biên dịch)",
    ["classes"] = perClassOut.Count,
    ["hook_counts"] = hookCounts,
    ["hook_counts_T0_index_394_classes"] = t0Biz,
    ["records"] = totalRecords,
    ["record_status"] = recStatus,
    ["effect_ops"] = effOps,
    ["effect_status"] = effStatus,
    ["formula_points"] = new Dictionary<string, object> { ["total"] = ptList.Count, ["by_status"] = ptCounts, ["business"] = ptList.Count(p => ((string)p["file"]).StartsWith("tax.")), ["app"] = ptList.Count(p => ((string)p["file"]).StartsWith("htkk/")) },
    ["opaque_call_names"] = opq.Count,
    ["truncated"] = truncated,
    ["class_list"] = classSummaries,
});
Console.WriteLine($"classes={perClassOut.Count} hooks={string.Join(",", hookCounts.Select(k => k.Key + ":" + k.Value))} records={totalRecords} points={ptList.Count} {string.Join(",", ptCounts.Select(k => k.Key + ":" + k.Value))} truncated={truncated.Count}");
return 0;
