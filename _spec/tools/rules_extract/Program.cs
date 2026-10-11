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
                foreach (var v in fd.Declaration.Variables) { c.FieldInit[v.Identifier.Text] = v.Initializer?.Value; c.Members.Add(v.Identifier.Text); c.FieldType[v.Identifier.Text] = fd.Declaration.Type.ToString().Split('.').Last(); }
                break;
            case PropertyDeclarationSyntax pd:
                c.Members.Add(pd.Identifier.Text); c.FieldType[pd.Identifier.Text] = pd.Type.ToString().Split('.').Last();
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
var recStatus = new SortedDictionary<string, int>(StringComparer.Ordinal) { ["resolved"] = 0, ["parametric"] = 0, ["unresolved"] = 0, ["opaque_only"] = 0, ["ui_only"] = 0 };
var UiCalls = new HashSet<string>(StringComparer.Ordinal) { "SuspendLayout", "ResumeLayout", "SetActiveCell", "Focus", "ShowCell", "ShowActiveCell", "Invalidate", "Select", "BringToFront", "SetViewportTopRow", "SetViewportLeftColumn", "SetActiveSheet", "SetNodeBackGroundPink", "SetRowHeight", "SetColumnWidth", "Refresh", "Update", "PerformLayout", "ScrollTo", "SetFocus" };
int effStmts = 0, effMatched = 0, effDropped = 0, effRaw = 0, effKept = 0;
var stmtCounts = new SortedDictionary<string, int>(StringComparer.Ordinal); var stmtOk = new SortedDictionary<string, int>(StringComparer.Ordinal); var stmtDropped = new SortedDictionary<string, int>(StringComparer.Ordinal);
// method cùng class có ghi ô (trực tiếp hoặc qua method cùng class khác, không giới hạn độ sâu, chống vòng)
var writersByClass = new Dictionary<ClassCtx, HashSet<BaseMethodDeclarationSyntax>>(ReferenceEqualityComparer.Instance);
// chọn overload theo số đối số, cùng quy tắc với Interp
BaseMethodDeclarationSyntax Resolve(ClassCtx c, string name, int argc)
{
    var cands = c.Methods.GetValueOrDefault(name);
    if (cands == null) return null;
    return cands.FirstOrDefault(x => x.ParameterList.Parameters.Count == argc) ?? cands.FirstOrDefault(x => x.ParameterList.Parameters.Count >= argc && x.ParameterList.Parameters.Skip(argc).All(p => p.Default != null));
}
HashSet<BaseMethodDeclarationSyntax> Writers(ClassCtx c)
{
    if (writersByClass.TryGetValue(c, out var w)) return w;
    // điểm bất động trên đồ thị gọi trong class (theo từng overload): ghi trực tiếp, rồi thêm method gọi tới method đã ghi
    w = new HashSet<BaseMethodDeclarationSyntax>(ReferenceEqualityComparer.Instance);
    var callees = new List<(BaseMethodDeclarationSyntax md, List<BaseMethodDeclarationSyntax> cs)>();
    foreach (var (nm0, mds) in c.Methods.OrderBy(k => k.Key, StringComparer.Ordinal))
        foreach (var md in mds)
        {
            var cs = new List<BaseMethodDeclarationSyntax>();
            foreach (var n in md.DescendantNodes())
            {
                if (n is AssignmentExpressionSyntax a && a.Left is MemberAccessExpressionSyntax lm && lm.Expression is ElementAccessExpressionSyntax ea && ea.Expression is MemberAccessExpressionSyntax em && em.Name.Identifier.Text is "Cells" or "Rows" or "Columns") w.Add(md);
                if (n is InvocationExpressionSyntax iv)
                {
                    string nm = iv.Expression switch { IdentifierNameSyntax i => i.Identifier.Text, MemberAccessExpressionSyntax m => m.Name.Identifier.Text, _ => "" };
                    if (nm is "ChangeValueCell" or "LockUnLockCol" or "LockUnLockCols" or "changFormatCell" or "checkLockUnLockCells") w.Add(md);
                    if (iv.Expression is MemberAccessExpressionSyntax rm && nm.StartsWith("Reset", StringComparison.Ordinal) && rm.Expression is ElementAccessExpressionSyntax rea && rea.Expression is MemberAccessExpressionSyntax rem && rem.Name.Identifier.Text is "Cells" or "Rows" or "Columns") w.Add(md);
                    string cn = iv.Expression is IdentifierNameSyntax ci ? ci.Identifier.Text : iv.Expression is MemberAccessExpressionSyntax cm && cm.Expression is ThisExpressionSyntax ? cm.Name.Identifier.Text : null;
                    var t = cn == null ? null : Resolve(c, cn, iv.ArgumentList.Arguments.Count);
                    if (t != null) cs.Add(t);
                }
            }
            callees.Add((md, cs));
        }
    for (bool changed = true; changed;)
    {
        changed = false;
        foreach (var (md, cs) in callees) if (!w.Contains(md) && cs.Any(w.Contains)) { w.Add(md); changed = true; }
    }
    writersByClass[c] = w;
    return w;
}
// theo tên: có overload nào ghi ô (cách đếm của verify.py)
bool WritesCells(ClassCtx c, string name, HashSet<string> unused) => (c.Methods.GetValueOrDefault(name) ?? new()).Any(Writers(c).Contains);
var effOps = new SortedDictionary<string, int>(StringComparer.Ordinal);
var effStatus = new SortedDictionary<string, int>(StringComparer.Ordinal);
var truncated = new List<string>();
var classSummaries = new List<object>();
int totalRecords = 0;

// ---------- hiển thị ----------
string ColText(V c1) { var fl = new Flags(); return c1 is VConst k && k.Val is long n && n >= 1 ? Fmt.Letters(n) : "{col(" + R.Expr(c1, fl) + ")}"; }
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
        var rest = w == null ? null : Conj(w).Where(x => !keys.Contains(x.Key)).ToList();
        if (rest != null && rest.Count == 0) { alts.Add((null, Prune(val, facts))); break; } // điều kiện nhánh đã là dữ kiện: các nhánh sau không tới
        alts.Add((rest == null ? null : C.And(rest), Prune(val, facts)));
    }
    // gộp nhánh liền kề cùng giá trị
    var res = new List<(Cond, V)>();
    foreach (var a in alts)
    {
        if (res.Count > 0 && (res[^1].Item2.Key == a.Item2.Key || Mk.Strip(R.Expr(res[^1].Item2, new Flags())) == Mk.Strip(R.Expr(a.Item2, new Flags())))) { res[^1] = (res[^1].Item1 == null || a.Item1 == null ? null : C.Or(res[^1].Item1, a.Item1), a.Item2); continue; }
        res.Add(a);
    }
    if (res.Count == 0) return v;
    if (res.Count == 1 || res[0].Item1 == null) return res[0].Item2;
    res[^1] = (null, res[^1].Item2);
    return new VCases(res);
}

object ValueJson(V v, Flags fl, bool typed)
{
    if (v is VCases cs)
        return new Dictionary<string, object> { ["cases"] = cs.Alts.Select(a => new Dictionary<string, object> { ["when"] = a.When == null ? "else" : R.Cond(a.When, fl), ["value"] = ValueJson(a.Val, fl, typed) }).ToList() };
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

// ---------- ký hiệu tự do, điều kiện mâu thuẫn ----------
var tokRx = new System.Text.RegularExpressions.Regex(@"(?<![\w.$])[A-Za-z_]\w*(?![\w(])", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
var litRx = new System.Text.RegularExpressions.Regex("\"(?:\\\\.|[^\"\\\\])*\"", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
string Placeholders(string t)
{
    var sb = new System.Text.StringBuilder(); int d = 0;
    foreach (char ch in t) { if (ch == '{') { if (d > 0) sb.Append(ch); d++; } else if (ch == '}') { d--; if (d > 0) sb.Append(ch); else sb.Append(' '); } else if (d > 0) sb.Append(ch); }
    return sb.ToString();
}
void Tokens(object o, string k, HashSet<string> acc)
{
    switch (o)
    {
        case string t:
            if (k is "formula" or "value") t = Placeholders(t);
            if (k is "file" or "call" or "via" or "note" or "base" or "reason" or "prop") return;
            foreach (var nm in Mk.Names(t)) acc.Add(nm);
            t = Mk.Strip(t);
            t = litRx.Replace(t, "\"\"");
            t = System.Text.RegularExpressions.Regex.Replace(t, @"field\(\w+\)|obj\(\w*\)|opaque\([^()]*\)", "");
            foreach (System.Text.RegularExpressions.Match m in tokRx.Matches(t)) acc.Add(m.Value);
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(t, @"@(\d+)")) acc.Add("@" + m.Groups[1].Value);
            break;
        case System.Collections.IDictionary d: foreach (System.Collections.DictionaryEntry e in d) Tokens(e.Value, (string)e.Key, acc); break;
        case System.Collections.IEnumerable l: foreach (var x in l) Tokens(x, k, acc); break;
    }
}
var kindRx = new System.Text.RegularExpressions.Regex("opaque\\((call|result|branches)\\)\u001F\\d+:([A-Za-z_]\\w*)\u001F", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
void KindNames(string t, object o, SortedDictionary<string, string> acc)
{
    if (t != null) { foreach (System.Text.RegularExpressions.Match m in kindRx.Matches(t)) acc.TryAdd(m.Groups[2].Value, m.Groups[1].Value); return; }
    switch (o)
    {
        case string st: KindNames(st, null, acc); break;
        case System.Collections.IDictionary d: foreach (System.Collections.DictionaryEntry e in d) KindNames(null, e.Value, acc); break;
        case System.Collections.IEnumerable l: foreach (var x in l) KindNames(null, x, acc); break;
    }
}
bool Contradictory(List<Cond> conj)
{
    for (int i = 0; i < conj.Count; i++)
    {
        var others = conj.Where((_, j) => j != i).ToList();
        if (Infeasible(conj[i], others, new HashSet<string>(others.Select(x => x.Key), StringComparer.Ordinal))) return true;
    }
    return false;
}
int recUnreachable = 0, recWithSyms = 0, symEntries = 0;

List<Dictionary<string, object>> BuildRecords(List<Effect> effects, Dictionary<string, object> key, string hook, string kind, int sheet, string file, string ctxLabel, SortedDictionary<string, List<SymInfo>> syms)
{
    var groupConj = new Dictionary<string, List<List<Cond>>>(StringComparer.Ordinal);
    var groupNames = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
    var groupKinds = new Dictionary<string, SortedDictionary<string, string>>(StringComparer.Ordinal); // tên -> loại nhãn opaque (call/result/branches) // tên trong dấu ẩn của điều kiện (trước khi bỏ dấu)
    var groups = new List<(string gk, List<Effect> effs, List<Cond> trig, List<string> guards)>();
    var effSeen = new HashSet<string>(StringComparer.Ordinal);
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
        // nhóm theo văn bản hiển thị (bỏ dấu ẩn); cờ unreachable chỉ khi mọi effect của nhóm có đường đi mâu thuẫn
        var rawNames = guards.SelectMany(Mk.Names).ToList();
        foreach (var gtxt in guards) KindNames(gtxt, null, groupKinds.TryGetValue(string.Join("\u0001", trig.Select(t => t.Key)) + "\u0002" + string.Join("\u0001", guards.Select(Mk.Strip).Distinct(StringComparer.Ordinal)), out var gkk) ? gkk : (groupKinds[string.Join("\u0001", trig.Select(t => t.Key)) + "\u0002" + string.Join("\u0001", guards.Select(Mk.Strip).Distinct(StringComparer.Ordinal))] = new SortedDictionary<string, string>(StringComparer.Ordinal)));
        guards = guards.Select(Mk.Strip).Distinct(StringComparer.Ordinal).ToList();
        string gk = string.Join("\u0001", trig.Select(t => t.Key)) + "\u0002" + string.Join("\u0001", guards);
        if (!gidx.TryGetValue(gk, out var gi)) { gi = groups.Count; gidx[gk] = gi; groups.Add((gk, new List<Effect>(), trig, guards)); groupConj[gk] = new List<List<Cond>>(); groupNames[gk] = new HashSet<string>(StringComparer.Ordinal); }
        groupNames[gk].UnionWith(rawNames);
        effRaw++;
        string ek = gk + "\u0003" + ef.Op + "|" + Mk.Strip(ef.Target == null ? "" : R.Expr(ef.Target, new Flags())) + "|" + Mk.Strip(ef.Val == null ? "" : R.Expr(ef.Val, new Flags())) + "|" + ef.Line + "|" + ef.Via + "|" + Mk.Strip(JsonSerializer.Serialize(ef.Params, jsonOpt));
        groupConj[gk].Add(conj);
        if (!effSeen.Add(ek)) continue;
        effKept++;
        groups[gi].effs.Add(ef);
    }
    var recs = new List<Dictionary<string, object>>();
    foreach (var g in groups)
    {
        int worst = -1; bool anyOpaque = false;
        bool guardOpq = g.guards.Any(x => x.Contains("opaque(", StringComparison.Ordinal));
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
            else if (ef.Val != null) R.Expr(ef.Val, fl);
            if (ef.Params.Count > 0) ej["params"] = new SortedDictionary<string, object>(ef.Params, StringComparer.Ordinal);
            ej["at"] = ef.Line;
            if (ef.Via != null) ej["via"] = ef.Via;
            string stt;
            // trạng thái: phần chưa giải ở giá trị, tham số, đích hoặc điều kiện đường đi thì không được là resolved/parametric
            bool txtOpq = Mk.Strip(JsonSerializer.Serialize(ej, jsonOpt)).Contains("opaque(", StringComparison.Ordinal);
            bool cfgTarget = ef.Target == null && (ef.Op == "setButtonText" || (ef.Op == "lock" && ef.Params.ContainsKey("targets")));
            if (txtOpq) fl.Notes.Add("giá trị hoặc tham số có phần chưa giải");
            if (guardOpq) fl.Notes.Add("điều kiện đường đi có phần chưa giải");
            if (cfgTarget) fl.Notes.Add("ô đích chưa xác định trong phạm vi phân tích");
            if (ef.Op == "opaque") { stt = "opaque"; anyOpaque = true; notes.Add("gọi chưa giải: " + ef.Params["call"]); }
            else { stt = fl.Opaque || txtOpq || guardOpq || cfgTarget ? "unresolved" : fl.Sym ? "parametric" : "resolved"; worst = Math.Max(worst, Array.IndexOf(Rank, stt)); foreach (var n in fl.Notes) notes.Add(Mk.Strip(n)); }
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
        bool uiOnly = worst < 0 && g.effs.All(e => e.Op == "opaque" && UiCalls.Contains(((string)e.Params["call"]).Split('.').Last()));
        string status = worst < 0 ? (uiOnly ? "ui_only" : "opaque_only") : Rank[worst];
        string reason = worst < 0 ? (uiOnly ? "chỉ có thao tác giao diện, không tác dụng lên ô" : "chỉ có lời gọi chưa giải") : status == "resolved" ? (anyOpaque ? "có lời gọi chưa giải kèm theo" : "") : "";
        if ((status != "resolved" || anyOpaque) && !uiOnly)
        {
            var ns = notes.Where(n => status is "unresolved" or "opaque_only" ? true : !n.StartsWith("gọi chưa giải")).Take(8).ToList();
            reason = (reason.Length > 0 ? reason + " | " : "") + (status == "parametric" ? "placeholder: " : status == "unresolved" ? "chưa giải: " : "") + string.Join(", ", ns);
        }
        recStatus[status] = recStatus.GetValueOrDefault(status) + 1;
        bool unreach = groupConj[g.gk].All(Contradictory);
        if (unreach) { recUnreachable++; reason = "không tới được: điều kiện trên đường đi mâu thuẫn" + (reason.Length > 0 ? " | " + reason : ""); }
        var trigJ = BuildTrigger(g.trig, kind);
        var toks = new HashSet<string>(StringComparer.Ordinal);
        Tokens(g.guards, "guards", toks); Tokens(trigJ, "trigger", toks); Tokens(effJson, "effects", toks);
        toks.UnionWith(groupNames[g.gk]);
        var symJ = new List<object>();
        foreach (var t in toks.OrderBy(x => x, StringComparer.Ordinal))
            if (syms != null && syms.TryGetValue(t, out var sl0))
            {
                // biến lặp: chỉ giữ mục của vòng lặp được nhắc trong bản ghi (loop@dòng); không xác định thì giữ hết
                var sl = sl0.Where(x => !(x.Origin is "loop" or "foreach") || toks.Contains("@" + x.Line)).ToList();
                if (sl.Count == 0) sl = sl0;
                foreach (var si in sl.OrderBy(x => x.Line).ThenBy(x => x.Origin, StringComparer.Ordinal))
                    symJ.Add(new Dictionary<string, object> { ["name"] = si.Name, ["origin"] = si.Origin, ["at"] = si.Line, ["desc"] = si.Desc });
            }
        // tên bị ẩn khỏi nhãn opaque(call/result/branches) luôn có mục symbols tương ứng
        var kinds = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (groupKinds.TryGetValue(g.gk, out var gk0)) foreach (var kv in gk0) kinds[kv.Key] = kv.Value;
        KindNames(null, trigJ, kinds); KindNames(null, effJson, kinds);
        foreach (var (nm, kd) in kinds)
        {
            string org = kd == "call" ? "method" : kd == "result" ? "method-result" : "merge";
            if (!symJ.Any(x => x is Dictionary<string, object> sd && (string)sd["name"] == nm && (string)sd["origin"] == org))
                symJ.Add(new Dictionary<string, object> { ["name"] = nm, ["origin"] = org, ["at"] = lmin == int.MaxValue ? 0 : lmin, ["desc"] = "tên bị ẩn khỏi nhãn opaque(" + kd + ")" });
        }
        if (symJ.Count > 0) { recWithSyms++; symEntries += symJ.Count; }
        recs.Add(new Dictionary<string, object>
        {
            ["key"] = key,
            ["hook"] = hook,
            ["kind"] = kind,
            ["sheetIndex"] = sheet >= 0 ? sheet : null,
            ["trigger"] = trigJ,
            ["guards"] = g.guards,
            ["effects"] = effJson,
            ["source"] = new Dictionary<string, object> { ["file"] = file, ["lines"] = new[] { lmin, lmax } },
            ["status"] = status,
            ["reason"] = reason,
            ["unreachable"] = unreach,
            ["symbols"] = symJ,
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
    Interp.CurSheet = -1;
    it.RunMethod(m, cls, BindParams(m, -1, null), null, null, null);
    string name = Interp.MethodName(m);
    if (it.Truncated) truncated.Add(rel + ":" + name);
    return BuildRecords(it.Effects, key, name, "helper", -1, rel, "standalone", it.Syms);
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
        Interp.CurSheet = sheet;
        it.RunMethod(m, cls, BindParams(m, sheet, kind), null, null, null);
        hookCounts[kind]++;
        var effLines = new HashSet<int>(it.Effects.Where(e => e.Via == null).Select(e => e.Line));
        var effOpsAt = it.Effects.Where(e => e.Via == null).GroupBy(e => e.Line).ToDictionary(g => g.Key, g => new HashSet<string>(g.Select(e => e.Op + (e.Params.TryGetValue("via", out var vv) ? "/" + vv : "")), StringComparer.Ordinal));
        var dropped = new List<object>(); var seenL = new HashSet<string>(StringComparer.Ordinal);
        string Generic(SyntaxNode n) => n.Ancestors().Any(a => a is LambdaExpressionSyntax) ? "nằm trong lambda, không phân tích" : n.Ancestors().Any(a => a is CatchClauseSyntax) ? "nằm trong khối catch, không phân tích" : "không tới được: nhánh điều kiện gấp thành hằng sai, hoặc sau lệnh trả về/ném ngoại lệ";
        foreach (var n in m.DescendantNodes())
        {
            int l0 = 0; string cat = null;
            // (a) gán thuộc tính / Reset* trên Cells/Rows/Columns; (b) màu thẻ sheet; (c) chữ nút; (d) checkLockUnLockCells; (e) gọi method cùng class có ghi ô
            if (n is AssignmentExpressionSyntax asx && asx.IsKind(SyntaxKind.SimpleAssignmentExpression) && asx.Left is MemberAccessExpressionSyntax lm)
            {
                l0 = LineOf(asx);
                if (lm.Expression is ElementAccessExpressionSyntax rea && rea.Expression is MemberAccessExpressionSyntax rem && rem.Name.Identifier.Text is "Cells" or "Rows" or "Columns") cat = "cell";
                else if (lm.Expression is ElementAccessExpressionSyntax tea && tea.Expression is MemberAccessExpressionSyntax tm && tm.Name.Identifier.Text == "TabStrip") cat = "tab";
                else if (lm.Name.Identifier.Text == "Text" && lm.Expression is IdentifierNameSyntax bid && bid.Identifier.Text.Contains("ButtonCellType", StringComparison.OrdinalIgnoreCase)) cat = "button";
            }
            else if (n is InvocationExpressionSyntax ivx)
            {
                l0 = LineOf(ivx);
                if (ivx.Expression is MemberAccessExpressionSyntax im && im.Name.Identifier.Text.StartsWith("Reset", StringComparison.Ordinal) && ivx.ArgumentList.Arguments.Count == 0 && im.Expression is ElementAccessExpressionSyntax rea2 && rea2.Expression is MemberAccessExpressionSyntax rem2 && rem2.Name.Identifier.Text is "Cells" or "Rows" or "Columns") cat = "cell";
                else if (ivx.Expression is MemberAccessExpressionSyntax im2 && im2.Name.Identifier.Text == "checkLockUnLockCells") cat = "chk";
                else
                {
                    string cn = ivx.Expression is IdentifierNameSyntax cid ? cid.Identifier.Text : ivx.Expression is MemberAccessExpressionSyntax cma && cma.Expression is ThisExpressionSyntax ? cma.Name.Identifier.Text : null;
                    if (cn != null && cls.Methods.ContainsKey(cn) && WritesCells(cls, cn, new HashSet<string>(StringComparer.Ordinal))) cat = "call:" + cn;
                }
            }
            if (cat == null || !seenL.Add(cat + "@" + l0)) continue;
            if (cat == "cell") effStmts++; else stmtCounts[cat.StartsWith("call:") ? "call" : cat] = stmtCounts.GetValueOrDefault(cat.StartsWith("call:") ? "call" : cat) + 1;
            var ops = effOpsAt.GetValueOrDefault(l0) ?? new HashSet<string>();
            bool ok2 = cat switch
            {
                "cell" => effLines.Contains(l0),
                "tab" => ops.Contains("setTabColor") || ops.Contains("setTabProp"),
                "button" => ops.Contains("setButtonText"),
                "chk" => ops.Contains("lock/checkLockUnLockCells"),
                _ => it.Effects.Any(e => e.Op != "opaque" && e.Via != null && (e.Via == cat.Substring(5) + "@" + l0 || e.Via.StartsWith(cat.Substring(5) + "@" + l0 + " > ", StringComparison.Ordinal))),
            };
            if (ok2) { if (cat == "cell") effMatched++; else stmtOk[cat.StartsWith("call:") ? "call" : cat] = stmtOk.GetValueOrDefault(cat.StartsWith("call:") ? "call" : cat) + 1; continue; }
            if (cat == "cell") effDropped++;
            string why = it.Dropped.FirstOrDefault(d => d.Line == l0).Reason;
            if (why == null && cat.StartsWith("call:", StringComparison.Ordinal))
            {
                string cn = cat.Substring(5);
                var crs = it.Calls.Where(c => c.Depth == 1 && c.Line == l0 && c.Name == cn).ToList();
                var nf = crs.FirstOrDefault(c => !c.Followed);
                var ivc = (InvocationExpressionSyntax)n;
                var md2 = Resolve(cls, cn, ivc.ArgumentList.Arguments.Count);
                if (md2 != null && !Writers(cls).Contains(md2)) why = "overload " + cn + " được gọi (" + ivc.ArgumentList.Arguments.Count + " đối số, khai báo dòng " + LineOf(md2) + ") không ghi ô, chỉ overload khác cùng tên có ghi ô";
                else if (nf != null) why = "không nội tuyến " + cn + "@" + l0 + ": " + nf.Reason;
                else if (crs.Count > 0)
                {
                    var inner = new List<string>();
                    foreach (var c in crs)
                    {
                        for (int k = c.DropBefore; k < c.DropAfter; k++) inner.Add("dòng " + it.Dropped[k].Line + " " + it.Dropped[k].Reason);
                        foreach (var c2 in it.Calls.Where(x => !x.Followed && x.Via != null && x.Via.StartsWith(c.Via + " > ", StringComparison.Ordinal))) inner.Add("lời gọi " + c2.Name + " dòng " + c2.Line + " không nội tuyến: " + c2.Reason);
                    }
                    why = "đã nội tuyến " + cn + "@" + l0 + " nhưng không lệnh ghi ô nào tới được: " + (inner.Count > 0 ? string.Join("; ", inner.Distinct().Take(3)) : "nhánh ghi ô bị loại theo đối số truyền vào hoặc trả về sớm");
                }
            }
            why ??= Generic(n);
            var dj = new Dictionary<string, object> { ["at"] = l0, ["reason"] = why };
            if (cat != "cell") dj["kind"] = cat.StartsWith("call:", StringComparison.Ordinal) ? "call" : cat;
            if (cat.StartsWith("call:", StringComparison.Ordinal)) dj["call"] = cat.Substring(5);
            dropped.Add(dj);
            if (!cat.StartsWith("call:", StringComparison.Ordinal) && cat != "cell") stmtDropped[cat] = stmtDropped.GetValueOrDefault(cat) + 1;
            else if (cat != "cell") stmtDropped["call"] = stmtDropped.GetValueOrDefault("call") + 1;
        }
        var r = BuildRecords(it.Effects, key, name, kind, sheet, rel, "hook", it.Syms);
        if (it.Truncated) truncated.Add(rel + ":" + name);
        var span = m.GetLocation().GetLineSpan();
        hookList.Add(new Dictionary<string, object> { ["hook"] = name, ["kind"] = kind, ["sheetIndex"] = sheet >= 0 ? sheet : null, ["lines"] = new[] { span.StartLinePosition.Line + 1, span.EndLinePosition.Line + 1 }, ["records"] = r.Count, ["effects"] = it.Effects.Count, ["truncated"] = it.Truncated, ["dropped"] = dropped });
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
var markRx = new System.Text.RegularExpressions.Regex(@"\\u001[Ff][^""\\]*?\\u001[Ff]", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
void Write(string name, object o)
{
    var txt = markRx.Replace(JsonSerializer.Serialize(o, jsonOpt), "");
    if (txt.Contains("\\u001F", StringComparison.OrdinalIgnoreCase) || txt.Contains(Mk.M)) throw new Exception("còn dấu ẩn trong " + name);
    File.WriteAllText(Path.Combine(outDir, name), txt + "\n");
}
foreach (var (cname, pc) in perClassOut)
{
    totalRecords += pc.recs.Count;
    var sc = new SortedDictionary<string, int>(StringComparer.Ordinal);
    foreach (var r in pc.recs) sc[(string)r["status"]] = sc.GetValueOrDefault((string)r["status"]) + 1;
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
    ["effects_dedup"] = new Dictionary<string, object> { ["before"] = effRaw, ["after"] = effKept },
    ["hook_effect_statements"] = new Dictionary<string, object> { ["total"] = effStmts, ["with_effect"] = effMatched, ["dropped_flagged"] = effDropped },
    ["hook_other_statements"] = stmtCounts.ToDictionary(k => k.Key, k => (object)new Dictionary<string, object> { ["total"] = k.Value, ["with_effect"] = stmtOk.GetValueOrDefault(k.Key), ["dropped_flagged"] = stmtDropped.GetValueOrDefault(k.Key) }),
    ["effect_ops"] = effOps,
    ["records_unreachable"] = recUnreachable,
    ["records_with_symbols"] = new Dictionary<string, object> { ["records"] = recWithSyms, ["entries"] = symEntries },
    ["effect_status"] = effStatus,
    ["formula_points"] = new Dictionary<string, object> { ["total"] = ptList.Count, ["by_status"] = ptCounts, ["business"] = ptList.Count(p => ((string)p["file"]).StartsWith("tax.")), ["app"] = ptList.Count(p => ((string)p["file"]).StartsWith("htkk/")) },
    ["opaque_call_names"] = opq.Count,
    ["truncated"] = truncated,
    ["class_list"] = classSummaries,
});
Console.WriteLine($"classes={perClassOut.Count} hooks={string.Join(",", hookCounts.Select(k => k.Key + ":" + k.Value))} records={totalRecords} points={ptList.Count} {string.Join(",", ptCounts.Select(k => k.Key + ":" + k.Value))} truncated={truncated.Count}");
return 0;
