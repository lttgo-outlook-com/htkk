// Trích lệnh kiểm tra dữ liệu (đánh dấu ô Pink/Yellow, lời gọi HTKKFile.GetMessage) từ mã decompile HTKK.
// Chỉ đọc text .cs; dựng CSharpCompilation trong bộ nhớ theo lô (một tệp nghiệp vụ + tệp utilities nó tham chiếu, tối đa 50 tệp)
// để lấy SemanticModel và ControlFlowGraph tính reaching definitions. Không Emit, không chạy mã decompile.
// Dùng: RoslynValidations <decompiledDir> <out.json> [--log <batch_log.tsv>] [--only <Class>] [--max-batch <n>]
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

if (args.Length < 2) { Console.Error.WriteLine("Dùng: RoslynValidations <decompiledDir> <out.json> [--log f] [--only Class] [--max-batch n]"); return 2; }
string decDir = Path.GetFullPath(args[0]);
string outPath = args[1];
string logPath = null, onlyClass = null; int maxBatch = 50;
for (int ai = 2; ai + 1 < args.Length; ai += 2)
{
    if (args[ai] == "--log") logPath = args[ai + 1];
    else if (args[ai] == "--only") onlyClass = args[ai + 1];
    else if (args[ai] == "--max-batch") maxBatch = int.Parse(args[ai + 1]);
}

var files = new List<string>();
foreach (var sub in new[] { "tax.htkk.business/TAX.HTKK.Business", "htkk" })
{
    var d = Path.Combine(decDir, sub);
    if (!Directory.Exists(d)) continue;
    files.AddRange(Directory.EnumerateFiles(d, "*.cs", SearchOption.AllDirectories)
        .Select(p => Path.GetRelativePath(decDir, p).Replace('\\', '/')));
}
files.Sort(StringComparer.Ordinal);

var classes = new List<ClassCtx>();
foreach (var rel in files)
{
    string text = File.ReadAllText(Path.Combine(decDir, rel));
    var tree = CSharpSyntaxTree.ParseText(text, new CSharpParseOptions(LanguageVersion.Latest));
    var root = tree.GetCompilationUnitRoot();
    foreach (var td in root.DescendantNodes().OfType<ClassDeclarationSyntax>())
    {
        if (td.Parent is TypeDeclarationSyntax) continue;
        var cc = new ClassCtx { File = rel, Name = td.Identifier.Text, Decl = td, Assembly = rel.Split('/')[0], Tree = tree };
        foreach (var fd in td.Members.OfType<FieldDeclarationSyntax>())
            foreach (var v in fd.Declaration.Variables)
            {
                cc.FieldTypes[v.Identifier.Text] = fd.Declaration.Type.ToString();
                if (v.Initializer != null) cc.Fields[v.Identifier.Text] = v.Initializer.Value;
            }
        foreach (var pd in td.Members.OfType<PropertyDeclarationSyntax>()) cc.FieldTypes.TryAdd(pd.Identifier.Text, pd.Type.ToString());
        foreach (var pd in td.Members.OfType<PropertyDeclarationSyntax>())
        {
            if (pd.ExpressionBody != null) cc.Fields[pd.Identifier.Text] = pd.ExpressionBody.Expression;
            else if (pd.Initializer != null) cc.Fields[pd.Identifier.Text] = pd.Initializer.Value;
            else
            {
                var get = pd.AccessorList?.Accessors.FirstOrDefault(a => a.Keyword.IsKind(SyntaxKind.GetKeyword));
                var ret = get?.ExpressionBody?.Expression ?? (get?.Body?.Statements.Count == 1 && get.Body.Statements[0] is ReturnStatementSyntax rs ? rs.Expression : null);
                if (ret != null) cc.Fields[pd.Identifier.Text] = ret;
            }
        }
        foreach (var md in td.Members.OfType<MethodDeclarationSyntax>()) cc.Methods.Add(md);
        classes.Add(cc);
    }
}
var byName = new Dictionary<string, ClassCtx>(StringComparer.Ordinal);
foreach (var c in classes) byName.TryAdd(c.Name, c);

// ---- đồ thị gọi để xác định phạm vi ----
bool IsMarkName(string n) => n == "SetNodeBackGroundPink" || n == "SetNodeBackGroundPink2" || n == "SetNodeBackGroundYellow";
bool IsGM(InvocationExpressionSyntax inv) => inv.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "GetMessage" && ma.Expression.ToString() == "HTKKFile";
string CalleeName(InvocationExpressionSyntax inv) => inv.Expression switch
{
    MemberAccessExpressionSyntax ma => ma.Name.Identifier.Text,
    IdentifierNameSyntax id => id.Identifier.Text,
    GenericNameSyntax g => g.Identifier.Text,
    _ => null,
};
string Kind(string n)
{
    if (n.StartsWith("Validate", StringComparison.Ordinal)) return "Validate";
    if (n == "isValidDuLieuTK") return "isValidDuLieuTK";
    if (n.StartsWith("checkPL", StringComparison.OrdinalIgnoreCase)) return "checkPL";
    if (Regex.IsMatch(n, @"^UpdateFormula(_\d+)?$")) return "UpdateFormula";
    if (Regex.IsMatch(n, @"^ButtonClick(_\d+)?$")) return "ButtonClick";
    if (Regex.IsMatch(n, @"^(CellChange|AutoBindData|Refresh|CellEditchange|ClearData)(_\d+)?$")) return Regex.Match(n, @"^[A-Za-z]+").Value;
    return null;
}
var nodes = new Dictionary<string, (ClassCtx c, string m)>(StringComparer.Ordinal);
var edges = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
var hasSignal = new HashSet<string>(StringComparer.Ordinal);
foreach (var c in classes)
    foreach (var md in c.Methods)
    {
        string key = c.Name + "::" + md.Identifier.Text;
        nodes[key] = (c, md.Identifier.Text);
        if (!edges.TryGetValue(key, out var set)) edges[key] = set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var inv in md.DescendantNodes().OfType<InvocationExpressionSyntax>())
        {
            string n = CalleeName(inv); if (n == null) continue;
            if (IsMarkName(n) || IsGM(inv)) hasSignal.Add(key);
        }
        if (c.Assembly == "tax.htkk.business" && md.DescendantNodes().OfType<ObjectCreationExpressionSyntax>().Any(o => o.Type.ToString() == "frmMessage")) hasSignal.Add(key);
        foreach (var inv in md.DescendantNodes().OfType<InvocationExpressionSyntax>())
        {
            string n = CalleeName(inv); if (n == null) continue;
            string target = null;
            if (inv.Expression is IdentifierNameSyntax || (inv.Expression is MemberAccessExpressionSyntax m0 && m0.Expression is ThisExpressionSyntax)) target = c.Name;
            else if (inv.Expression is MemberAccessExpressionSyntax m1 && m1.Expression is IdentifierNameSyntax q && byName.ContainsKey(q.Identifier.Text)) target = q.Identifier.Text;
            if (target != null) set.Add(target + "::" + n);
        }
    }
var inScope = new Dictionary<string, string>(StringComparer.Ordinal); // key -> root key gần nhất
var queue = new Queue<string>();
foreach (var kv in nodes.OrderBy(k => k.Key, StringComparer.Ordinal))
{
    string k = Kind(kv.Value.m);
    bool root = k == "Validate" || k == "isValidDuLieuTK" || k == "checkPL" || ((k == "UpdateFormula" || k == "ButtonClick") && hasSignal.Contains(kv.Key));
    if (root) { inScope[kv.Key] = kv.Key; queue.Enqueue(kv.Key); }
}
while (queue.Count > 0)
{
    var k = queue.Dequeue();
    foreach (var t in edges[k].OrderBy(x => x, StringComparer.Ordinal))
        if (nodes.ContainsKey(t) && !inScope.ContainsKey(t)) { inScope[t] = inScope[k]; queue.Enqueue(t); }
}

// ---- lô biên dịch: tệp utilities (và 3 tệp dùng chung của business) theo tên kiểu khai báo ----
var batches = new Batches(decDir, classes, maxBatch);
var mon = new MemMonitor();
var log = new List<string> { "class\tfiles\tpeak_ws_mb\tmethods\tcfg_ok\tcfg_fail\tms" };
int totalCfgOk = 0, totalCfgFail = 0;
long globalPeak = 0;

// ---- xuất ----
var outClasses = new List<object>();
foreach (var c in classes.OrderBy(c => c.File, StringComparer.Ordinal).ThenBy(c => c.Name, StringComparer.Ordinal))
{
    if (onlyClass != null && c.Name != onlyClass) continue;
    var methods = new List<object>();
    var todo = c.Methods.Where(md => { string k0 = c.Name + "::" + md.Identifier.Text; string kd = Kind(md.Identifier.Text); return hasSignal.Contains(k0) || kd == "Validate" || kd == "isValidDuLieuTK" || kd == "checkPL"; }).ToList();
    if (todo.Count == 0) continue;
    var sw = Stopwatch.StartNew();
    mon.Reset();
    var (comp, nfiles) = batches.Compile(c);
    var model = comp.GetSemanticModel(c.Tree);
    int cfgOk = 0, cfgFail = 0;
    foreach (var md in todo)
    {
        string name = md.Identifier.Text;
        string key = c.Name + "::" + name;
        string kind = Kind(name);
        bool scope = inScope.ContainsKey(key);
        bool signal = hasSignal.Contains(key);
        RDResult rd = null;
        try { rd = FlowRD.Analyze(model, md); } catch (Exception ex) { Console.Error.WriteLine($"CFG lỗi {key}: {ex.GetType().Name}: {ex.Message}"); rd = null; }
        if (rd != null && rd.Overflow) rd = null;
        if (rd != null) cfgOk++; else cfgFail++;
        var mx = new MethodCtx(c, md, rd);
        var span = md.GetLocation().GetLineSpan();
        int bodyLines = 0;
        if (md.Body != null) { var bs = md.Body.GetLocation().GetLineSpan(); bodyLines = bs.EndLinePosition.Line - bs.StartLinePosition.Line + 1; }
        methods.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
        {
            ["method"] = name,
            ["params"] = string.Join(", ", md.ParameterList.Parameters.Select(p => p.Identifier.Text)),
            ["return_type"] = md.ReturnType.ToString(),
            ["start_line"] = span.StartLinePosition.Line + 1,
            ["end_line"] = span.EndLinePosition.Line + 1,
            ["body_lines"] = bodyLines,
            ["statement_count"] = md.Body?.Statements.Count ?? 0,
            ["kind"] = kind ?? (scope ? "helper" : "other"),
            ["in_scope"] = scope,
            ["reached_from"] = scope && inScope[key] != key ? inScope[key].Split("::")[1] : null,
            ["sheetIndex"] = mx.SheetIndex(),
            ["calls"] = edges[key].Where(nodes.ContainsKey).Select(t => t.StartsWith(c.Name + "::") ? t.Split("::")[1] : t.Replace("::", ".")).OrderBy(x => x, StringComparer.Ordinal).ToList(),
            ["checks"] = mx.Extract(),
            ["unused_messages"] = mx.Unused,
            ["binding_method"] = rd != null ? "cfg" : "fallback",
        });
    }
    model = null; comp = null;
    GC.Collect(2, GCCollectionMode.Forced, true, true);
    long pk = mon.Peak;
    globalPeak = Math.Max(globalPeak, pk);
    totalCfgOk += cfgOk; totalCfgFail += cfgFail;
    log.Add($"{c.Name}\t{nfiles}\t{pk / 1048576}\t{todo.Count}\t{cfgOk}\t{cfgFail}\t{sw.ElapsedMilliseconds}");
    if (pk > 3L * 1024 * 1024 * 1024) { batches.Shrink(); Console.Error.WriteLine($"RAM lô {c.Name} = {pk / 1048576} MB > 3 GB: giảm cỡ lô xuống {batches.Max}"); }
    if (methods.Count == 0) continue;
    outClasses.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
    {
        ["file"] = c.File, ["class"] = c.Name, ["assembly"] = c.Assembly, ["methods"] = methods,
    });
}
var jo = new JsonSerializerOptions { WriteIndented = false, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
File.WriteAllText(outPath, JsonSerializer.Serialize(new { classes = outClasses }, jo) + "\n", new UTF8Encoding(false));
mon.Stop();
if (logPath != null) File.WriteAllText(logPath, string.Join("\n", log) + "\n", new UTF8Encoding(false));
Console.WriteLine($"classes_out={outClasses.Count} cfg_ok={totalCfgOk} cfg_fail={totalCfgFail} peak_ws_mb={globalPeak / 1048576}");
return 0;

sealed class MemMonitor
{
    long peak;
    readonly Timer t;
    public MemMonitor() { t = new Timer(_ => Sample(), null, 0, 200); }
    void Sample() { long w = Environment.WorkingSet; long p; do { p = Interlocked.Read(ref peak); if (w <= p) return; } while (Interlocked.CompareExchange(ref peak, w, p) != p); }
    public void Reset() { Interlocked.Exchange(ref peak, Environment.WorkingSet); }
    public long Peak { get { Sample(); return Interlocked.Read(ref peak); } }
    public void Stop() => t.Dispose();
}

sealed class Batches
{
    readonly List<MetadataReference> refs = new();
    readonly Dictionary<string, List<SyntaxTree>> typeTrees = new(StringComparer.Ordinal);
    readonly Dictionary<SyntaxTree, string> treePath = new();
    public int Max;
    public Batches(string decDir, List<ClassCtx> classes, int max)
    {
        Max = max;
        string core = Path.GetDirectoryName(typeof(object).Assembly.Location);
        string ver = Path.GetFileName(core);
        string refDir = Path.GetFullPath(Path.Combine(core, "..", "..", "..", "packs", "Microsoft.NETCore.App.Ref", ver, "ref", "net8.0"));
        IEnumerable<string> dlls = Directory.Exists(refDir) ? Directory.EnumerateFiles(refDir, "*.dll") : ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")).Split(Path.PathSeparator);
        foreach (var d in dlls.OrderBy(x => x, StringComparer.Ordinal)) refs.Add(MetadataReference.CreateFromFile(d));
        var util = Path.Combine(decDir, "tax.htkk.utilities");
        var opts = new CSharpParseOptions(LanguageVersion.Latest);
        foreach (var f in Directory.EnumerateFiles(util, "*.cs", SearchOption.AllDirectories).OrderBy(x => x, StringComparer.Ordinal))
        {
            var tr = CSharpSyntaxTree.ParseText(File.ReadAllText(f), opts);
            treePath[tr] = Path.GetRelativePath(decDir, f);
            foreach (var td in tr.GetCompilationUnitRoot().DescendantNodes().OfType<BaseTypeDeclarationSyntax>())
                Add(td.Identifier.Text, tr);
        }
        foreach (var c in classes)
            if (c.Assembly == "tax.htkk.business" && (c.Name is "GeneralFunctions" or "ValidateBase" or "Common"))
            { treePath[c.Tree] = c.File; Add(c.Name, c.Tree); }
    }
    void Add(string n, SyntaxTree t) { if (!typeTrees.TryGetValue(n, out var l)) typeTrees[n] = l = new(); if (!l.Contains(t)) l.Add(t); }
    public void Shrink() { Max = Math.Max(1, Max / 2); }
    public (CSharpCompilation, int) Compile(ClassCtx c)
    {
        var cnt = new Dictionary<SyntaxTree, int>();
        foreach (var tk in c.Tree.GetRoot().DescendantTokens())
        {
            if (!tk.IsKind(SyntaxKind.IdentifierToken)) continue;
            if (!typeTrees.TryGetValue(tk.ValueText, out var l)) continue;
            foreach (var t in l) if (t != c.Tree) cnt[t] = cnt.GetValueOrDefault(t) + 1;
        }
        var pick = cnt.OrderByDescending(kv => kv.Value).ThenBy(kv => treePath[kv.Key], StringComparer.Ordinal).Take(Max - 1).Select(kv => kv.Key);
        var trees = new List<SyntaxTree> { c.Tree };
        trees.AddRange(pick);
        var comp = CSharpCompilation.Create("batch_" + c.Name, trees, refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, concurrentBuild: false, allowUnsafe: true));
        return (comp, trees.Count);
    }
}

class ClassCtx
{
    public string File, Name, Assembly;
    public ClassDeclarationSyntax Decl;
    public SyntaxTree Tree;
    public Dictionary<string, ExpressionSyntax> Fields = new(StringComparer.Ordinal);
    public Dictionary<string, string> FieldTypes = new(StringComparer.Ordinal);
    public List<MethodDeclarationSyntax> Methods = new();
}

class MethodCtx
{
    readonly ClassCtx C;
    readonly MethodDeclarationSyntax M;
    readonly List<(string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node)> assigns = new();
    readonly HashSet<string> forVars = new(StringComparer.Ordinal);
    readonly HashSet<string> paramNames;
    readonly RDResult rd;
    public MethodCtx(ClassCtx c, MethodDeclarationSyntax m, RDResult rd)
    {
        C = c; M = m; this.rd = rd;
        paramNames = m.ParameterList.Parameters.Select(p => p.Identifier.Text).ToHashSet(StringComparer.Ordinal);
        foreach (var n in m.DescendantNodes())
        {
            if (n is VariableDeclaratorSyntax vd && vd.Initializer != null)
            {
                assigns.Add((vd.Identifier.Text, vd.SpanStart, vd.Initializer.Value, Line(vd), "init", vd));
                if (vd.Parent?.Parent is ForStatementSyntax) forVars.Add(vd.Identifier.Text);
            }
            else if (n is AssignmentExpressionSyntax a && (a.Left is IdentifierNameSyntax || a.Left is ElementAccessExpressionSyntax))
                assigns.Add((Norm(a.Left.ToString()), a.SpanStart, a.Right, Line(a), a.IsKind(SyntaxKind.SimpleAssignmentExpression) ? "set" : a.OperatorToken.Text, a));
            else if (n is PostfixUnaryExpressionSyntax pu && (pu.IsKind(SyntaxKind.PostIncrementExpression) || pu.IsKind(SyntaxKind.PostDecrementExpression)) && (pu.Operand is IdentifierNameSyntax || pu.Operand is ElementAccessExpressionSyntax))
                assigns.Add((Norm(pu.Operand.ToString()), pu.SpanStart, null, Line(pu), pu.OperatorToken.Text, pu));
            else if (n is PrefixUnaryExpressionSyntax pr && (pr.IsKind(SyntaxKind.PreIncrementExpression) || pr.IsKind(SyntaxKind.PreDecrementExpression)) && (pr.Operand is IdentifierNameSyntax || pr.Operand is ElementAccessExpressionSyntax))
                assigns.Add((Norm(pr.Operand.ToString()), pr.SpanStart, null, Line(pr), pr.OperatorToken.Text, pr));
            else if (n is ArgumentSyntax arg && (arg.RefKindKeyword.IsKind(SyntaxKind.OutKeyword) || arg.RefKindKeyword.IsKind(SyntaxKind.RefKeyword)))
            {
                string k = arg.Expression is DeclarationExpressionSyntax de && de.Designation is SingleVariableDesignationSyntax sv ? sv.Identifier.Text : Norm(arg.Expression.ToString());
                assigns.Add((k, arg.SpanStart, null, Line(arg), arg.RefKindKeyword.IsKind(SyntaxKind.RefKeyword) ? "ref" : "out", arg));
            }
            else if (n is ForEachStatementSyntax fe)
                assigns.Add((fe.Identifier.Text, fe.SpanStart, null, Line(fe), "foreach", fe));
        }
    }
    static int Line(SyntaxNode n) => n.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
    static string Norm(string s) => Regex.Replace(s, @"\s+", "");

    static bool IsLoop(SyntaxNode p) => p is ForStatementSyntax || p is WhileStatementSyntax || p is DoStatementSyntax || p is ForEachStatementSyntax;
    SyntaxNode NodeAt(int pos) => M.FindToken(pos).Parent;
    // phép gán d chắc chắn chạy trước điểm dùng: đi lên tới tổ tiên chung không qua if/else/switch/vòng lặp/catch/lambda/&&,||
    bool Dominates(SyntaxNode d, SyntaxNode use)
    {
        var anc = new HashSet<SyntaxNode>(use.AncestorsAndSelf());
        for (var p = d.Parent; p != null && !anc.Contains(p); p = p.Parent)
            if (p is IfStatementSyntax || p is ElseClauseSyntax || p is SwitchSectionSyntax || IsLoop(p) || p is CatchClauseSyntax || p is ConditionalExpressionSyntax || p is LambdaExpressionSyntax || p is BinaryExpressionSyntax) return false;
        return true;
    }
    // ---- reaching definitions theo CFG (rd); rơi về phân tích cú pháp cũ khi không có CFG hoặc điểm dùng nằm ngoài CFG ----
    static bool IsIdentKey(string k) => Regex.IsMatch(k, @"^@?[A-Za-z_][A-Za-z0-9_]*$");
    public ISymbol SymAt(int pos)
    {
        if (rd == null) return null;
        try { var si = rd.Model.GetSymbolInfo(NodeAt(pos)); return si.Symbol ?? si.CandidateSymbols.FirstOrDefault(); } catch { return null; }
    }
    // true nếu điểm dùng có kết quả CFG
    public bool CfgUse(int pos, out UseInfo u) { u = null; return rd != null && rd.Uses.TryGetValue(pos, out u) && u.Reachable; }
    (string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node) MapDef(DefSite d, string key)
    {
        foreach (var a in assigns) if (a.node == d.Node && a.key == key) return d.Kind == "ref" ? (a.key, a.pos, a.val, a.line, "ref", a.node) : a;
        for (var p = d.Node.Parent; p != null && p != M; p = p.Parent)
            foreach (var a in assigns) if (a.node == p && a.key == key && a.kind == "foreach") return a;
        string kind = d.Kind switch { "compound" or "inc" => d.Node is AssignmentExpressionSyntax ae ? ae.OperatorToken.Text : d.Node is PostfixUnaryExpressionSyntax pu ? pu.OperatorToken.Text : d.Node is PrefixUnaryExpressionSyntax pr ? pr.OperatorToken.Text : d.Kind, "set" => d.Node is AssignmentExpressionSyntax ? "set" : "other", _ => d.Kind };
        ExpressionSyntax val = d.Node is AssignmentExpressionSyntax ae2 ? ae2.Right : d.Node is VariableDeclaratorSyntax vd ? vd.Initializer?.Value : null;
        return (key, d.Node.SpanStart, val, Line(d.Node), kind, d.Node);
    }
    void Closure(List<int> ids, HashSet<int> seen, List<(string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node)> res, string key)
    {
        foreach (var i in ids)
        {
            if (!seen.Add(i)) continue;
            var d = rd.Defs[i];
            res.Add(MapDef(d, key));
            if (d.TargetUsePos >= 0 && rd.Uses.TryGetValue(d.TargetUsePos, out var u2)) Closure(u2.Defs, seen, res, key);
        }
    }
    public List<(string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node)> ReachingAll(string key, int pos)
    {
        if (!M.Span.Contains(pos)) return new();
        if (rd != null && IsIdentKey(key))
        {
            if (CfgUse(pos, out var u) && u.Sym.Name == key.TrimStart('@'))
            {
                var res = new List<(string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node)>();
                Closure(u.Defs, new HashSet<int>(), res, key);
                return res.OrderBy(a => a.pos).ToList();
            }
            if (!rd.Uses.ContainsKey(pos))
            {
                var sym = SymAt(pos);
                if (sym != null && !(sym is ILocalSymbol) && !(sym is IParameterSymbol)) return new();
            }
        }
        return ReachingOld(key, pos);
    }
    // phân tích cũ (fallback): từ phép gán chắc chắn gần nhất trở đi, cộng phép gán phía sau trong cùng vòng lặp
    public List<(string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node)> ReachingOld(string key, int pos)
    {
        // biểu thức nằm ngoài thân hàm (khởi tạo field/thuộc tính): không có biến cục bộ
        if (!M.Span.Contains(pos)) return new();
        var use = NodeAt(pos);
        var useLoops = new HashSet<SyntaxNode>(use.Ancestors().Where(IsLoop));
        var before = assigns.Where(a => a.key == key && a.pos < pos).OrderBy(a => a.pos).ToList();
        int start = 0;
        for (int i = before.Count - 1; i >= 0; i--)
            if ((before[i].kind == "init" || before[i].kind == "set" || before[i].kind == "foreach") && Dominates(before[i].node, use)) { start = i; break; }
        var res = before.Skip(start).ToList();
        res.AddRange(assigns.Where(a => a.key == key && a.pos > pos && a.node.Ancestors().Any(useLoops.Contains)).OrderBy(a => a.pos));
        return res;
    }
    (ExpressionSyntax val, int line, int count)? Reaching(string key, int pos)
    {
        var all = ReachingAll(key, pos);
        if (all.Count == 0) return null;
        var near = all.Where(a => a.pos < pos).LastOrDefault();
        if (near.node == null) near = all[0];
        return (near.val, near.line, all.Count);
    }
    // giá trị duy nhất tới được (không +=, ++, nhánh hay vòng lặp khác)
    // print=true: chỉ thay khi các biến trong vế phải có cùng tập phép gán tại điểm dùng như tại phép gán (tránh in sai ngữ cảnh)
    ExpressionSyntax SingleDef(string key, int pos, bool print = false)
    {
        var all = ReachingAll(key, pos);
        if (all.Count != 1 || all[0].val == null || !(all[0].kind == "init" || all[0].kind == "set")) return null;
        if (IsForVar(key, pos)) return null;
        var use = NodeAt(pos);
        var dl = all[0].node.Ancestors().FirstOrDefault(IsLoop);
        if (dl != null && !use.Ancestors().Contains(dl)) return null;
        if (print && CfgUse(pos, out var u))
            foreach (var x in all[0].val.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>())
            {
                if (!rd.Uses.TryGetValue(x.SpanStart, out var xu)) continue;
                var atUse = rd.ReachingAt(xu.Sym, u.State).Select(d => d.Id).OrderBy(i => i).ToList();
                if (!atUse.SequenceEqual(xu.Defs.OrderBy(i => i))) return null;
            }
        return all[0].val;
    }
    bool IsForVar(string n, int pos)
    {
        if (CfgUse(pos, out var u))
            return u.Sym.DeclaringSyntaxReferences.FirstOrDefault()?.GetSyntax() is VariableDeclaratorSyntax vd && vd.Parent?.Parent is ForStatementSyntax;
        return forVars.Contains(n);
    }

    // ---------- hằng số ----------
    public long? Eval(ExpressionSyntax e, int depth = 0)
    {
        if (e == null || depth > 10) return null;
        switch (e)
        {
            case LiteralExpressionSyntax l when l.Token.Value is int or long or short or byte:
                return Convert.ToInt64(l.Token.Value);
            case ParenthesizedExpressionSyntax p: return Eval(p.Expression, depth + 1);
            case InvocationExpressionSyntax ci when ci.Expression is MemberAccessExpressionSyntax cim && cim.Name.Identifier.Text == "ColumnLettersToIndex" && ci.ArgumentList.Arguments.Count == 1:
                { var cs = StrOf(ci.ArgumentList.Arguments[0].Expression, depth + 1); return cs != null && Regex.IsMatch(cs, "^[A-Z]+$") ? ColIndex(cs) + 1 : null; }
            case CastExpressionSyntax ce: return Eval(ce.Expression, depth + 1);
            case PrefixUnaryExpressionSyntax pu when pu.IsKind(SyntaxKind.UnaryMinusExpression): return -Eval(pu.Operand, depth + 1);
            case BinaryExpressionSyntax b:
                var x = Eval(b.Left, depth + 1); var y = Eval(b.Right, depth + 1);
                if (x == null || y == null) return null;
                return b.Kind() switch
                {
                    SyntaxKind.AddExpression => x + y, SyntaxKind.SubtractExpression => x - y,
                    SyntaxKind.MultiplyExpression => x * y, _ => null,
                };
            case IdentifierNameSyntax id:
                {
                    string n = id.Identifier.Text;
                    if (assigns.Any(a => a.key == n)) { var sd = SingleDef(n, e.SpanStart); return sd == null ? null : Eval(sd, depth + 1); }
                    if (paramNames.Contains(n)) return null;
                    if (C.Fields.TryGetValue(n, out var fe)) return Eval(fe, depth + 1);
                    return null;
                }
            case ElementAccessExpressionSyntax ea when ea.ArgumentList.Arguments.Count == 1:
                {
                    var idx = Eval(ea.ArgumentList.Arguments[0].Expression, depth + 1);
                    var arr = ArrayOf(ea.Expression, depth + 1);
                    if (idx != null && arr != null && idx >= 0 && idx < arr.Length) return arr[idx.Value];
                    var sd2 = SingleDef(Norm(ea.ToString()), e.SpanStart);
                    return sd2 == null ? null : Eval(sd2, depth + 1);
                }
        }
        return null;
    }
    bool InLoopAssigned(string name)
    {
        // biến được gán trong thân vòng lặp hoặc là biến đếm vòng lặp: không gập hằng
        foreach (var a in assigns.Where(a => a.key == name))
        {
            var node = M.FindNode(new Microsoft.CodeAnalysis.Text.TextSpan(a.pos, 1));
            for (var p = node; p != null && p != M; p = p.Parent)
                if (p is ForStatementSyntax || p is WhileStatementSyntax || p is DoStatementSyntax || p is ForEachStatementSyntax) return true;
        }
        return false;
    }
    public long[] ArrayOf(ExpressionSyntax e, int depth = 0)
    {
        if (e == null || depth > 10) return null;
        switch (e)
        {
            case InvocationExpressionSyntax inv when inv.Expression.ToString().EndsWith("ArrayCellId") && inv.ArgumentList.Arguments.Count == 1:
                {
                    var s = StrOf(inv.ArgumentList.Arguments[0].Expression, depth + 1);
                    var m = s == null ? null : Regex.Match(s, @"^([A-Z]+)_(\d+)$");
                    if (m == null || !m.Success) return null;
                    return new long[] { ColIndex(m.Groups[1].Value), long.Parse(m.Groups[2].Value) - 1 };
                }
            case ArrayCreationExpressionSyntax ac when ac.Initializer != null:
                {
                    var vals = ac.Initializer.Expressions.Select(x => Eval(x, depth + 1)).ToList();
                    return vals.Any(v => v == null) ? null : vals.Select(v => v.Value).ToArray();
                }
            case ImplicitArrayCreationExpressionSyntax ic:
                {
                    var vals = ic.Initializer.Expressions.Select(x => Eval(x, depth + 1)).ToList();
                    return vals.Any(v => v == null) ? null : vals.Select(v => v.Value).ToArray();
                }
            case InitializerExpressionSyntax ie:
                {
                    var vals = ie.Expressions.Select(x => Eval(x, depth + 1)).ToList();
                    return vals.Any(v => v == null) ? null : vals.Select(v => v.Value).ToArray();
                }
            case IdentifierNameSyntax id:
                {
                    if (assigns.Any(a => a.key == id.Identifier.Text)) { var sd = SingleDef(id.Identifier.Text, e.SpanStart); return sd == null ? null : ArrayOf(sd, depth + 1); }
                    if (C.Fields.TryGetValue(id.Identifier.Text, out var fe)) return ArrayOf(fe, depth + 1);
                    return null;
                }
        }
        return null;
    }
    public string StrOf(ExpressionSyntax e, int depth = 0)
    {
        if (e == null || depth > 10) return null;
        if (e is LiteralExpressionSyntax l && l.IsKind(SyntaxKind.StringLiteralExpression)) return l.Token.ValueText;
        if (e is ParenthesizedExpressionSyntax p) return StrOf(p.Expression, depth + 1);
        if (e is IdentifierNameSyntax id)
        {
            var r = Reaching(id.Identifier.Text, e.SpanStart);
            if (r != null) return r.Value.val == null || r.Value.count > 1 ? null : StrOf(r.Value.val, depth + 1);
            if (C.Fields.TryGetValue(id.Identifier.Text, out var fe)) return StrOf(fe, depth + 1);
        }
        return null;
    }
    static long ColIndex(string letters) { long n = 0; foreach (char ch in letters) n = n * 26 + (ch - 'A' + 1); return n - 1; }
    public static string ColLetters(long c) { var sb = new StringBuilder(); c += 1; while (c > 0) { long r = (c - 1) % 26; sb.Insert(0, (char)('A' + r)); c = (c - 1) / 26; } return sb.ToString(); }

    public int? SheetIndex()
    {
        var m = Regex.Match(M.Identifier.Text, @"_(\d+)$");
        if (M.Identifier.Text.StartsWith("Validate") && m.Success) return int.Parse(m.Groups[1].Value);
        foreach (var inv in M.DescendantNodes().OfType<InvocationExpressionSyntax>())
            if (inv.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "TryGetValue" && ma.Expression.ToString() == "levelWarning")
            { var v = Eval(inv.ArgumentList.Arguments[0].Expression); if (v != null) return (int)v; }
        foreach (var ea in M.DescendantNodes().OfType<ElementAccessExpressionSyntax>())
            if (ea.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "Sheets" && ea.ArgumentList.Arguments.Count == 1)
            { var v = Eval(ea.ArgumentList.Arguments[0].Expression); if (v != null) return (int)v; }
        if (m.Success && Regex.IsMatch(M.Identifier.Text, @"^(UpdateFormula|ButtonClick|CellChange|AutoBindData|Refresh|CellEditchange|ClearData)_\d+$")) return int.Parse(m.Groups[1].Value);
        return null;
    }

    // sheet của biểu thức SheetView
    string SheetOf(ExpressionSyntax e, int depth = 0)
    {
        if (e == null || depth > 6) return null;
        if (e is ElementAccessExpressionSyntax ea && ea.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "Sheets" && ea.ArgumentList.Arguments.Count == 1)
        {
            var a = ea.ArgumentList.Arguments[0].Expression;
            var v = Eval(a);
            if (v != null) return "S" + v;
            if (Regex.IsMatch(Norm(a.ToString()), @"\.Sheets\.Count-1$")) return "S[last]";
            var s = StrOf(a); if (s != null) return s;
            return "S[" + R(a) + "]";
        }
        if (e is IdentifierNameSyntax id)
        {
            if (paramNames.Contains(id.Identifier.Text)) return "param:" + id.Identifier.Text;
            var r = Reaching(id.Identifier.Text, e.SpanStart);
            if (r != null && r.Value.val != null) return SheetOf(r.Value.val, depth + 1);
        }
        if (e is MemberAccessExpressionSyntax m2 && m2.Name.Identifier.Text == "ActiveSheet") return "active";
        return null;
    }

    // ---------- in trung tính ----------
    readonly Dictionary<string, string> bindings = new(StringComparer.Ordinal);
    int inlineDepth = 0;
    public string R(ExpressionSyntax e)
    {
        string r0 = R0(e);
        if (r0 != null && r0.StartsWith(OLD) && !(e.Parent is MemberAccessExpressionSyntax pm0 && pm0.Expression == e) && !(e.Parent is ElementAccessExpressionSyntax pe0 && pe0.Expression == e) && !(e.Parent is ConditionalAccessExpressionSyntax))
            r0 = r0 == OLD ? "model.old()" : "model.old(" + r0.Substring(OLD.Length).TrimStart('.') + ")";
        return r0;
    }
    string R0(ExpressionSyntax e)
    {
        if (e == null) return "?";
        switch (e)
        {
            case LiteralExpressionSyntax l:
                if (l.IsKind(SyntaxKind.StringLiteralExpression)) return JsonSerializer.Serialize(l.Token.ValueText, new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
                if (l.IsKind(SyntaxKind.NullLiteralExpression)) return "null";
                if (l.IsKind(SyntaxKind.TrueLiteralExpression)) return "true";
                if (l.IsKind(SyntaxKind.FalseLiteralExpression)) return "false";
                return Convert.ToString(l.Token.Value, System.Globalization.CultureInfo.InvariantCulture);
            case ParenthesizedExpressionSyntax p: return "(" + R(p.Expression) + ")";
            case CastExpressionSyntax ce: return R(ce.Expression);
            case PrefixUnaryExpressionSyntax pu:
                return pu.Kind() switch
                {
                    SyntaxKind.LogicalNotExpression => "NOT " + Wrap(pu.Operand),
                    SyntaxKind.UnaryMinusExpression => "-" + Wrap(pu.Operand),
                    _ => pu.OperatorToken.Text + R(pu.Operand),
                };
            case PostfixUnaryExpressionSyntax po: return R(po.Operand) + po.OperatorToken.Text;
            case BinaryExpressionSyntax b:
                {
                    string op = b.Kind() switch
                    {
                        SyntaxKind.LogicalAndExpression or SyntaxKind.BitwiseAndExpression => "AND",
                        SyntaxKind.LogicalOrExpression or SyntaxKind.BitwiseOrExpression => "OR",
                        SyntaxKind.ExclusiveOrExpression => "XOR",
                        SyntaxKind.EqualsExpression => "=", SyntaxKind.NotEqualsExpression => "<>",
                        SyntaxKind.CoalesceExpression => "DEFAULT",
                        SyntaxKind.IsExpression => "IS", SyntaxKind.AsExpression => "AS",
                        _ => b.OperatorToken.Text,
                    };
                    return R(b.Left) + " " + op + " " + R(b.Right);
                }
            case ConditionalExpressionSyntax c: return "IF(" + R(c.Condition) + ", " + R(c.WhenTrue) + ", " + R(c.WhenFalse) + ")";
            case IsPatternExpressionSyntax ip: return R(ip.Expression) + " IS " + ip.Pattern.ToString();
            case ConditionalAccessExpressionSyntax ca: return RChain(ca.WhenNotNull, R(ca.Expression));
            case IdentifierNameSyntax id: return RIdent(id);
            case ThisExpressionSyntax: return "this";
            case ElementAccessExpressionSyntax ea: return RElement(ea);
            case MemberAccessExpressionSyntax ma: return RMember(ma);
            case InvocationExpressionSyntax inv: return RInvoke(inv);
            case ObjectCreationExpressionSyntax od when od.Type.ToString() == "DateTime" && od.ArgumentList != null && od.ArgumentList.Arguments.Count >= 3 && od.ArgumentList.Arguments.Take(3).All(x => Eval(x.Expression) != null):
                {
                    var dv = od.ArgumentList.Arguments.Select(x => Eval(x.Expression)).ToList();
                    return "date(" + dv[0].Value.ToString("D4") + "-" + dv[1].Value.ToString("D2") + "-" + dv[2].Value.ToString("D2") + (dv.Count >= 6 && dv.All(x => x != null) ? " " + dv[3].Value.ToString("D2") + ":" + dv[4].Value.ToString("D2") + ":" + dv[5].Value.ToString("D2") : "") + ")";
                }
            case ObjectCreationExpressionSyntax odt when odt.Type.ToString() == "DateTime" && odt.ArgumentList != null:
                return "date(" + string.Join(", ", odt.ArgumentList.Arguments.Select(a => R(a.Expression))) + ")";
            case ObjectCreationExpressionSyntax oc: return "create(" + string.Join(", ", new[] { oc.Type.ToString() }.Concat((oc.ArgumentList?.Arguments ?? default).Select(a => R(a.Expression)))) + ")";
            case ArrayCreationExpressionSyntax ac: return ac.Initializer == null ? "new_array(" + string.Join(", ", ac.Type.RankSpecifiers.SelectMany(r => r.Sizes).Select(R)) + ")" : "[" + string.Join(", ", ac.Initializer.Expressions.Select(R)) + "]";
            case ImplicitArrayCreationExpressionSyntax ic: return "[" + string.Join(", ", ic.Initializer.Expressions.Select(R)) + "]";
            case InterpolatedStringExpressionSyntax isx: return "format(" + JsonSerializer.Serialize(isx.ToString()) + ")";
            case PredefinedTypeSyntax pt: return pt.Keyword.Text;
            case TypeOfExpressionSyntax to: return "type(" + to.Type + ")";
            case AssignmentExpressionSyntax asg: return R(asg.Left) + " := " + R(asg.Right);
            case SimpleLambdaExpressionSyntax or ParenthesizedLambdaExpressionSyntax: return RLambda((LambdaExpressionSyntax)e);
            case DeclarationExpressionSyntax de: return de.Designation.ToString();
            case GenericNameSyntax g: return g.Identifier.Text;
        }
        return "expr(" + e.Kind().ToString() + ")";
    }
    string RLambda(LambdaExpressionSyntax l)
    {
        var ps = l is SimpleLambdaExpressionSyntax sl ? new[] { sl.Parameter.Identifier.Text } : ((ParenthesizedLambdaExpressionSyntax)l).ParameterList.Parameters.Select(p => p.Identifier.Text).ToArray();
        if (l.ExpressionBody == null) return "(" + string.Join(", ", ps) + " => {khối lệnh})";
        return "(" + string.Join(", ", ps) + " => " + R(l.ExpressionBody) + ")";
    }
    // biểu thức lọc dạng x => x.Thuộc_tính == E: trả (thuộc tính, E)
    static (string prop, ExpressionSyntax val)? SimpleFilter(ExpressionSyntax arg)
    {
        if (arg is not LambdaExpressionSyntax l || l.ExpressionBody == null) return null;
        string pn = l is SimpleLambdaExpressionSyntax sl ? sl.Parameter.Identifier.Text : ((ParenthesizedLambdaExpressionSyntax)l).ParameterList.Parameters.Count == 1 ? ((ParenthesizedLambdaExpressionSyntax)l).ParameterList.Parameters[0].Identifier.Text : null;
        if (pn == null) return null;
        ExpressionSyntax left = null, right = null;
        if (l.ExpressionBody is BinaryExpressionSyntax b && b.IsKind(SyntaxKind.EqualsExpression)) { left = b.Left; right = b.Right; }
        else if (l.ExpressionBody is InvocationExpressionSyntax iv && iv.Expression is MemberAccessExpressionSyntax im && im.Name.Identifier.Text == "Equals" && iv.ArgumentList.Arguments.Count == 1) { left = im.Expression; right = iv.ArgumentList.Arguments[0].Expression; }
        if (left is MemberAccessExpressionSyntax lm && lm.Expression is IdentifierNameSyntax li && li.Identifier.Text == pn && !right.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>().Any(x => x.Identifier.Text == pn))
            return (lm.Name.Identifier.Text, right);
        return null;
    }
    string RFind(string name, string recv, ExpressionSyntax lambdaArg)
    {
        var f = SimpleFilter(lambdaArg);
        if (f == null) return null;
        string v = R(f.Value.val);
        if (f.Value.prop == "Path" && (name is "Find" or "FirstOrDefault" or "First" or "SingleOrDefault")) return "cell_by_path(" + recv + ", " + v + ")";
        string op = name switch { "Find" or "FirstOrDefault" or "First" or "SingleOrDefault" => "find_by", "Where" or "FindAll" => "filter_by", "Exists" or "Any" => "exists_by", "Count" => "count_by", _ => null };
        return op == null ? null : op + "(" + recv + ", " + f.Value.prop + " = " + v + ")";
    }
    // chuỗi sau toán tử ?. : prefix là phần nhận đã in
    string RChain(ExpressionSyntax e, string prefix)
    {
        switch (e)
        {
            case MemberBindingExpressionSyntax mb: return prefix + "." + mb.Name.Identifier.Text;
            case ElementBindingExpressionSyntax eb: return prefix + "[" + string.Join(", ", eb.ArgumentList.Arguments.Select(x => R(x.Expression))) + "]";
            case MemberAccessExpressionSyntax ma: return RChain(ma.Expression, prefix) + "." + ma.Name.Identifier.Text;
            case ElementAccessExpressionSyntax ea: return RChain(ea.Expression, prefix) + "[" + string.Join(", ", ea.ArgumentList.Arguments.Select(x => R(x.Expression))) + "]";
            case ConditionalAccessExpressionSyntax c2: return RChain(c2.WhenNotNull, RChain(c2.Expression, prefix));
            case InvocationExpressionSyntax inv:
                {
                    var args = inv.ArgumentList.Arguments.Select(x => x.Expression).ToList();
                    string name; string recv;
                    if (inv.Expression is MemberAccessExpressionSyntax im) { name = im.Name.Identifier.Text; recv = RChain(im.Expression, prefix); }
                    else if (inv.Expression is MemberBindingExpressionSyntax ib) { name = ib.Name.Identifier.Text; recv = prefix; }
                    else return R(e);
                    if (args.Count == 1) { var fr = RFind(name, recv, args[0]); if (fr != null) return fr; }
                    string fn = FnMap.TryGetValue(name, out var mp) ? mp : name;
                    return fn + "(" + string.Join(", ", new[] { recv }.Concat(args.Select(R))) + ")";
                }
        }
        return R(e);
    }
    string Wrap(ExpressionSyntax e) { var s = R(e); return e is BinaryExpressionSyntax || e is ConditionalExpressionSyntax ? "(" + s + ")" : s; }

    string RIdent(IdentifierNameSyntax id)
    {
        string n = id.Identifier.Text;
        var v = Eval(id); if (v != null) return v.Value.ToString();
        if (id.Parent is ArgumentSyntax oa && !oa.RefKindKeyword.IsKind(SyntaxKind.None)) return n;
        if (id.Parent is DeclarationExpressionSyntax) return n;
        string ty = TypeNameOf(id);
        if (ty == "SheetView") return "sheet(" + SheetName(id) + ")";
        if (ty == "FpSpread") return "workbook";
        if (ty == "HTKKCMCToKhaiInfo" && n != "ToKhaiInfo") return OLD;
        if (paramNames.Contains(n) || IsForVar(n, id.SpanStart)) return n;
        var all = ReachingAll(n, id.SpanStart);
        if (all.Count == 0)
        {
            if (C.Fields.ContainsKey(n) && StrOf(id) is string fs) return JsonSerializer.Serialize(fs);
            return n;
        }
        string pk = KeyFor(n, id.SpanStart);
        if (inProgress.Contains(pk)) return pk;
        var sd = SingleDef(n, id.SpanStart, true);
        if (sd != null && !(sd is ArrayCreationExpressionSyntax ac0 && ac0.Initializer == null) && !(sd is ObjectCreationExpressionSyntax) && inlineDepth < 2)
        {
            inProgress.Add(pk);
            inlineDepth++; string s1 = R(sd); inlineDepth--;
            inProgress.Remove(pk);
            if (s1.Length <= 140) return s1.Contains(' ') ? "(" + s1 + ")" : s1;
            Bind(pk, s1, all[0].line, id);
            return pk;
        }
        BindAll(pk, all, id);
        return pk;
    }
    const string OLD = "\u0001";
    // tên kiểu của định danh: theo SemanticModel (kể cả kiểu lỗi), nếu không có thì theo khai báo cú pháp
    string TypeNameOf(IdentifierNameSyntax id)
    {
        string n = id.Identifier.Text;
        var sym = SymAt(id.SpanStart);
        ITypeSymbol t = sym switch { ILocalSymbol l => l.Type, IParameterSymbol p => p.Type, IFieldSymbol f => f.Type, IPropertySymbol pp => pp.Type, _ => null };
        if (t != null && !(t is IErrorTypeSymbol && t.Name == "var") && t.Name != "") return t.Name;
        if (sym != null && !(sym is ILocalSymbol || sym is IParameterSymbol || sym is IFieldSymbol || sym is IPropertySymbol)) return null;
        var pd = M.ParameterList.Parameters.FirstOrDefault(x => x.Identifier.Text == n);
        if (pd != null) return pd.Type?.ToString().Split('.').Last();
        foreach (var vd in M.DescendantNodes().OfType<VariableDeclaratorSyntax>())
            if (vd.Identifier.Text == n && vd.Parent is VariableDeclarationSyntax decl)
            {
                if (!decl.Type.IsVar) return decl.Type.ToString().Split('.').Last();
                if (vd.Initializer?.Value is ElementAccessExpressionSyntax ie && ie.Expression is MemberAccessExpressionSyntax im && im.Name.Identifier.Text == "Sheets") return "SheetView";
                if (vd.Initializer?.Value is MemberAccessExpressionSyntax am && am.Name.Identifier.Text == "ActiveSheet") return "SheetView";
            }
        if (C.FieldTypes.TryGetValue(n, out var ft)) return ft.Split('.').Last();
        return null;
    }
    // tên sheet của biến SheetView: S<n> | S[biểu thức] | S[tham số] (giống tiền tố trong cells_marked)
    string SheetName(ExpressionSyntax e)
    {
        string sh = SheetOf(e);
        if (sh != null && !sh.StartsWith("param:") && sh != "active") return sh;
        sheetIdxCache ??= SheetIndex() ?? -1;
        if (sheetIdxCache >= 0) return "S" + sheetIdxCache;
        if (sh != null && sh.StartsWith("param:")) return "S[" + sh.Substring(6) + "]";
        if (sh == "active") return "S[active]";
        return e is IdentifierNameSyntax ii ? "S[" + ii.Identifier.Text + "]" : "Sx";
    }
    // khóa binding theo ISymbol: hai biến cùng tên khác phạm vi trong một check có khóa khác nhau (tên@L<dòng khai báo>)
    readonly Dictionary<ISymbol, string> symKey = new(SymbolEqualityComparer.Default);
    readonly Dictionary<string, ISymbol> keySym = new(StringComparer.Ordinal);
    readonly HashSet<string> inProgress = new(StringComparer.Ordinal);
    readonly SortedDictionary<string, object> bindInfo = new(StringComparer.Ordinal);
    static int DeclLine(ISymbol s) { var l = s.Locations.FirstOrDefault(); return l == null || !l.IsInSource ? 0 : l.GetLineSpan().StartLinePosition.Line + 1; }
    string KeyFor(string n, int pos)
    {
        if (!CfgUse(pos, out var u)) return n;
        if (symKey.TryGetValue(u.Sym, out var k)) return k;
        k = keySym.ContainsKey(n) ? n + "@L" + DeclLine(u.Sym) : n;
        symKey[u.Sym] = k; keySym[k] = u.Sym;
        return k;
    }
    void Info(string key, SyntaxNode use, IEnumerable<int> defLines)
    {
        bool cfg = CfgUse(use.SpanStart, out var u);
        var d = new SortedDictionary<string, object>(StringComparer.Ordinal)
        {
            ["method"] = cfg ? "cfg" : "fallback",
            ["use_line"] = Line(use),
            ["def_lines"] = defLines.Distinct().OrderBy(x => x).ToList(),
        };
        if (cfg) d["decl_line"] = DeclLine(u.Sym);
        bindInfo[key] = d;
    }
    // liệt kê mọi phép gán có thể tới điểm dùng, kèm điều kiện riêng của từng phép gán
    void BindAll(string n, List<(string key, int pos, ExpressionSyntax val, int line, string kind, SyntaxNode node)> all, SyntaxNode use)
    {
        if (bindings.ContainsKey(n) || inProgress.Contains(n)) return;
        inProgress.Add(n);
        inlineDepth++;
        var saved = LastCondLines;
        var parts = new List<string>();
        var useAnc = new HashSet<SyntaxNode>(use.AncestorsAndSelf());
        foreach (var a in all)
        {
            var stop = a.node.Ancestors().FirstOrDefault(useAnc.Contains);
            var g = ConditionsUpTo(a.node, stop);
            string val = a.kind switch
            {
                "init" or "set" or "other" => a.val == null ? "?" : R(a.val),
                "out" or "ref" => "<gán qua tham số out/ref>",
                "foreach" => "<phần tử vòng lặp>",
                "pattern" => "<gán khi khớp mẫu kiểu>",
                "catch" => "<ngoại lệ bắt được>",
                "decon" => "<gán qua phân rã bộ giá trị>",
                "param" => "<giá trị tham số khi gọi hàm>",
                "++" or "--" => a.kind,
                _ => a.kind + " " + (a.val == null ? "?" : R(a.val)),
            };
            string gs = g.Count == 0 ? "" : " [khi " + string.Join(" AND ", g.Select(WrapOr)) + "]";
            parts.Add("L" + a.line + gs + ": " + val);
        }
        LastCondLines = saved;
        inlineDepth--;
        string sv = all.Count == 1 ? parts[0] : "<" + all.Count + " phép gán có thể tới> " + string.Join(" | ", parts);
        bindings[n] = sv;
        inProgress.Remove(n);
        Info(n, use, all.Select(a => a.line));
    }
    void Bind(string n, string s, int line, SyntaxNode use) { bindings[n] = s + " @L" + line; Info(n, use, new[] { line }); }

    string RElement(ElementAccessExpressionSyntax ea)
    {
        var args = ea.ArgumentList.Arguments;
        if (ea.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "Cells" && args.Count == 2)
            return "cell(" + CellRef(ma.Expression, args[0].Expression, args[1].Expression) + ")";
        if (ea.Expression is MemberAccessExpressionSyntax ms && ms.Name.Identifier.Text == "Sheets" && args.Count == 1)
            return "sheet(" + (SheetOf(ea) ?? R(args[0].Expression)) + ")";
        var v = Eval(ea); if (v != null) return v.Value.ToString();
        var sdE = SingleDef(Norm(ea.ToString()), ea.SpanStart, true);
        if (sdE != null && inlineDepth < 2)
        {
            inlineDepth++; string s = R(sdE); inlineDepth--;
            if (s.Length <= 140) return s.Contains(' ') ? "(" + s + ")" : s;
        }
        return R(ea.Expression) + "[" + string.Join(", ", args.Select(a => R(a.Expression))) + "]";
    }
    int? sheetIdxCache;
    public string LastSheet;
    public string CellRef(ExpressionSyntax sheetExpr, ExpressionSyntax row, ExpressionSyntax col)
    {
        string sh = SheetName(sheetExpr); // S<n>; S[tham số]/S[active]/Sx khi không xác định tĩnh
        string prefix = sh + "!";
        var r = Eval(row); var c = Eval(col);
        string cs = c != null ? ColLetters(c.Value) : "col{" + R(col) + "}";
        string rs = r != null ? (r.Value + 1).ToString() : "{" + R(row) + " +1}";
        LastSheet = sh; // gán sau khi in hàng/cột (biểu thức hàng/cột có thể chứa ô của sheet khác)
        return prefix + cs + rs;
    }
    string RMember(MemberAccessExpressionSyntax ma)
    {
        string name = ma.Name.Identifier.Text;
        if (ma.Expression is ElementAccessExpressionSyntax ea && ea.Expression is MemberAccessExpressionSyntax cm && cm.Name.Identifier.Text == "Cells" && ea.ArgumentList.Arguments.Count == 2)
        {
            string cref = CellRef(cm.Expression, ea.ArgumentList.Arguments[0].Expression, ea.ArgumentList.Arguments[1].Expression);
            return name switch { "Text" => "text(" + cref + ")", "Value" => "value(" + cref + ")", _ => name.ToLowerInvariant() + "(" + cref + ")" };
        }
        if ((name == "Length" || name == "Count") && ma.Expression is MemberAccessExpressionSyntax sc && sc.Name.Identifier.Text == "Sheets") return "sheet_count()";
        if (name == "Length" || name == "Count") return "length(" + R(ma.Expression) + ")";
        if (name == "Visible") return "visible(" + R(ma.Expression) + ")";
        if (name == "ActiveSheet") return "sheet(S[active])";
        if (name == "Sheets") return "sheets()";
        if (ma.Expression is PredefinedTypeSyntax || ma.Expression is IdentifierNameSyntax q && char.IsUpper(q.Identifier.Text[0]) && !paramNames.Contains(q.Identifier.Text) && Reaching(q.Identifier.Text, ma.SpanStart) == null && !C.Fields.ContainsKey(q.Identifier.Text) && !(TypeNameOf(q) is "SheetView" or "FpSpread" or "HTKKCMCToKhaiInfo"))
            return StaticMember(ma.Expression.ToString(), name); // hằng/thuộc tính tĩnh -> ký hiệu trung tính
        var v = Eval(ma); if (v != null) return v.Value.ToString();
        return R(ma.Expression) + "." + name;
    }
    static string Snake(string n) => Regex.Replace(n, "(?<=[a-z0-9])([A-Z])", "_$1").ToLowerInvariant();
    // hằng/thuộc tính tĩnh: DateTime.Now -> now(), Color.Pink -> color(pink), GeneralFunctions.KEY_PRESS -> key_press, string.Empty -> ""
    static string StaticMember(string type, string name)
    {
        string t = type.Split('.').Last();
        switch (t)
        {
            case "DateTime":
                return name switch { "Now" => "now()", "Today" => "today()", "MinValue" => "date_min()", "MaxValue" => "date_max()", _ => "date_" + Snake(name) + "()" };
            case "Color": case "SystemColors": return "color(" + name.ToLowerInvariant() + ")";
            case "string": case "String": if (name == "Empty") return "\"\""; break;
            case "int": case "Int32": case "long": case "Int64": case "decimal": case "Decimal": case "double": case "Double":
                return name switch { "MaxValue" => "max_value(" + t.ToLowerInvariant() + ")", "MinValue" => "min_value(" + t.ToLowerInvariant() + ")", "Zero" => "0", "One" => "1", _ => Snake(name) };
            case "Environment": if (name == "NewLine") return "newline()"; break;
            case "Keys": return "key(" + name.ToLowerInvariant() + ")";
        }
        if (t == "ToKhaiInfo" || t == "PeriodInfo" || t == "ToKhaiInfoOld") return type + "." + name; // dữ liệu tờ khai (model)
        if (char.IsUpper(t[0]) && Regex.IsMatch(t, "^[A-Z][A-Za-z0-9]*$") && (t.EndsWith("Functions") || t.EndsWith("File") || t.EndsWith("Base") || t == "Common" || t.EndsWith("Helpers") || t.EndsWith("Constants") || t.EndsWith("Const") || t == "WinSettings" || t == "HTKKConstants"))
            return Snake(name);
        return type + "." + name;
    }
    static readonly Dictionary<string, string> FnMap = new(StringComparer.Ordinal)
    {
        ["ColumnLettersToIndex"] = "col_index",
        ["isNullString"] = "is_empty", ["IsNullOrEmpty"] = "is_empty", ["IsNullOrWhiteSpace"] = "is_blank",
        ["ToDouble"] = "to_number", ["ToDecimal"] = "to_number", ["ToInt32"] = "to_int", ["ToInt64"] = "to_int", ["ToString"] = "to_text",
        ["Abs"] = "abs", ["Round"] = "round", ["Max"] = "max", ["Min"] = "min", ["Trim"] = "trim", ["ToUpper"] = "upper", ["ToLower"] = "lower",
        ["Contains"] = "contains", ["StartsWith"] = "starts_with", ["EndsWith"] = "ends_with", ["Substring"] = "substring", ["Replace"] = "replace",
        ["IndexOf"] = "index_of", ["Split"] = "split", ["Parse"] = "parse", ["TryParse"] = "try_parse", ["Compare"] = "compare",
        ["GetText"] = "text", ["GetValue"] = "value", ["Format"] = "format", ["Concat"] = "concat",
    };
    string RInvoke(InvocationExpressionSyntax inv)
    {
        var args = inv.ArgumentList.Arguments.Select(a => a.Expression).ToList();
        string name; ExpressionSyntax recv = null;
        if (inv.Expression is MemberAccessExpressionSyntax ma) { name = ma.Name.Identifier.Text; recv = ma.Expression; }
        else if (inv.Expression is IdentifierNameSyntax id) name = id.Identifier.Text;
        else if (inv.Expression is GenericNameSyntax g) name = g.Identifier.Text;
        else name = "call";
        if (recv != null && recv.ToString() == "HTKKFile" && name == "GetMessage") return "message(" + string.Join(", ", args.Select(R)) + ")";
        if (recv != null && recv.ToString().EndsWith("GeneralFunctions") && name == "ArrayCellId" && args.Count == 1)
        { var s = StrOf(args[0]); if (s != null) return "cellid(" + s + ")"; }
        if ((name == "GetText" || name == "GetValue") && args.Count == 2 && recv != null)
            return (name == "GetText" ? "text(" : "value(") + CellRef(recv, args[0], args[1]) + ")";
        if (name == "Equals" && recv != null && args.Count >= 1 && !(recv is PredefinedTypeSyntax)) return Wrap(recv) + " = " + Wrap(args[0]);
        bool staticRecv = recv == null || recv is PredefinedTypeSyntax || recv is IdentifierNameSyntax q && char.IsUpper(q.Identifier.Text[0]) && !paramNames.Contains(q.Identifier.Text) && Reaching(q.Identifier.Text, inv.SpanStart) == null && !C.Fields.ContainsKey(q.Identifier.Text);
        if (args.Count == 1 && recv != null && args[0] is LambdaExpressionSyntax)
        { var fr = RFind(name, R(recv), args[0]); if (fr != null) return fr; }
        string fn = FnMap.TryGetValue(name, out var mapped) ? mapped : name;
        var parts = new List<string>();
        if (!staticRecv) parts.Add(R(recv));
        parts.AddRange(args.Select(R));
        if (name == "Trim" && parts.Count == 1) return "trim(" + parts[0] + ")";
        return fn + "(" + string.Join(", ", parts) + ")";
    }

    // vế có OR ở mức ngoài cùng (ngoài ngoặc và chuỗi) thì bọc ngoặc trước khi nối bằng AND
    static string WrapOr(string x)
    {
        int depth = 0; bool q = false;
        for (int i = 0; i < x.Length; i++)
        {
            char ch = x[i];
            if (q) { if (ch == '\\') i++; else if (ch == '"') q = false; continue; }
            if (ch == '"') q = true;
            else if (ch == '(' || ch == '[' || ch == '{') depth++;
            else if (ch == ')' || ch == ']' || ch == '}') depth--;
            else if (depth == 0 && string.CompareOrdinal(x, i, " OR ", 0, 4) == 0) return "(" + x + ")";
        }
        return x;
    }
    // ---------- điều kiện ----------
    string Neg(ExpressionSyntax c)
    {
        while (c is ParenthesizedExpressionSyntax pp) c = pp.Expression;
        if (c is PrefixUnaryExpressionSyntax pu && pu.IsKind(SyntaxKind.LogicalNotExpression)) return R(pu.Operand);
        return "NOT (" + R(c) + ")";
    }
    public List<int> LastCondLines = new();
    List<string> Conditions(SyntaxNode n) => ConditionsUpTo(n, null);
    List<string> ConditionsUpTo(SyntaxNode n, SyntaxNode stop)
    {
        var conds = new List<string>();
        var lines = new List<int>();
        SyntaxNode child = n;
        for (var p = n.Parent; p != null && p != M && p != stop; child = p, p = p.Parent)
        {
            switch (p)
            {
                case IfStatementSyntax ifs when child == ifs.Statement: conds.Add(R(ifs.Condition)); break;
                case IfStatementSyntax ifs2 when child == ifs2.Else: conds.Add(Neg(ifs2.Condition)); break;
                case ElseClauseSyntax: break;
                case ConditionalExpressionSyntax ce when child == ce.WhenTrue: conds.Add(R(ce.Condition)); break;
                case ConditionalExpressionSyntax ce2 when child == ce2.WhenFalse: conds.Add(Neg(ce2.Condition)); break;
                case ForStatementSyntax fs:
                    if (child != fs.Condition && !fs.Initializers.Contains(child))
                    {
                        string v = fs.Declaration?.Variables.FirstOrDefault()?.Identifier.Text;
                        string init = fs.Declaration?.Variables.FirstOrDefault()?.Initializer?.Value is ExpressionSyntax iv ? R(iv) : null;
                        if (v == null && fs.Initializers.FirstOrDefault() is AssignmentExpressionSyntax fa) { v = R(fa.Left); init = R(fa.Right); }
                        v ??= "?"; init ??= "?";
                        conds.Add("FOR " + v + " FROM " + init + " WHILE " + (fs.Condition != null ? R(fs.Condition) : "true"));
                    }
                    break;
                case ForEachStatementSyntax fe when child == fe.Statement: conds.Add("FOR EACH " + fe.Identifier.Text + " IN " + R(fe.Expression)); break;
                case WhileStatementSyntax ws when child == ws.Statement: conds.Add("WHILE " + R(ws.Condition)); break;
                case DoStatementSyntax ds when child == ds.Statement: conds.Add("REPEAT UNTIL NOT (" + R(ds.Condition) + ")"); break;
                case SwitchSectionSyntax ss:
                    {
                        var sw = ss.Parent as SwitchStatementSyntax;
                        var labels = ss.Labels.Select(l => l is CaseSwitchLabelSyntax cl ? R(cl.Value) : l is DefaultSwitchLabelSyntax ? "default" : l.ToString()).ToList();
                        conds.Add("CASE " + (sw != null ? R(sw.Expression) : "?") + " IN {" + string.Join(", ", labels) + "}");
                        break;
                    }
                case BinaryExpressionSyntax be when be.IsKind(SyntaxKind.LogicalAndExpression) && child == be.Right: conds.Add(R(be.Left)); break;
                case BinaryExpressionSyntax bo when bo.IsKind(SyntaxKind.LogicalOrExpression) && child == bo.Right: conds.Add(Neg(bo.Left)); break;
                case CatchClauseSyntax: conds.Add("ON EXCEPTION"); break;
                case BlockSyntax blk:
                    {
                        // điều kiện chặn sớm: if (X) { ...; return|continue|break; } đứng trước trong cùng block
                        var gc = new List<(string, int)>();
                        foreach (var st in blk.Statements)
                        {
                            if (st.SpanStart >= child.SpanStart) break;
                            if (st is IfStatementSyntax g && g.Else == null && EndsWithJump(g.Statement)) gc.Add((Neg(g.Condition), Line(g)));
                            else if (st is IfStatementSyntax g2 && g2.Else != null && !(g2.Else.Statement is IfStatementSyntax))
                            {   // if (A) {...} else {...; continue;} => sau lệnh này A đúng; và ngược lại
                                bool tj = EndsWithJump(g2.Statement), ej = EndsWithJump(g2.Else.Statement);
                                if (ej && !tj) gc.Add((R(g2.Condition), Line(g2)));
                                else if (tj && !ej) gc.Add((Neg(g2.Condition), Line(g2)));
                            }
                        }
                        gc.Reverse();
                        foreach (var (t, l) in gc) { conds.Add(t); lines.Add(l); }
                        break;
                    }
            }
            while (lines.Count < conds.Count) lines.Add(Line(p));
        }
        conds.Reverse();
        lines.Reverse();
        LastCondLines = lines;
        return conds;
    }
    static bool EndsWithJump(StatementSyntax s)
    {
        if (s is BlockSyntax b) s = b.Statements.LastOrDefault();
        return s is ReturnStatementSyntax || s is ContinueStatementSyntax || s is BreakStatementSyntax || s is ThrowStatementSyntax;
    }

    // ---------- thông điệp ----------
    object MsgInfo(InvocationExpressionSyntax gm)
    {
        var a = gm.ArgumentList.Arguments.Select(x => x.Expression).ToList();
        var pos = new[] { Line(gm), gm.GetLocation().GetLineSpan().StartLinePosition.Character + 1 };
        var d = new SortedDictionary<string, object>(StringComparer.Ordinal) { ["gm_pos"] = pos, ["api"] = IsGM(gm) ? "HTKKFile.GetMessage" : "XMLMessage.GetMessage" };
        if (a.Count == 1)
        {
            var s = StrOf(a[0]);
            d["kind"] = "xml_id"; d["id"] = s; if (s == null) d["reason"] = "id không phải hằng chuỗi: " + R(a[0]);
        }
        else if (a.Count == 2)
        {
            var arr = ArrayOf(a[1]);
            d["kind"] = "cell_text"; d["sheet"] = "last";
            d["id"] = arr != null && arr.Length >= 2 && arr[0] >= 0 && arr[1] >= 0 ? ColLetters(arr[0]) + "_" + (arr[1] + 1) : null;
            if (d["id"] == null) d["reason"] = "ô thông điệp không gập được hằng: " + R(a[1]);
        }
        else if (a.Count == 3)
        {
            var r = Eval(a[1]); var c = Eval(a[2]);
            d["kind"] = "cell_text"; d["sheet"] = "Header";
            d["id"] = r != null && c != null ? ColLetters(c.Value) + "_" + (r.Value + 1) : null;
            if (d["id"] == null) d["reason"] = "ô Header không gập được hằng";
        }
        else { d["kind"] = "dynamic"; d["reason"] = "số tham số lạ"; }
        return d;
    }
    IEnumerable<InvocationExpressionSyntax> GMsIn(ExpressionSyntax e, int depth, HashSet<string> seen, List<string> notes)
    {
        if (e == null || depth > 4) yield break;
        foreach (var inv in e.DescendantNodesAndSelf().OfType<InvocationExpressionSyntax>().Where(IsGM)) yield return inv;
        foreach (var id in e.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>())
        {
            if (id.Parent is MemberAccessExpressionSyntax pm && pm.Name == id) continue;
            if (id.Parent is InvocationExpressionSyntax) continue;
            string n = id.Identifier.Text;
            if (!seen.Add(n + "@" + id.SpanStart)) continue;
            var r = Reaching(n, id.SpanStart);
            if (r == null) { if (paramNames.Contains(n)) notes.Add("param:" + n); continue; }
            if (r.Value.count > 1) notes.Add("multi:" + n);
            foreach (var x in GMsIn(r.Value.val, depth + 1, seen, notes)) yield return x;
        }
    }
    static bool IsGM(InvocationExpressionSyntax inv) => inv.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "GetMessage" && ma.Expression.ToString() == "HTKKFile";
    static bool IsXmlGM(InvocationExpressionSyntax inv) => inv.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "GetMessage" && ma.Expression.ToString() != "HTKKFile" && inv.ArgumentList.Arguments.Count == 1;

    // thông điệp không qua HTKKFile.GetMessage: XMLMessage.GetMessage("id") hoặc đọc thẳng ô của sheet cuối / sheet "Header"
    List<object> OtherMsgs(ExpressionSyntax e, int depth, HashSet<string> seen)
    {
        var res = new List<object>();
        if (e == null || depth > 4) return res;
        foreach (var inv in e.DescendantNodesAndSelf().OfType<InvocationExpressionSyntax>().Where(IsXmlGM))
            res.Add(MsgInfo(inv));
        foreach (var ea in e.DescendantNodesAndSelf().OfType<ElementAccessExpressionSyntax>())
        {
            if (!(ea.Expression is MemberAccessExpressionSyntax cm && cm.Name.Identifier.Text == "Cells" && ea.ArgumentList.Arguments.Count == 2)) continue;
            if (!(cm.Expression is ElementAccessExpressionSyntax se && se.Expression is MemberAccessExpressionSyntax sm && sm.Name.Identifier.Text == "Sheets" && se.ArgumentList.Arguments.Count == 1)) continue;
            var sarg = se.ArgumentList.Arguments[0].Expression;
            string sheet = Regex.IsMatch(Norm(sarg.ToString()), @"Count-1$") ? "last" : StrOf(sarg) == "Header" ? "Header" : null;
            if (sheet == null) continue;
            var r = Eval(ea.ArgumentList.Arguments[0].Expression); var c = Eval(ea.ArgumentList.Arguments[1].Expression);
            var d = new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["gm_pos"] = new[] { Line(ea), ea.GetLocation().GetLineSpan().StartLinePosition.Character + 1 },
                ["api"] = "direct_cell_read", ["kind"] = "cell_text", ["sheet"] = sheet,
                ["id"] = r != null && c != null ? ColLetters(c.Value) + "_" + (r.Value + 1) : null,
            };
            if (d["id"] == null) d["reason"] = "ô thông điệp không gập được hằng";
            res.Add(d);
        }
        if (res.Count > 0) return res;
        foreach (var id in e.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>())
        {
            if (id.Parent is MemberAccessExpressionSyntax pm && pm.Name == id) continue;
            if (!seen.Add(id.Identifier.Text + "@" + id.SpanStart)) continue;
            var rr = Reaching(id.Identifier.Text, id.SpanStart);
            if (rr != null && rr.Value.val != null) res.AddRange(OtherMsgs(rr.Value.val, depth + 1, seen));
        }
        return res;
    }

    public List<object> Unused = new();
    static readonly HashSet<string> Collect = new(StringComparer.Ordinal) { "Add", "Insert", "AddRange", "Push", "Enqueue", "TryAdd" };
    // mẫu thông điệp ghép: phần GetMessage/ô thông điệp thành {ID}, phần tính lúc chạy thành {biểu thức}
    string RTemplate(ExpressionSyntax e, int depth)
    {
        if (e == null) return "";
        if (depth > 4) return "{" + R(e) + "}";
        switch (e)
        {
            case ParenthesizedExpressionSyntax p: return RTemplate(p.Expression, depth);
            case LiteralExpressionSyntax l when l.IsKind(SyntaxKind.StringLiteralExpression): return l.Token.ValueText;
            case BinaryExpressionSyntax b when b.IsKind(SyntaxKind.AddExpression): return RTemplate(b.Left, depth) + RTemplate(b.Right, depth);
            case InterpolatedStringExpressionSyntax isx:
                return string.Concat(isx.Contents.Select(c => c is InterpolatedStringTextSyntax t ? t.TextToken.ValueText : c is InterpolationSyntax ip ? RTemplate(ip.Expression, depth + 1) : ""));
            case InvocationExpressionSyntax gi when IsGM(gi):
                { var mi = (SortedDictionary<string, object>)MsgInfo(gi); return "{" + (mi.TryGetValue("id", out var idv) && idv != null ? idv : "message") + "}"; }
            case InvocationExpressionSyntax fi when fi.Expression is MemberAccessExpressionSyntax fm && fm.Expression.ToString() is "string" or "String" && fm.Name.Identifier.Text == "Concat":
                return string.Concat(fi.ArgumentList.Arguments.Select(x => RTemplate(x.Expression, depth + 1)));
            case InvocationExpressionSyntax ti when ti.Expression is MemberAccessExpressionSyntax tm && tm.Name.Identifier.Text == "Trim" && ti.ArgumentList.Arguments.Count == 0:
                return RTemplate(tm.Expression, depth);
            case InvocationExpressionSyntax ff when ff.Expression is MemberAccessExpressionSyntax fmm && fmm.Expression.ToString() is "string" or "String" && fmm.Name.Identifier.Text == "Format" && ff.ArgumentList.Arguments.Count >= 1:
                {
                    var fa = ff.ArgumentList.Arguments.Select(x => x.Expression).ToList();
                    string fmt = RTemplate(fa[0], depth + 1);
                    for (int i = 1; i < fa.Count; i++) fmt = Regex.Replace(fmt, "\\{" + (i - 1) + "(:[^}]*)?\\}", "{" + R(fa[i]).Replace("$", "$$") + "}");
                    return fmt;
                }
            case IdentifierNameSyntax id:
                {
                    var sd = SingleDef(id.Identifier.Text, id.SpanStart, true);
                    if (sd != null) return RTemplate(sd, depth + 1);
                    return "{" + id.Identifier.Text + "}";
                }
        }
        var om = OtherMsgs(e, 0, new HashSet<string>(StringComparer.Ordinal));
        if (om.Count == 1 && om[0] is SortedDictionary<string, object> od && od.TryGetValue("id", out var oid) && oid != null && e is not BinaryExpressionSyntax) return "{" + oid + "}";
        return "{" + R(e) + "}";
    }
    static readonly HashSet<string> NotHelper = new(StringComparer.Ordinal) { "Show", "ShowDialog", "Format", "Concat", "Add", "Replace", "Trim", "Equals", "Contains", "ToString", "Split", "Insert", "Append", "AppendLine", "Join", "Log" };
    static string CalleeOf(InvocationExpressionSyntax i) => i.Expression is MemberAccessExpressionSyntax m ? m.Name.Identifier.Text : i.Expression is IdentifierNameSyntax id ? id.Identifier.Text : i.Expression.ToString();
    string UseOf(string name, int after)
    {
        var uses = M.DescendantNodes().OfType<IdentifierNameSyntax>().Where(i => i.Identifier.Text == name && i.SpanStart > after).ToList();
        if (uses.Count == 0) return "gán biến " + name + "; biến không được dùng trong hàm";
        var callees = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var u in uses)
            for (var q = u.Parent; q != null && q != M; q = q.Parent)
                if (q is InvocationExpressionSyntax qi) { callees.Add(qi.Expression is MemberAccessExpressionSyntax qm ? qm.Name.Identifier.Text : qi.Expression.ToString()); break; }
                else if (q is StatementSyntax) { callees.Add("(" + q.Kind() + ")"); break; }
        return "gán biến " + name + "; dùng tại: " + string.Join(", ", callees);
    }

    void ClearBind() { bindings.Clear(); bindInfo.Clear(); symKey.Clear(); keySym.Clear(); inProgress.Clear(); }
    public List<object> Extract()
    {
        var checks = new List<object>();
        var linked = new HashSet<int>();
        var gms = M.DescendantNodes().OfType<InvocationExpressionSyntax>().Where(IsGM).ToList();
        // 1) đánh dấu ô
        var marks = new List<(SyntaxNode node, string api, string level, ExpressionSyntax sheet, ExpressionSyntax row, ExpressionSyntax col, ExpressionSyntax msg)>();
        foreach (var n in M.DescendantNodes())
        {
            if (n is InvocationExpressionSyntax inv)
            {
                string nm = inv.Expression is MemberAccessExpressionSyntax ma ? ma.Name.Identifier.Text : inv.Expression is IdentifierNameSyntax id ? id.Identifier.Text : null;
                if (nm == "SetNodeBackGroundPink" || nm == "SetNodeBackGroundPink2" || nm == "SetNodeBackGroundYellow")
                {
                    var a = inv.ArgumentList.Arguments.Select(x => x.Expression).ToList();
                    string owner = inv.Expression is MemberAccessExpressionSyntax m2 ? m2.Expression.ToString() : C.Name;
                    string api = owner + "." + nm + "/" + a.Count;
                    marks.Add((inv, api, nm.Contains("Pink") ? "pink" : "yellow", a.ElementAtOrDefault(0), a.ElementAtOrDefault(1), a.ElementAtOrDefault(2), a.ElementAtOrDefault(3)));
                }
            }
            else if (n is AssignmentExpressionSyntax asg && asg.Left is MemberAccessExpressionSyntax lm && lm.Name.Identifier.Text == "BackColor"
                     && lm.Expression is ElementAccessExpressionSyntax lea && lea.Expression is MemberAccessExpressionSyntax lc && lc.Name.Identifier.Text == "Cells" && lea.ArgumentList.Arguments.Count == 2)
            {
                string col = asg.Right.ToString();
                if (col is "Color.Pink" or "Color.LightPink" or "Color.Yellow" or "Color.LemonChiffon")
                    marks.Add((asg, "direct BackColor=" + col, "other", lc.Expression, lea.ArgumentList.Arguments[0].Expression, lea.ArgumentList.Arguments[1].Expression, null));
            }
        }
        foreach (var mk in marks)
        {
            ClearBind();
            var notes = new List<string>();
            var msgs = new List<object>();
            if (mk.msg != null)
                foreach (var g in GMsIn(mk.msg, 0, new HashSet<string>(StringComparer.Ordinal), notes).Distinct())
                { linked.Add(g.SpanStart); msgs.Add(MsgInfo(g)); }
            var message = new SortedDictionary<string, object>(StringComparer.Ordinal);
            if (mk.msg == null) { message["kind"] = "literal"; message["text"] = ""; message["reason"] = "gán màu trực tiếp, không có thông điệp"; }
            else if (msgs.Count > 0)
            {
                message["parts"] = msgs;
                bool composite = !(mk.msg is InvocationExpressionSyntax mi && IsGM(mi)) && !(mk.msg is IdentifierNameSyntax);
                if (composite) { message["composite"] = true; message["template"] = RTemplate(mk.msg, 0); }
            }
            else if (OtherMsgs(mk.msg, 0, new HashSet<string>(StringComparer.Ordinal)) is var om && om.Count > 0)
            {
                message["parts"] = om;
                if (!(mk.msg is IdentifierNameSyntax) && !(mk.msg is MemberAccessExpressionSyntax) && !(mk.msg is InvocationExpressionSyntax ti && ti.Expression is MemberAccessExpressionSyntax tm && tm.Name.Identifier.Text is "Trim" or "GetMessage")) { message["composite"] = true; message["template"] = RTemplate(mk.msg, 0); }
            }
            else
            {
                var s = StrOf(mk.msg);
                if (s != null) { message["kind"] = "literal"; message["text"] = s; }
                else { message["kind"] = "dynamic"; message["expr"] = R(mk.msg); }
            }
            if (notes.Count > 0) message["notes"] = notes.Distinct().OrderBy(x => x, StringComparer.Ordinal).ToList();
            LastSheet = null;
            string cell = mk.row != null && mk.col != null ? CellRef(mk.sheet, mk.row, mk.col) : "?";
            string cellSheet = LastSheet;
            var cond = Conditions(mk.node);
            var condLines = new List<int>(LastCondLines);
            var sp = mk.node.GetLocation().GetLineSpan();
            checks.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = mk.api, ["level"] = mk.level,
                ["cells_marked"] = new[] { cell },
                ["cell_const"] = mk.row != null && Eval(mk.row) != null && Eval(mk.col) != null,
                ["sheet"] = cellSheet,
                ["condition"] = cond.Count == 0 ? "always" : string.Join(" AND ", cond.Select(WrapOr)),
                ["condition_lines"] = condLines,
                ["bindings"] = new SortedDictionary<string, string>(bindings, StringComparer.Ordinal),
                ["binding_info"] = new SortedDictionary<string, object>(bindInfo, StringComparer.Ordinal),
                ["message"] = message,
                ["lines"] = new[] { sp.StartLinePosition.Line + 1, sp.EndLinePosition.Line + 1 },
                ["col"] = sp.StartLinePosition.Character + 1,
            });
        }
        // 2) GetMessage không nối được tới lệnh đánh dấu ô
        foreach (var g in gms)
        {
            if (linked.Contains(g.SpanStart)) continue;
            ClearBind();
            string ctx = "khác";
            for (var p = g.Parent; p != null && p != M; p = p.Parent)
            {
                if (p is InvocationExpressionSyntax pi && !IsGM(pi)) { ctx = "đối số của " + (pi.Expression is MemberAccessExpressionSyntax pm ? pm.Name.Identifier.Text : pi.Expression.ToString()); break; }
                if (p is ReturnStatementSyntax) { ctx = "giá trị trả về"; break; }
                if (p is VariableDeclaratorSyntax vd) { ctx = UseOf(vd.Identifier.Text, vd.SpanStart); break; }
                if (p is AssignmentExpressionSyntax pa) { ctx = "gán " + R(pa.Left); break; }
                if (p is StatementSyntax) break;
            }
            var cond = Conditions(g);
            var condLines = new List<int>(LastCondLines);
            var sp = g.GetLocation().GetLineSpan();
            var msg = (SortedDictionary<string, object>)MsgInfo(g);
            string helper = null;
            var vdecl = g.Ancestors().TakeWhile(a => !(a is StatementSyntax)).OfType<VariableDeclaratorSyntax>().FirstOrDefault();
            var dinv = g.Ancestors().TakeWhile(a => !(a is StatementSyntax)).OfType<InvocationExpressionSyntax>().FirstOrDefault(a => !IsGM(a));
            if (vdecl != null)
            {
                var uses = M.DescendantNodes().OfType<IdentifierNameSyntax>().Where(i => i.Identifier.Text == vdecl.Identifier.Text && i.SpanStart > vdecl.SpanStart).ToList();
                if (uses.Count == 0)
                {
                    msg["line"] = sp.StartLinePosition.Line + 1; msg["variable"] = vdecl.Identifier.Text; msg["usage"] = "biến không được dùng";
                    Unused.Add(msg);
                    continue;
                }
                var uinv = uses.Select(u => u.Ancestors().TakeWhile(a => !(a is StatementSyntax)).OfType<InvocationExpressionSyntax>().FirstOrDefault()).ToList();
                if (uinv.All(x => x != null && Collect.Contains(CalleeOf(x))))
                {
                    msg["line"] = sp.StartLinePosition.Line + 1; msg["variable"] = vdecl.Identifier.Text; msg["usage"] = "chỉ đưa vào tập hợp (" + CalleeOf(uinv[0]) + ")";
                    Unused.Add(msg);
                    continue;
                }
                var hinv = uinv.FirstOrDefault(x => x != null);
                if (hinv != null && !NotHelper.Contains(CalleeOf(hinv))) { helper = CalleeOf(hinv); cond = Conditions(hinv); condLines = new List<int>(LastCondLines); condLines.Insert(0, Line(hinv)); }
            }
            else if (dinv != null && Collect.Contains(CalleeOf(dinv)))
            {
                msg["line"] = sp.StartLinePosition.Line + 1; msg["variable"] = null; msg["usage"] = "chỉ đưa vào tập hợp (" + CalleeOf(dinv) + ")";
                Unused.Add(msg);
                continue;
            }
            else if (dinv != null && !NotHelper.Contains(CalleeOf(dinv))) { helper = CalleeOf(dinv); condLines.Insert(0, Line(dinv)); }
            if (helper != null) cond.Insert(0, "via helper " + helper);
            checks.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = "HTKKFile.GetMessage", ["level"] = "other", ["cells_marked"] = Array.Empty<string>(), ["cell_const"] = true,
                ["sheet"] = null,
                ["condition"] = cond.Count == 0 ? "always" : string.Join(" AND ", cond.Select(WrapOr)),
                ["condition_lines"] = condLines,
                ["bindings"] = new SortedDictionary<string, string>(bindings, StringComparer.Ordinal),
                ["binding_info"] = new SortedDictionary<string, object>(bindInfo, StringComparer.Ordinal),
                ["message"] = new SortedDictionary<string, object>(StringComparer.Ordinal) { ["parts"] = new[] { msg } },
                ["context"] = ctx,
                ["lines"] = new[] { sp.StartLinePosition.Line + 1, sp.EndLinePosition.Line + 1 },
                ["col"] = sp.StartLinePosition.Character + 1,
            });
        }
        // 3) hộp thoại frmMessage("ID", ...): ID tra trong message.xml (chỉ assembly business)
        foreach (var oc in C.Assembly == "tax.htkk.business" ? M.DescendantNodes().OfType<ObjectCreationExpressionSyntax>() : Enumerable.Empty<ObjectCreationExpressionSyntax>())
        {
            if (oc.Type.ToString() != "frmMessage" || oc.ArgumentList == null || oc.ArgumentList.Arguments.Count == 0) continue;
            ClearBind();
            var oa = oc.ArgumentList.Arguments.Select(x => x.Expression).ToList();
            var sp = oc.GetLocation().GetLineSpan();
            var part = new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = "frmMessage",
                ["gm_pos"] = new[] { sp.StartLinePosition.Line + 1, sp.StartLinePosition.Character + 1 },
            };
            if (oa.Count == 3)
            {   // frmMessage(kiểu, tiêu đề, nội dung): nội dung truyền thẳng
                var lit = StrOf(oa[2]);
                part["kind"] = lit != null ? "literal" : "dynamic";
                if (lit != null) part["text"] = lit; else part["expr"] = R(oa[2]);
                if (lit == null) part["reason"] = "nội dung hộp thoại tính lúc chạy";
            }
            else
            {   // frmMessage(ID, [nội dung thêm,] kiểu, biểu tượng, tiêu đề)
                var id = StrOf(oa[0]);
                part["kind"] = "xml_id"; part["id"] = id;
                if (id == null) part["reason"] = "id không phải hằng chuỗi: " + R(oa[0]);
                if (oa.Count == 5) part["extra"] = R(oa[1]);
            }
            var cond = Conditions(oc);
            var condLines = new List<int>(LastCondLines);
            checks.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = "frmMessage", ["level"] = "other", ["cells_marked"] = Array.Empty<string>(), ["cell_const"] = true,
                ["sheet"] = null,
                ["condition"] = cond.Count == 0 ? "always" : string.Join(" AND ", cond.Select(WrapOr)),
                ["bindings"] = new SortedDictionary<string, string>(bindings, StringComparer.Ordinal),
                ["binding_info"] = new SortedDictionary<string, object>(bindInfo, StringComparer.Ordinal),
                ["message"] = new SortedDictionary<string, object>(StringComparer.Ordinal) { ["parts"] = new[] { part } },
                ["context"] = "hộp thoại frmMessage/" + oa.Count + " tham số",
                ["condition_lines"] = condLines,
                ["lines"] = new[] { sp.StartLinePosition.Line + 1, sp.EndLinePosition.Line + 1 },
                ["col"] = sp.StartLinePosition.Character + 1,
            });
        }
        return checks.OrderBy(c => ((int[])((SortedDictionary<string, object>)c)["lines"])[0]).ThenBy(c => (int)((SortedDictionary<string, object>)c)["col"]).ToList();
    }
}
