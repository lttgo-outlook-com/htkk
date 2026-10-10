// Trích lệnh kiểm tra dữ liệu (đánh dấu ô Pink/Yellow, lời gọi HTKKFile.GetMessage) từ mã decompile HTKK.
// Chỉ đọc text .cs và gọi CSharpSyntaxTree.ParseText. Không biên dịch, không chạy mã decompile.
// Dùng: RoslynValidations <decompiledDir> <out.json>
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

if (args.Length < 2) { Console.Error.WriteLine("Dùng: RoslynValidations <decompiledDir> <out.json>"); return 2; }
string decDir = Path.GetFullPath(args[0]);
string outPath = args[1];

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
        var cc = new ClassCtx { File = rel, Name = td.Identifier.Text, Decl = td, Assembly = rel.Split('/')[0] };
        foreach (var fd in td.Members.OfType<FieldDeclarationSyntax>())
            foreach (var v in fd.Declaration.Variables)
                if (v.Initializer != null) cc.Fields[v.Identifier.Text] = v.Initializer.Value;
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

// ---- xuất ----
var outClasses = new List<object>();
foreach (var c in classes.OrderBy(c => c.File, StringComparer.Ordinal).ThenBy(c => c.Name, StringComparer.Ordinal))
{
    var methods = new List<object>();
    foreach (var md in c.Methods)
    {
        string name = md.Identifier.Text;
        string key = c.Name + "::" + name;
        string kind = Kind(name);
        bool scope = inScope.ContainsKey(key);
        bool signal = hasSignal.Contains(key);
        bool rootish = kind == "Validate" || kind == "isValidDuLieuTK" || kind == "checkPL";
        if (!signal && !rootish) continue;
        var mx = new MethodCtx(c, md);
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
        });
    }
    if (methods.Count == 0) continue;
    outClasses.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
    {
        ["file"] = c.File, ["class"] = c.Name, ["assembly"] = c.Assembly, ["methods"] = methods,
    });
}
var jo = new JsonSerializerOptions { WriteIndented = false, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
File.WriteAllText(outPath, JsonSerializer.Serialize(new { classes = outClasses }, jo) + "\n", new UTF8Encoding(false));
Console.WriteLine($"classes_out={outClasses.Count}");
return 0;

class ClassCtx
{
    public string File, Name, Assembly;
    public ClassDeclarationSyntax Decl;
    public Dictionary<string, ExpressionSyntax> Fields = new(StringComparer.Ordinal);
    public List<MethodDeclarationSyntax> Methods = new();
}

class MethodCtx
{
    readonly ClassCtx C;
    readonly MethodDeclarationSyntax M;
    readonly List<(string key, int pos, ExpressionSyntax val, int line)> assigns = new();
    readonly HashSet<string> paramNames;
    public MethodCtx(ClassCtx c, MethodDeclarationSyntax m)
    {
        C = c; M = m;
        paramNames = m.ParameterList.Parameters.Select(p => p.Identifier.Text).ToHashSet(StringComparer.Ordinal);
        foreach (var n in m.DescendantNodes())
        {
            if (n is VariableDeclaratorSyntax vd && vd.Initializer != null)
                assigns.Add((vd.Identifier.Text, vd.SpanStart, vd.Initializer.Value, Line(vd)));
            else if (n is AssignmentExpressionSyntax a && a.IsKind(SyntaxKind.SimpleAssignmentExpression) && (a.Left is IdentifierNameSyntax || a.Left is ElementAccessExpressionSyntax))
                assigns.Add((Norm(a.Left.ToString()), a.SpanStart, a.Right, Line(a)));
            else if (n is ArgumentSyntax arg && arg.RefKindKeyword.IsKind(SyntaxKind.OutKeyword))
            {
                string k = arg.Expression is DeclarationExpressionSyntax de && de.Designation is SingleVariableDesignationSyntax sv ? sv.Identifier.Text : Norm(arg.Expression.ToString());
                assigns.Add((k, arg.SpanStart, null, Line(arg)));
            }
            else if (n is ForEachStatementSyntax fe)
                assigns.Add((fe.Identifier.Text, fe.SpanStart, null, Line(fe)));
        }
    }
    static int Line(SyntaxNode n) => n.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
    static string Norm(string s) => Regex.Replace(s, @"\s+", "");

    (ExpressionSyntax val, int line, int count)? Reaching(string key, int pos)
    {
        (ExpressionSyntax, int, int)? best = null; int bpos = -1, cnt = 0;
        foreach (var a in assigns) if (a.key == key) { cnt++; if (a.pos < pos && a.pos > bpos) { bpos = a.pos; best = (a.val, a.line, 0); } }
        if (best == null) return null;
        return (best.Value.Item1, best.Value.Item2, cnt);
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
                    var r = Reaching(n, e.SpanStart);
                    if (r != null) return r.Value.val == null || InLoopAssigned(n) ? null : Eval(r.Value.val, depth + 1);
                    if (paramNames.Contains(n)) return null;
                    if (C.Fields.TryGetValue(n, out var fe)) return Eval(fe, depth + 1);
                    return null;
                }
            case ElementAccessExpressionSyntax ea when ea.ArgumentList.Arguments.Count == 1:
                {
                    var idx = Eval(ea.ArgumentList.Arguments[0].Expression, depth + 1);
                    var arr = ArrayOf(ea.Expression, depth + 1);
                    if (idx != null && arr != null && idx >= 0 && idx < arr.Length) return arr[idx.Value];
                    var r = Reaching(Norm(ea.ToString()), e.SpanStart);
                    if (r != null && r.Value.val != null) return Eval(r.Value.val, depth + 1);
                    return null;
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
                    var r = Reaching(id.Identifier.Text, e.SpanStart);
                    if (r != null) return r.Value.val == null ? null : ArrayOf(r.Value.val, depth + 1);
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
            case ConditionalAccessExpressionSyntax ca: return RChain(ca.WhenNotNull, R(ca.Expression) + "?");
            case IdentifierNameSyntax id: return RIdent(id);
            case ThisExpressionSyntax: return "this";
            case ElementAccessExpressionSyntax ea: return RElement(ea);
            case MemberAccessExpressionSyntax ma: return RMember(ma);
            case InvocationExpressionSyntax inv: return RInvoke(inv);
            case ObjectCreationExpressionSyntax oc: return "new " + oc.Type + "(" + string.Join(", ", (oc.ArgumentList?.Arguments ?? default).Select(a => R(a.Expression))) + ")";
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
            case ConditionalAccessExpressionSyntax c2: return RChain(c2.WhenNotNull, RChain(c2.Expression, prefix) + "?");
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
        if (paramNames.Contains(n)) return n;
        var r = Reaching(n, id.SpanStart);
        if (r != null && (r.Value.val is ArrayCreationExpressionSyntax ac0 && ac0.Initializer == null || r.Value.val is ObjectCreationExpressionSyntax)) return n; // tập hợp có thể bị sửa sau khi tạo: giữ tên
        if (r != null && r.Value.val != null && inlineDepth < 2 && !InLoopAssigned(n) && r.Value.count == 1)
        {
            inlineDepth++;
            string s = R(r.Value.val);
            inlineDepth--;
            if (s.Length <= 140) return s.Contains(' ') ? "(" + s + ")" : s;
            Bind(n, s, r.Value.line);
            return n;
        }
        if (r != null && r.Value.val != null && inlineDepth < 2)
        {
            inlineDepth++; string s = R(r.Value.val); inlineDepth--;
            if (r.Value.count > 1) s = "<gán " + r.Value.count + " lần; phép gán gần nhất phía trước> " + s;
            else if (InLoopAssigned(n)) s = "<gán trong vòng lặp> " + s;
            Bind(n, s, r.Value.line);
        }
        if (C.Fields.ContainsKey(n) && StrOf(id) is string fs) return JsonSerializer.Serialize(fs);
        return n;
    }
    void Bind(string n, string s, int line) { if (s.Length > 400) s = s.Substring(0, 400) + "…"; bindings[n] = s + " @L" + line; }

    string RElement(ElementAccessExpressionSyntax ea)
    {
        var args = ea.ArgumentList.Arguments;
        if (ea.Expression is MemberAccessExpressionSyntax ma && ma.Name.Identifier.Text == "Cells" && args.Count == 2)
            return "cell(" + CellRef(ma.Expression, args[0].Expression, args[1].Expression) + ")";
        if (ea.Expression is MemberAccessExpressionSyntax ms && ms.Name.Identifier.Text == "Sheets" && args.Count == 1)
            return "sheet(" + (SheetOf(ea) ?? R(args[0].Expression)) + ")";
        var v = Eval(ea); if (v != null) return v.Value.ToString();
        var r = Reaching(Norm(ea.ToString()), ea.SpanStart);
        if (r != null && r.Value.val != null && inlineDepth < 2)
        {
            inlineDepth++; string s = R(r.Value.val); inlineDepth--;
            if (s.Length <= 140) return s.Contains(' ') ? "(" + s + ")" : s;
        }
        return R(ea.Expression) + "[" + string.Join(", ", args.Select(a => R(a.Expression))) + "]";
    }
    public string CellRef(ExpressionSyntax sheetExpr, ExpressionSyntax row, ExpressionSyntax col)
    {
        string sh = SheetOf(sheetExpr);
        string prefix = sh != null && !sh.StartsWith("param:") && sh != "active" ? sh + "!" : "";
        var r = Eval(row); var c = Eval(col);
        string cs = c != null ? ColLetters(c.Value) : "col{" + R(col) + "}";
        string rs = r != null ? (r.Value + 1).ToString() : "{" + R(row) + " +1}";
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
        if (name == "Length" || name == "Count") return "length(" + R(ma.Expression) + ")";
        if (name == "Visible" && ma.Expression is ElementAccessExpressionSyntax) return "visible(" + R(ma.Expression) + ")";
        if (ma.Expression is PredefinedTypeSyntax || ma.Expression is IdentifierNameSyntax q && char.IsUpper(q.Identifier.Text[0]) && !paramNames.Contains(q.Identifier.Text) && Reaching(q.Identifier.Text, ma.SpanStart) == null && !C.Fields.ContainsKey(q.Identifier.Text))
            return ma.Expression + "." + name; // hằng/thuộc tính tĩnh (Color.Pink, string.Empty, ...)
        var v = Eval(ma); if (v != null) return v.Value.ToString();
        return R(ma.Expression) + "." + name;
    }
    static readonly Dictionary<string, string> FnMap = new(StringComparer.Ordinal)
    {
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

    // ---------- điều kiện ----------
    string Neg(ExpressionSyntax c)
    {
        while (c is ParenthesizedExpressionSyntax pp) c = pp.Expression;
        if (c is PrefixUnaryExpressionSyntax pu && pu.IsKind(SyntaxKind.LogicalNotExpression)) return R(pu.Operand);
        return "NOT (" + R(c) + ")";
    }
    List<string> Conditions(SyntaxNode n)
    {
        var conds = new List<string>();
        SyntaxNode child = n;
        for (var p = n.Parent; p != null && p != M; child = p, p = p.Parent)
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
                    // điều kiện chặn sớm: if (X) { ...; return|continue|break; } đứng trước trong cùng block
                    foreach (var st in blk.Statements)
                    {
                        if (st.SpanStart >= child.SpanStart) break;
                        if (st is IfStatementSyntax g && g.Else == null && EndsWithJump(g.Statement)) conds.Add(Neg(g.Condition));
                        else if (st is IfStatementSyntax g2 && g2.Else != null && !(g2.Else.Statement is IfStatementSyntax))
                        {   // if (A) {...} else {...; continue;} => sau lệnh này A đúng; và ngược lại
                            bool tj = EndsWithJump(g2.Statement), ej = EndsWithJump(g2.Else.Statement);
                            if (ej && !tj) conds.Add(R(g2.Condition));
                            else if (tj && !ej) conds.Add(Neg(g2.Condition));
                        }
                    }
                    break;
            }
        }
        conds.Reverse();
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
            bindings.Clear();
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
                if (composite) message["composite"] = true;
            }
            else if (OtherMsgs(mk.msg, 0, new HashSet<string>(StringComparer.Ordinal)) is var om && om.Count > 0)
            {
                message["parts"] = om;
                if (!(mk.msg is IdentifierNameSyntax) && !(mk.msg is MemberAccessExpressionSyntax) && !(mk.msg is InvocationExpressionSyntax ti && ti.Expression is MemberAccessExpressionSyntax tm && tm.Name.Identifier.Text is "Trim" or "GetMessage")) message["composite"] = true;
            }
            else
            {
                var s = StrOf(mk.msg);
                if (s != null) { message["kind"] = "literal"; message["text"] = s; }
                else { message["kind"] = "dynamic"; message["expr"] = R(mk.msg); }
            }
            if (notes.Count > 0) message["notes"] = notes.Distinct().OrderBy(x => x, StringComparer.Ordinal).ToList();
            string cell = mk.row != null && mk.col != null ? CellRef(mk.sheet, mk.row, mk.col) : "?";
            var cond = Conditions(mk.node);
            var sp = mk.node.GetLocation().GetLineSpan();
            checks.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = mk.api, ["level"] = mk.level,
                ["cells_marked"] = new[] { cell },
                ["cell_const"] = mk.row != null && Eval(mk.row) != null && Eval(mk.col) != null,
                ["sheet"] = SheetOf(mk.sheet),
                ["condition"] = cond.Count == 0 ? "always" : string.Join(" AND ", cond.Select(x => x.Contains(" OR ") && !x.StartsWith("NOT (") ? "(" + x + ")" : x)),
                ["bindings"] = new SortedDictionary<string, string>(bindings, StringComparer.Ordinal),
                ["message"] = message,
                ["lines"] = new[] { sp.StartLinePosition.Line + 1, sp.EndLinePosition.Line + 1 },
                ["col"] = sp.StartLinePosition.Character + 1,
            });
        }
        // 2) GetMessage không nối được tới lệnh đánh dấu ô
        foreach (var g in gms)
        {
            if (linked.Contains(g.SpanStart)) continue;
            bindings.Clear();
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
                    msg["line"] = sp.StartLinePosition.Line + 1; msg["variable"] = vdecl.Identifier.Text;
                    Unused.Add(msg);
                    continue;
                }
                var hinv = uses.Select(u => u.Ancestors().TakeWhile(a => !(a is StatementSyntax)).OfType<InvocationExpressionSyntax>().FirstOrDefault()).FirstOrDefault(x => x != null);
                if (hinv != null && !NotHelper.Contains(CalleeOf(hinv))) { helper = CalleeOf(hinv); cond = Conditions(hinv); }
            }
            else if (dinv != null && !NotHelper.Contains(CalleeOf(dinv))) helper = CalleeOf(dinv);
            if (helper != null) cond.Insert(0, "via helper " + helper);
            checks.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = "HTKKFile.GetMessage", ["level"] = "other", ["cells_marked"] = Array.Empty<string>(), ["cell_const"] = true,
                ["sheet"] = null,
                ["condition"] = cond.Count == 0 ? "always" : string.Join(" AND ", cond),
                ["bindings"] = new SortedDictionary<string, string>(bindings, StringComparer.Ordinal),
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
            bindings.Clear();
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
            checks.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["api"] = "frmMessage", ["level"] = "other", ["cells_marked"] = Array.Empty<string>(), ["cell_const"] = true,
                ["sheet"] = null,
                ["condition"] = cond.Count == 0 ? "always" : string.Join(" AND ", cond),
                ["bindings"] = new SortedDictionary<string, string>(bindings, StringComparer.Ordinal),
                ["message"] = new SortedDictionary<string, object>(StringComparer.Ordinal) { ["parts"] = new[] { part } },
                ["context"] = "hộp thoại frmMessage/" + oa.Count + " tham số",
                ["lines"] = new[] { sp.StartLinePosition.Line + 1, sp.EndLinePosition.Line + 1 },
                ["col"] = sp.StartLinePosition.Character + 1,
            });
        }
        return checks.OrderBy(c => ((int[])((SortedDictionary<string, object>)c)["lines"])[0]).ThenBy(c => (int)((SortedDictionary<string, object>)c)["col"]).ToList();
    }
}
