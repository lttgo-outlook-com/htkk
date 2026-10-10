// Thông dịch ký hiệu trên cây cú pháp Roslyn (chỉ ParseText). Không biên dịch, không chạy mã decompile.
using System.Collections.Immutable;
using System.Globalization;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RulesExtract;

public sealed class ClassCtx
{
    public string Name, File;
    public readonly Dictionary<string, List<BaseMethodDeclarationSyntax>> Methods = new(StringComparer.Ordinal);
    public readonly Dictionary<string, ExpressionSyntax> FieldInit = new(StringComparer.Ordinal);
    public readonly Dictionary<string, SyntaxNode> PropBody = new(StringComparer.Ordinal);
    public readonly HashSet<string> Mutable = new(StringComparer.Ordinal);
    public readonly HashSet<string> Members = new(StringComparer.Ordinal);
    public readonly Dictionary<string, V> Cache = new(StringComparer.Ordinal);
    public readonly HashSet<string> Evaluating = new(StringComparer.Ordinal);
    public readonly Dictionary<string, string> FieldType = new(StringComparer.Ordinal);
}

public sealed class Entry { public readonly V Val; public readonly ImmutableSortedSet<int> Lines; public Entry(V v, ImmutableSortedSet<int> l) { Val = v; Lines = l; } }
public sealed record PathItem(Cond C, int Line);
public sealed record LoopInfo(string Desc, int Line);

public sealed class State
{
    public Dictionary<string, Entry> Env;
    public ImmutableList<PathItem> Path;
    public ImmutableList<LoopInfo> Loops;
    public bool Dead, Threw;
    public State Clone() => new() { Env = new Dictionary<string, Entry>(Env, StringComparer.Ordinal), Path = Path, Loops = Loops, Dead = Dead, Threw = Threw };
}

public sealed class Frame
{
    public ClassCtx Cls; public string Method; public int Depth; public string Via;
    public List<(ImmutableList<PathItem> Path, V Val)> Returns = new();
    public ImmutableHashSet<string> Stack = ImmutableHashSet<string>.Empty;
    public int BasePathLen;
}

public sealed class Effect
{
    public string Op; public V Target; public V Val; public SortedDictionary<string, object> Params = new(StringComparer.Ordinal);
    public ImmutableList<PathItem> Path; public ImmutableList<LoopInfo> Loops; public int Line; public string Via; public string Point;
    public ImmutableSortedSet<int> Lines;
}

public sealed class Interp
{
    public readonly List<Effect> Effects = new();
    public readonly HashSet<string> OpaqueSites; // "name\tfile:line\tkind"
    public readonly Func<string, ClassCtx> NoClass;
    readonly string file;
    HashSet<int> used = new();
    int steps;
    public bool Truncated;
    public const int MaxSteps = 400000, MaxDepth = 4;
    public int HookSheet = -1;
    public static int CurSheet = -1;
    public readonly List<(int Line, string Reason)> Dropped = new();
    bool pendingThrow;
    void CheckThrow(State st, int line, string what)
    {
        if (!pendingThrow) return;
        pendingThrow = false; Dropped.Add((line, "ngoại lệ chắc chắn: " + what)); st.Dead = true; st.Threw = true;
    }

    public Interp(string file, HashSet<string> opaqueSites) { this.file = file; OpaqueSites = opaqueSites; }

    static int Line(SyntaxNode n) => n.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
    static string PointId(string file, SyntaxNode n) { var p = n.GetLocation().GetLineSpan().StartLinePosition; return file + ":" + (p.Line + 1) + ":" + (p.Character + 1); }

    void Opaque(string name, SyntaxNode at, string kind) => OpaqueSites.Add(name + "\t" + file + ":" + Line(at) + "\t" + kind);

    // ---------- tham số ----------
    public static V ParamValue(string type, string name, int sheetIdx, string hookKind, int pos, int count)
    {
        type = type.Trim();
        if (type == "FpSpread") return VSpread.I;
        if (type == "SheetView") return sheetIdx >= 0 && hookKind != null ? new VSheet(new VConst((long)sheetIdx)) : new VSheet(new VSym(name));
        if (type == "HTKKCMCToKhaiInfo") return new VSym("tk");
        if (type == "HTKKCMCTTinChungInfo") return new VSym("ttc");
        if (type == "HTKKPeriodInfo") return new VSym("period");
        if (type == "HTKKHeaderInfo") return new VSym("header");
        if (hookKind is "CellChange" or "CellEditchange" or "ButtonClick" && pos >= count - 2 && (type is "long" or "int"))
        {
            var at = VLin.Atom(pos == count - 2 ? R.TRow : R.TCol);
            // ButtonClick: app truyền idxRow/idxCol gốc 0 (so với ArrayCellId và RowIndex - 1), quy về trigger 1-based
            return hookKind == "ButtonClick" ? Ar.Add(at, new VConst(-1L)) : at;
        }
        if (type is "int" or "long" or "short") return VLin.Atom(name);
        return new VSym(name);
    }

    public void RunMethod(BaseMethodDeclarationSyntax m, ClassCtx cls, Dictionary<string, V> args, string via, Frame parent, State callerState)
    {
        var st = new State
        {
            Env = new Dictionary<string, Entry>(StringComparer.Ordinal),
            Path = callerState?.Path ?? ImmutableList<PathItem>.Empty,
            Loops = callerState?.Loops ?? ImmutableList<LoopInfo>.Empty,
        };
        foreach (var kv in args) st.Env[kv.Key] = new Entry(kv.Value, ImmutableSortedSet<int>.Empty);
        var f = new Frame { Cls = cls, Method = MethodName(m), Depth = parent == null ? 0 : parent.Depth + 1, Via = via, BasePathLen = st.Path.Count, Stack = (parent?.Stack ?? ImmutableHashSet<string>.Empty).Add(MethodName(m) + "/" + m.SpanStart) };
        _lastFrame = f;
        if (m.Body != null) Exec(m.Body, st, f);
        else if (m.ExpressionBody != null) { used = new(); var v = Eval(m.ExpressionBody.Expression, st, f); f.Returns.Add((st.Path, v)); }
    }
    Frame _lastFrame;
    public static string MethodName(BaseMethodDeclarationSyntax m) => m is MethodDeclarationSyntax mm ? mm.Identifier.Text : ".ctor";

    // ---------- câu lệnh ----------
    void Exec(StatementSyntax s, State st, Frame f)
    {
        if (st.Dead) return;
        if (++steps > MaxSteps) { Truncated = true; st.Dead = true; return; }
        switch (s)
        {
            case BlockSyntax b:
                foreach (var x in b.Statements) { if (st.Dead) break; Exec(x, st, f); }
                break;
            case LocalDeclarationStatementSyntax ld:
                foreach (var v in ld.Declaration.Variables)
                {
                    used = new();
                    V val = v.Initializer != null ? Eval(v.Initializer.Value, st, f) : new VSym(v.Identifier.Text);
                    Assign(st, v.Identifier.Text, val, Line(v));
                    CheckThrow(st, Line(v), "chỉ số ngoài mảng");
                    if (st.Dead) break;
                }
                break;
            case ExpressionStatementSyntax es:
                used = new();
                ExecExpr(es.Expression, st, f);
                CheckThrow(st, Line(es), "chỉ số ngoài mảng");
                break;
            case IfStatementSyntax ifs:
                {
                    used = new();
                    var c = C.Of(Eval(ifs.Condition, st, f));
                    int ln = Line(ifs);
                    var ts = st.Clone(); ts.Path = ts.Path.Add(new PathItem(c, ln));
                    var es2 = st.Clone(); es2.Path = es2.Path.Add(new PathItem(C.Not(c), ln));
                    if (c is CConst k) { if (k.B) es2.Dead = true; else ts.Dead = true; ts.Path = st.Path; es2.Path = st.Path; }
                    if (!ts.Dead) Exec(ifs.Statement, ts, f);
                    if (!es2.Dead && ifs.Else != null) Exec(ifs.Else.Statement, es2, f);
                    Merge(st, new List<(Cond, State)> { (c, ts), (C.Not(c), es2) }, ln);
                    break;
                }
            case ReturnStatementSyntax rs:
                if (rs.Expression != null) { used = new(); var v = Eval(rs.Expression, st, f); f.Returns.Add((st.Path, v)); }
                st.Dead = true; break;
            case ThrowStatementSyntax: st.Dead = true; st.Threw = true; break;
            case TryStatementSyntax ts2:
                Exec(ts2.Block, st, f);
                if (st.Threw) { st.Dead = false; st.Threw = false; }
                if (ts2.Finally != null) { st.Dead = false; Exec(ts2.Finally.Block, st, f); }
                break;
            case UsingStatementSyntax us:
                if (us.Declaration != null) foreach (var v in us.Declaration.Variables) { used = new(); Assign(st, v.Identifier.Text, v.Initializer != null ? Eval(v.Initializer.Value, st, f) : new VSym(v.Identifier.Text), Line(v)); }
                Exec(us.Statement, st, f); break;
            case ForStatementSyntax fs:
                {
                    var vars = new List<string>();
                    if (fs.Declaration != null) foreach (var v in fs.Declaration.Variables) { used = new(); var init = v.Initializer != null ? Eval(v.Initializer.Value, st, f) : new VSym(v.Identifier.Text); vars.Add(v.Identifier.Text + " from " + R.Expr(init, new Flags())); }
                    foreach (var ie in fs.Initializers) { used = new(); ExecExpr(ie, st, f); }
                    var assigned = AssignedIn(fs);
                    if (TryUnroll(fs, st, f)) break;
                    foreach (var n in assigned) st.Env[n] = new Entry(VLin.Atom(n), ImmutableSortedSet<int>.Empty);
                    if (fs.Declaration != null) foreach (var v in fs.Declaration.Variables) st.Env[v.Identifier.Text] = new Entry(VLin.Atom(v.Identifier.Text), ImmutableSortedSet.Create(Line(v)));
                    used = new();
                    string cond = fs.Condition != null ? R.Cond(C.Of(Eval(fs.Condition, st, f)), new Flags()) : "true";
                    Loop(st, f, fs.Statement, "for " + string.Join(", ", vars) + " while " + cond, Line(fs), assigned);
                    break;
                }
            case ForEachStatementSyntax fe:
                {
                    used = new();
                    var coll = Eval(fe.Expression, st, f);
                    var assigned = AssignedIn(fe);
                    st.Env[fe.Identifier.Text] = new Entry(new VSym(fe.Identifier.Text), ImmutableSortedSet.Create(Line(fe)));
                    Loop(st, f, fe.Statement, "foreach " + fe.Identifier.Text + " in " + R.Expr(coll, new Flags()), Line(fe), assigned);
                    break;
                }
            case WhileStatementSyntax ws:
                {
                    var assigned = AssignedIn(ws);
                    foreach (var n in assigned) st.Env[n] = new Entry(VLin.Atom(n), ImmutableSortedSet<int>.Empty);
                    used = new();
                    Loop(st, f, ws.Statement, "while " + R.Cond(C.Of(Eval(ws.Condition, st, f)), new Flags()), Line(ws), assigned);
                    break;
                }
            case DoStatementSyntax ds:
                {
                    var assigned = AssignedIn(ds);
                    foreach (var n in assigned) st.Env[n] = new Entry(VLin.Atom(n), ImmutableSortedSet<int>.Empty);
                    used = new();
                    Loop(st, f, ds.Statement, "do while " + R.Cond(C.Of(Eval(ds.Condition, st, f)), new Flags()), Line(ds), assigned);
                    break;
                }
            case SwitchStatementSyntax sw:
                {
                    used = new();
                    var g = Eval(sw.Expression, st, f);
                    int ln = Line(sw);
                    var branches = new List<(Cond, State)>();
                    var labels = new List<Cond>();
                    SwitchSectionSyntax def = null;
                    foreach (var sec in sw.Sections)
                    {
                        var cs = new List<Cond>();
                        bool isDef = false;
                        foreach (var lab in sec.Labels)
                        {
                            if (lab is CaseSwitchLabelSyntax cl) { used = new(); cs.Add(C.Cmp("==", g, Eval(cl.Value, st, f))); }
                            else if (lab is DefaultSwitchLabelSyntax) isDef = true;
                            else cs.Add(new CAtom("pattern", new VOpaque("pattern")));
                        }
                        if (isDef) { def = sec; continue; }
                        var c = C.Or(cs); labels.Add(c);
                        var bs = st.Clone(); bs.Path = bs.Path.Add(new PathItem(c, Line(sec)));
                        if (c is CConst k0 && !k0.B) bs.Dead = true;
                        foreach (var x in sec.Statements) { if (bs.Dead) break; Exec(x, bs, f); }
                        branches.Add((c, bs));
                    }
                    var dc = C.Not(C.Or(labels));
                    var ds2 = st.Clone(); ds2.Path = ds2.Path.Add(new PathItem(dc, def != null ? Line(def) : ln));
                    if (dc is CConst k1 && !k1.B) ds2.Dead = true;
                    if (def != null) foreach (var x in def.Statements) { if (ds2.Dead) break; Exec(x, ds2, f); }
                    branches.Add((dc, ds2));
                    Merge(st, branches, ln);
                    break;
                }
            case LabeledStatementSyntax ls: Exec(ls.Statement, st, f); break;
            case CheckedStatementSyntax cs3: Exec(cs3.Block, st, f); break;
            case LockStatementSyntax lk: Exec(lk.Statement, st, f); break;
            case FixedStatementSyntax fx: Exec(fx.Statement, st, f); break;
            default: break; // break/continue/goto/empty/local function: bỏ qua
        }
    }

    const int MaxUnroll = 40;
    bool TryUnroll(ForStatementSyntax fs, State st, Frame f)
    {
        if (fs.Declaration == null || fs.Declaration.Variables.Count != 1 || fs.Condition == null) return false;
        string name = fs.Declaration.Variables[0].Identifier.Text;
        var initX = fs.Declaration.Variables[0].Initializer?.Value;
        if (initX == null || AssignedIn(fs.Statement).Contains(name)) return false;
        var savedUsed = used;
        var init = Eval(initX, st, f);
        if (init is not VConst k0 || k0.Val is not long) { used = savedUsed; return false; }
        var scratch = st.Clone(); var vals = new List<V>();
        scratch.Env[name] = new Entry(init, ImmutableSortedSet<int>.Empty);
        for (int it = 0; ; it++)
        {
            used = new();
            var c = C.Of(Eval(fs.Condition, scratch, f));
            if (c is not CConst ck) { used = savedUsed; return false; }
            if (!ck.B) break;
            if (it >= MaxUnroll) { used = savedUsed; return false; }
            vals.Add(scratch.Env[name].Val);
            foreach (var inc in fs.Incrementors) ExecExpr(inc, scratch, f);
            if (scratch.Env[name].Val is not VConst) { used = savedUsed; return false; }
        }
        used = savedUsed;
        foreach (var v in vals)
        {
            if (st.Dead) break;
            st.Env[name] = new Entry(v, ImmutableSortedSet.Create(Line(fs)));
            Exec(fs.Statement, st, f);
        }
        return true;
    }

    void Loop(State st, Frame f, StatementSyntax body, string desc, int line, HashSet<string> assigned)
    {
        var bs = st.Clone(); bs.Loops = bs.Loops.Add(new LoopInfo(desc, line));
        Exec(body, bs, f);
        foreach (var n in assigned) st.Env[n] = new Entry(VLin.Atom(n), ImmutableSortedSet.Create(line));
    }

    static HashSet<string> AssignedIn(SyntaxNode n)
    {
        var h = new HashSet<string>(StringComparer.Ordinal);
        foreach (var a in n.DescendantNodes())
        {
            if (a is AssignmentExpressionSyntax ae && ae.Left is IdentifierNameSyntax id) h.Add(id.Identifier.Text);
            else if (a is PostfixUnaryExpressionSyntax pu && pu.Operand is IdentifierNameSyntax id2) h.Add(id2.Identifier.Text);
            else if (a is PrefixUnaryExpressionSyntax pr && (pr.IsKind(SyntaxKind.PreIncrementExpression) || pr.IsKind(SyntaxKind.PreDecrementExpression)) && pr.Operand is IdentifierNameSyntax id3) h.Add(id3.Identifier.Text);
            else if (a is InvocationExpressionSyntax iv && iv.Expression is MemberAccessExpressionSyntax im && im.Expression is IdentifierNameSyntax id4 && im.Name.Identifier.Text is "Add" or "Insert" or "Remove" or "RemoveAt" or "Clear" or "AddRange") h.Add(id4.Identifier.Text);
        }
        return h;
    }

    void Merge(State st, List<(Cond C, State S)> branches, int line)
    {
        var live = branches.Where(b => !b.S.Dead).ToList();
        if (live.Count == 0) { st.Dead = true; return; }
        if (live.Count == 1)
        {
            st.Env = live[0].S.Env;
            if (live.Count < branches.Count && !(live[0].C is CConst)) st.Path = st.Path.Add(new PathItem(live[0].C, line));
            return;
        }
        var env = new Dictionary<string, Entry>(StringComparer.Ordinal);
        var keys = new SortedSet<string>(live.SelectMany(b => b.S.Env.Keys), StringComparer.Ordinal);
        foreach (var k in keys)
        {
            if (!live.All(b => b.S.Env.ContainsKey(k))) { if (st.Env.ContainsKey(k)) env[k] = st.Env[k]; continue; }
            var es = live.Select(b => b.S.Env[k]).ToList();
            var lines = es.Aggregate(ImmutableSortedSet<int>.Empty, (a, e) => a.Union(e.Lines));
            if (es.All(e => ReferenceEquals(e.Val, es[0].Val) || e.Val.Key == es[0].Val.Key)) { env[k] = new Entry(es[0].Val, lines); continue; }
            var alts = new List<(Cond, V)>();
            for (int i = 0; i < live.Count; i++) alts.Add((i == live.Count - 1 ? null : live[i].C, es[i].Val));
            var v = Ar.Cases(alts);
            if (v is VOpaque) v = VLin.Atom(k);
            env[k] = new Entry(v, lines.Add(line));
        }
        st.Env = env;
    }

    void Assign(State st, string name, V v, int line)
    {
        st.Env[name] = new Entry(v, ImmutableSortedSet.CreateRange(used).Add(line));
    }

    // ---------- biểu thức mang tác dụng ----------
    void ExecExpr(ExpressionSyntax e, State st, Frame f)
    {
        switch (e)
        {
            case AssignmentExpressionSyntax a: DoAssign(a, st, f); break;
            case InvocationExpressionSyntax inv: EvalInvocation(inv, st, f, true); break;
            case PostfixUnaryExpressionSyntax pu when pu.Operand is IdentifierNameSyntax id:
                IncDec(id.Identifier.Text, pu.IsKind(SyntaxKind.PostIncrementExpression) ? 1 : -1, st, f, Line(pu)); break;
            case PrefixUnaryExpressionSyntax pr when pr.Operand is IdentifierNameSyntax id2 && (pr.IsKind(SyntaxKind.PreIncrementExpression) || pr.IsKind(SyntaxKind.PreDecrementExpression)):
                IncDec(id2.Identifier.Text, pr.IsKind(SyntaxKind.PreIncrementExpression) ? 1 : -1, st, f, Line(pr)); break;
            case AwaitExpressionSyntax aw: ExecExpr(aw.Expression, st, f); break;
            case ParenthesizedExpressionSyntax p: ExecExpr(p.Expression, st, f); break;
            default: Eval(e, st, f); break;
        }
    }
    void IncDec(string n, long d, State st, Frame f, int line)
    {
        var cur = Lookup(n, st, f, null);
        if (st.Env.ContainsKey(n)) Assign(st, n, Ar.Add(cur, new VConst(d)), line);
        else Assign(st, "this." + n, Ar.Add(cur, new VConst(d)), line);
    }

    public static readonly HashSet<string> FormatProps = new(StringComparer.Ordinal) { "CellType", "BackColor", "ForeColor", "HorizontalAlignment", "VerticalAlignment", "Formatter", "Font" };

    V DoAssign(AssignmentExpressionSyntax a, State st, Frame f)
    {
        int ln = Line(a);
        V rhs = Eval(a.Right, st, f);
        if (!a.IsKind(SyntaxKind.SimpleAssignmentExpression))
        {
            var cur = Eval(a.Left, st, f);
            string op = a.Kind() switch { SyntaxKind.AddAssignmentExpression => "+", SyntaxKind.SubtractAssignmentExpression => "-", SyntaxKind.MultiplyAssignmentExpression => "*", SyntaxKind.DivideAssignmentExpression => "/", _ => "?" };
            rhs = op == "+" && (Ar.IsStringy(cur) || Ar.IsStringy(rhs)) ? Ar.Concat(cur, rhs) : op == "?" ? new VOpaque("compound") : Ar.Lift(cur, rhs, (x, y) => Ar.Arith(op, x, y));
        }
        switch (a.Left)
        {
            case IdentifierNameSyntax id:
                {
                    string n = id.Identifier.Text;
                    if (!st.Env.ContainsKey(n) && f.Cls != null && f.Cls.Members.Contains(n)) n = "this." + n;
                    Assign(st, n, rhs, ln); return rhs;
                }
            case MemberAccessExpressionSyntax ma when ma.Expression is ThisExpressionSyntax:
                Assign(st, "this." + ma.Name.Identifier.Text, rhs, ln); return rhs;
            case MemberAccessExpressionSyntax ma when ma.Expression is IdentifierNameSyntax oid && st.Env.TryGetValue(oid.Identifier.Text, out var oe) && oe.Val is VSym osy && osy.Text.StartsWith("celltype(", StringComparison.Ordinal):
                Assign(st, oid.Identifier.Text, new VSym(osy.Text.Substring(0, osy.Text.Length - 1) + ", " + ma.Name.Identifier.Text + "=" + R.Expr(rhs, new Flags()) + ")"), ln);
                return rhs;
            case MemberAccessExpressionSyntax ma:
                {
                    string prop = ma.Name.Identifier.Text;
                    if (prop == "Formula" || prop == "Value" || prop == "Text" || prop == "Locked" || FormatProps.Contains(prop))
                    {
                        var tgt = Eval(ma.Expression, st, f);
                        if (pendingThrow) return rhs;
                        bool cellSyntax = ma.Expression is ElementAccessExpressionSyntax eas && eas.Expression is MemberAccessExpressionSyntax em && em.Name.Identifier.Text is "Cells" or "Rows" or "Columns";
                        if (tgt is VCell || prop == "Formula" || cellSyntax)
                        {
                            var ef = NewEffect(st, f, ln);
                            ef.Target = tgt is VCell ? tgt : new VOpaque("target:" + Kind(tgt));
                            if (prop == "Formula") { ef.Op = "setFormula"; ef.Val = rhs; ef.Point = PointId(file, a); }
                            else if (prop == "Value" || prop == "Text") { ef.Op = "setValue"; ef.Val = rhs; ef.Params["prop"] = prop.ToLowerInvariant(); }
                            else if (prop == "Locked")
                            {
                                if (rhs is VConst k && k.Val is bool b) ef.Op = b ? "lock" : "unlock";
                                else { ef.Op = "lock"; ef.Params["locked"] = R.Expr(rhs, new Flags()); ef.Val = rhs; }
                            }
                            else { ef.Op = "changeFormat"; ef.Params["prop"] = prop; ef.Params["value"] = R.Expr(rhs, new Flags()); }
                            Effects.Add(ef);
                        }
                    }
                    return rhs;
                }
            case ElementAccessExpressionSyntax ea when ea.Expression is IdentifierNameSyntax aid && st.Env.TryGetValue(aid.Identifier.Text, out var ent) && ent.Val is VArr arr && ea.ArgumentList.Arguments.Count == 1:
                {
                    var idx = Eval(ea.ArgumentList.Arguments[0].Expression, st, f);
                    if (idx is VConst ik && ik.Val is long ii && ii >= 0 && ii < arr.Items.Count)
                    {
                        var items = new List<V>(arr.Items); items[(int)ii] = rhs;
                        Assign(st, aid.Identifier.Text, new VArr(items), ln);
                    }
                    else Assign(st, aid.Identifier.Text, VLin.Atom(aid.Identifier.Text), ln);
                    return rhs;
                }
        }
        return rhs;
    }

    static string Kind(V v) => v switch { VSym s => s.Text.Length > 40 ? "runtime" : s.Text, VOpaque o => o.What, _ => v.GetType().Name.Substring(1).ToLowerInvariant() };

    Effect NewEffect(State st, Frame f, int ln) => new()
    {
        Path = st.Path, Loops = st.Loops, Line = ln, Via = f.Depth > 0 ? f.Via : null,
        Lines = ImmutableSortedSet.CreateRange(used).Add(ln).Union(st.Path.Skip(f.Depth > 0 ? 0 : 0).Select(p => p.Line))
    };

    // ---------- đánh giá ----------
    V Lookup(string n, State st, Frame f, SyntaxNode at)
    {
        if (st.Env.TryGetValue(n, out var e)) { used.UnionWith(e.Lines); return e.Val; }
        if (st.Env.TryGetValue("this." + n, out var e2)) { used.UnionWith(e2.Lines); return e2.Val; }
        if (f.Cls != null && f.Cls.Members.Contains(n)) return Field(n, f.Cls);
        return new VSym(n);
    }

    public V Field(string n, ClassCtx cls)
    {
        if (cls.FieldType.TryGetValue(n, out var ft)) { if (ft == "FpSpread") return VSpread.I; if (ft == "SheetView") return new VSheet(new VSym("field(" + n + ")")); }
        if (cls.Mutable.Contains(n)) return new VSym("field(" + n + ")");
        if (cls.Cache.TryGetValue(n, out var v)) return v;
        if (!cls.Evaluating.Add(n)) return new VSym("field(" + n + ")");
        var saved = used; used = new();
        var st = new State { Env = new(StringComparer.Ordinal), Path = ImmutableList<PathItem>.Empty, Loops = ImmutableList<LoopInfo>.Empty };
        var f = new Frame { Cls = cls, Method = "<field>" };
        if (cls.FieldInit.TryGetValue(n, out var init)) v = init == null ? VConst.Null : Eval(init, st, f);
        else if (cls.PropBody.TryGetValue(n, out var body))
        {
            if (body is ExpressionSyntax ex) v = Eval(ex, st, f);
            else if (body is BlockSyntax b) { Exec(b, st, f); v = f.Returns.Count == 1 ? f.Returns[0].Val : new VSym("field(" + n + ")"); }
            else v = new VSym("field(" + n + ")");
        }
        else v = new VSym("field(" + n + ")");
        used = saved;
        cls.Evaluating.Remove(n);
        cls.Cache[n] = v;
        return v;
    }

    public V Eval(ExpressionSyntax e, State st, Frame f)
    {
        if (++steps > MaxSteps) { Truncated = true; return new VOpaque("budget"); }
        switch (e)
        {
            case LiteralExpressionSyntax lit:
                {
                    var val = lit.Token.Value;
                    return val switch
                    {
                        null => VConst.Null,
                        string s => new VConst(s),
                        char ch => new VConst(ch.ToString()),
                        bool b => new VConst(b),
                        int i => new VConst((long)i), long l => new VConst(l), uint ui => new VConst((long)ui),
                        double d => new VConst(d), float fl => new VConst((double)fl), decimal dm => new VConst((double)dm),
                        _ => new VConst(val.ToString())
                    };
                }
            case InterpolatedStringExpressionSyntax istr:
                return VCat.Make(istr.Contents.Select(c => c is InterpolatedStringTextSyntax t ? new VConst(t.TextToken.ValueText) : Ar.ToStr(Eval(((InterpolationSyntax)c).Expression, st, f))));
            case IdentifierNameSyntax id: return Lookup(id.Identifier.Text, st, f, id);
            case ParenthesizedExpressionSyntax p: return Eval(p.Expression, st, f);
            case CastExpressionSyntax ce:
                {
                    var v = Eval(ce.Expression, st, f);
                    var t = ce.Type.ToString();
                    if ((t == "int" || t == "long") && v is VConst k && k.Val is double d) return new VConst((long)d);
                    if ((t == "double" || t == "decimal") && v is VConst k2 && k2.Val is long l) return new VConst((double)l);
                    return v;
                }
            case CheckedExpressionSyntax ck: return Eval(ck.Expression, st, f);
            case ThisExpressionSyntax: return new VSym("this");
            case PredefinedTypeSyntax pt: return new VSym(pt.Keyword.ValueText);
            case MemberAccessExpressionSyntax ma: return EvalMember(ma, st, f);
            case ElementAccessExpressionSyntax ea: return EvalElement(ea, st, f);
            case InvocationExpressionSyntax inv: return EvalInvocation(inv, st, f, false);
            case AssignmentExpressionSyntax asg: return DoAssign(asg, st, f);
            case BinaryExpressionSyntax be: return EvalBinary(be, st, f);
            case PrefixUnaryExpressionSyntax pu:
                {
                    var v = Eval(pu.Operand, st, f);
                    if (pu.IsKind(SyntaxKind.LogicalNotExpression)) return new VCond(C.Not(C.Of(v)));
                    if (pu.IsKind(SyntaxKind.UnaryMinusExpression)) return Ar.Lift(new VConst(0L), v, (x, y) => Ar.Arith("-", x, y)) is var r && v is VConst kd && kd.Val is double dd ? new VConst(-dd) : r;
                    if (pu.IsKind(SyntaxKind.PreIncrementExpression) || pu.IsKind(SyntaxKind.PreDecrementExpression)) { ExecExpr(pu, st, f); return Eval(pu.Operand, st, f); }
                    return v;
                }
            case PostfixUnaryExpressionSyntax po:
                {
                    var v = Eval(po.Operand, st, f);
                    if (po.IsKind(SyntaxKind.SuppressNullableWarningExpression)) return v;
                    ExecExpr(po, st, f); return v;
                }
            case ConditionalExpressionSyntax ce2:
                {
                    var c = C.Of(Eval(ce2.Condition, st, f));
                    if (c is CConst k) return k.B ? Eval(ce2.WhenTrue, st, f) : Eval(ce2.WhenFalse, st, f);
                    return Ar.Cases(new List<(Cond, V)> { (c, Eval(ce2.WhenTrue, st, f)), (null, Eval(ce2.WhenFalse, st, f)) });
                }
            case ObjectCreationExpressionSyntax oc:
                {
                    if (oc.Initializer != null && oc.Initializer.IsKind(SyntaxKind.CollectionInitializerExpression))
                        return new VArr(oc.Initializer.Expressions.Select(x => Eval(x, st, f)).ToList());
                    if (oc.ArgumentList != null) foreach (var a in oc.ArgumentList.Arguments) Eval(a.Expression, st, f);
                    var tn = oc.Type.ToString(); int lt = tn.IndexOf('<'); if (lt > 0) tn = tn.Substring(0, lt);
                    if ((tn == "List" || tn == "HashSet") && oc.Initializer == null && (oc.ArgumentList == null || oc.ArgumentList.Arguments.Count == 0)) return new VArr(new List<V>());
                    tn = tn.Split('.').Last();
                    if (tn.EndsWith("CellType", StringComparison.Ordinal))
                    {
                        var props = oc.Initializer?.Expressions.OfType<AssignmentExpressionSyntax>().Select(a => a.Left.ToString().Trim() + "=" + R.Expr(Eval(a.Right, st, f), new Flags())) ?? Enumerable.Empty<string>();
                        return new VSym("celltype(" + string.Join(", ", new[] { tn.Substring(0, tn.Length - 8) }.Concat(props)) + ")");
                    }
                    return new VSym("obj(" + tn + ")");
                }
            case ImplicitObjectCreationExpressionSyntax: return new VSym("obj()");
            case ArrayCreationExpressionSyntax ac:
                {
                    if (ac.Initializer != null) return new VArr(ac.Initializer.Expressions.Select(x => Eval(x, st, f)).ToList());
                    var rank = ac.Type.RankSpecifiers.FirstOrDefault()?.Sizes.FirstOrDefault();
                    if (rank != null && rank is not OmittedArraySizeExpressionSyntax && Eval(rank, st, f) is VConst k && k.Val is long n && n >= 0 && n <= 64)
                        return new VArr(Enumerable.Range(0, (int)n).Select(_ => (V)new VConst(0L)).ToList());
                    return new VSym("array");
                }
            case ImplicitArrayCreationExpressionSyntax ia: return new VArr(ia.Initializer.Expressions.Select(x => Eval(x, st, f)).ToList());
            case InitializerExpressionSyntax ie: return new VArr(ie.Expressions.Select(x => Eval(x, st, f)).ToList());
            case LambdaExpressionSyntax le: return new VLambda(le);
            case DefaultExpressionSyntax: return VConst.Null;
            case TypeOfExpressionSyntax to: return new VSym("type(" + to.Type + ")");
            case AwaitExpressionSyntax aw: return Eval(aw.Expression, st, f);
            case GenericNameSyntax gn: return new VSym(gn.Identifier.Text);
            case QualifiedNameSyntax qn: return new VSym(qn.Right.Identifier.Text);
            case IsPatternExpressionSyntax ip: Eval(ip.Expression, st, f); return new VCond(new CAtom("pattern", new VOpaque("pattern")));
            default: return new VOpaque("expr:" + e.Kind());
        }
    }

    static readonly Dictionary<SyntaxKind, string> CmpOps = new()
    {
        [SyntaxKind.EqualsExpression] = "==", [SyntaxKind.NotEqualsExpression] = "!=", [SyntaxKind.LessThanExpression] = "<",
        [SyntaxKind.LessThanOrEqualExpression] = "<=", [SyntaxKind.GreaterThanExpression] = ">", [SyntaxKind.GreaterThanOrEqualExpression] = ">=",
    };

    V EvalBinary(BinaryExpressionSyntax be, State st, Frame f)
    {
        var k = be.Kind();
        if (k == SyntaxKind.LogicalAndExpression || k == SyntaxKind.LogicalOrExpression)
        {
            var a = C.Of(Eval(be.Left, st, f));
            if (k == SyntaxKind.LogicalAndExpression && a is CConst ka && !ka.B) return new VCond(CConst.F);
            if (k == SyntaxKind.LogicalOrExpression && a is CConst kb && kb.B) return new VCond(CConst.T);
            var b = C.Of(Eval(be.Right, st, f));
            return new VCond(k == SyntaxKind.LogicalAndExpression ? C.And(a, b) : C.Or(a, b));
        }
        if (k == SyntaxKind.CoalesceExpression)
        {
            var a = Eval(be.Left, st, f);
            if (a is VConst kn && kn.Val == null) return Eval(be.Right, st, f);
            return a;
        }
        if (k == SyntaxKind.AsExpression || k == SyntaxKind.IsExpression) { var a = Eval(be.Left, st, f); return k == SyntaxKind.AsExpression ? a : new VCond(new CAtom("isType", a)); }
        var l = Eval(be.Left, st, f); var r = Eval(be.Right, st, f);
        if (CmpOps.TryGetValue(k, out var op)) return new VCond(C.Cmp(op, l, r));
        if ((k == SyntaxKind.BitwiseAndExpression || k == SyntaxKind.BitwiseOrExpression) && !(l is VConst ln && ln.Val is long && r is VConst rn && rn.Val is long))
            return new VCond(k == SyntaxKind.BitwiseAndExpression ? C.And(C.Of(l), C.Of(r)) : C.Or(C.Of(l), C.Of(r)));
        switch (k)
        {
            case SyntaxKind.AddExpression:
                if (Ar.IsStringy(l) || Ar.IsStringy(r)) return Ar.Concat(Ar.ToStr(l), Ar.ToStr(r));
                return Ar.Lift(l, r, (x, y) => Ar.Arith("+", x, y));
            case SyntaxKind.SubtractExpression: return Ar.Lift(l, r, (x, y) => Ar.Arith("-", x, y));
            case SyntaxKind.MultiplyExpression: return Ar.Lift(l, r, (x, y) => Ar.Arith("*", x, y));
            case SyntaxKind.DivideExpression: return Ar.Lift(l, r, (x, y) => Ar.Arith("/", x, y));
            case SyntaxKind.ModuloExpression: return Ar.Lift(l, r, (x, y) => Ar.Arith("%", x, y));
        }
        return new VOpaque("binary:" + k);
    }

    V EvalMember(MemberAccessExpressionSyntax ma, State st, Frame f)
    {
        string n = ma.Name.Identifier.Text;
        if (ma.Expression is ThisExpressionSyntax) return Lookup(n, st, f, ma);
        if (ma.Expression is PredefinedTypeSyntax pt)
        {
            if (n == "Empty") return new VConst("");
            return new VSym(pt.Keyword.ValueText + "." + n);
        }
        if (ma.Expression is IdentifierNameSyntax tid && !st.Env.ContainsKey(tid.Identifier.Text) && !st.Env.ContainsKey("this." + tid.Identifier.Text) && (f.Cls == null || !f.Cls.Members.Contains(tid.Identifier.Text)))
        {
            // tham chiếu tĩnh kiểu khác: String.Empty, GeneralFunctions.X ...
            if (n == "Empty" && (tid.Identifier.Text == "String" || tid.Identifier.Text == "string")) return new VConst("");
            return new VSym(tid.Identifier.Text + "." + n);
        }
        var b = Eval(ma.Expression, st, f);
        return Member(b, n);
    }

    public static V Member(V b, string n)
    {
        switch (b)
        {
            case VSpread:
                if (n == "Sheets") return VSheets.I;
                if (n == "ActiveSheet") return new VSheet(new VSym("activeSheet"));
                if (n == "ActiveSheetIndex") return VLin.Atom("activeSheet");
                return new VSym("spread." + n);
            case VSheets: return new VSym("sheets." + n);
            case VSheet s:
                if (n == "Cells") return new VCells(s, "cells");
                if (n == "Rows") return new VCells(s, "rows");
                if (n == "Columns") return new VCells(s, "cols");
                if (n == "FpSpread" || n == "Parent") return VSpread.I;
                if (n == "SheetName") return new VSym("sheetName[" + R.Sheet(s, new Flags()) + "]");
                return new VSym("sheet[" + R.Sheet(s, new Flags()) + "]." + n);
            case VCell c:
                {
                    var fl = new Flags(); string rf = R.CellRef(c, fl);
                    string shn = R.Sheet(c.S, new Flags());
                    string pre = c.S.Index is VConst sk && sk.Val is long si ? ("S" + si + "!") : shn == "activeSheet" ? "" : "S{" + shn + "}!";
                    return new VSym("cell(" + pre + rf + ")." + n.ToLowerInvariant());
                }
            case VCellInfo ci:
                if (n == "CellTemplate" || n == "Location") return ci;
                if (n == "RowIndex") return VLin.Atom("row(" + ci.Path + ")");
                if (n == "ColIndex") return VLin.Atom("col(" + ci.Path + ")");
                return new VSym("cellInfo(" + ci.Path + ")." + n);
            case VConst k when k.Val is string s2 && n == "Length": return new VConst((long)s2.Length);
            case VArr a when n == "Length" || n == "Count": return new VConst((long)a.Items.Count);
            case VOpaque: return b;
            case VCases cs: return Ar.Cases(cs.Alts.Select(x => (x.When, Member(x.Val, n))).ToList());
            case VSym y:
                if (n == "Location" || n == "CellTemplate" || n == "Value" && y.Text.EndsWith(".value")) return y;
                if (n == "Cells") return new VSym(y.Text + ".cells");
                return new VSym(y.Text + "." + n);
            default:
                {
                    var fl = new Flags(); var t = R.Expr(b, fl);
                    if (fl.Opaque) return new VOpaque("member:" + n);
                    return new VSym(t + "." + n);
                }
        }
    }

    V EvalElement(ElementAccessExpressionSyntax ea, State st, Frame f)
    {
        var b = Eval(ea.Expression, st, f);
        var idx = ea.ArgumentList.Arguments.Select(a => Eval(a.Expression, st, f)).ToList();
        return Element(b, idx);
    }
    V Element(V b, List<V> idx)
    {
        switch (b)
        {
            case VArr a when idx.Count == 1:
                if (idx[0] is VConst k && k.Val is long i) { if (i >= 0 && i < a.Items.Count) return a.Items[(int)i]; pendingThrow = true; return new VOpaque("index-out-of-range"); }
                return new VOpaque("array-index");
            case VSheets when idx.Count == 1: return new VSheet(idx[0]);
            case VCells c:
                if (c.Kind == "cells" && idx.Count == 2) return new VCell(c.S, "cell", idx[0], idx[1]);
                if (c.Kind == "cells" && idx.Count == 4) return new VCell(c.S, "range", idx[0], idx[1], Ar.Add(Ar.Add(idx[0], idx[2]), new VConst(-1L)), Ar.Add(Ar.Add(idx[1], idx[3]), new VConst(-1L)));
                if (c.Kind == "rows" && idx.Count == 1) return new VCell(c.S, "row", idx[0], null);
                if (c.Kind == "cols" && idx.Count == 1) return new VCell(c.S, "col", null, idx[0]);
                return new VOpaque("cells-index");
            case VCases cs: return Ar.Cases(cs.Alts.Select(x => (x.When, Element(x.Val, idx))).ToList());
            case VOpaque: return b;
            case VConst ks when ks.Val is string s && idx.Count == 1 && idx[0] is VConst ki && ki.Val is long ii: return ii >= 0 && ii < s.Length ? new VConst(s[(int)ii].ToString()) : new VOpaque("index-out-of-range");
            default:
                {
                    var fl = new Flags(); var t = R.Expr(b, fl) + "[" + string.Join(",", idx.Select(i => R.Expr(i, fl))) + "]";
                    if (fl.Opaque) return new VOpaque("element");
                    return new VSym(t);
                }
        }
    }

    static string StrConst(V v) => v is VConst k && k.Val is string s ? s : null;

    static V ArrayCellIdOf(V v)
    {
        if (v is VCases cs) return Ar.Cases(cs.Alts.Select(a => (a.When, ArrayCellIdOf(a.Val))).ToList());
        var s = StrConst(v);
        if (s == null) return new VArr(new List<V> { VLin.Atom("colIdx(" + R.Expr(v, new Flags()) + ")"), VLin.Atom("rowIdx(" + R.Expr(v, new Flags()) + ")") });
        var p = s.Split('_');
        if (p.Length < 2 || !long.TryParse(p[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var row)) return new VArr(new List<V> { new VConst(-1L), new VConst(-1L) });
        return new VArr(new List<V> { new VConst(Fmt.LettersToIndex(p[0]) - 1), new VConst(row - 1) });
    }

    V Arg(InvocationExpressionSyntax inv, int i, State st, Frame f) => i < inv.ArgumentList.Arguments.Count ? Eval(inv.ArgumentList.Arguments[i].Expression, st, f) : null;

    V EvalInvocation(InvocationExpressionSyntax inv, State st, Frame f, bool isStmt)
    {
        string name; ExpressionSyntax recv = null;
        switch (inv.Expression)
        {
            case IdentifierNameSyntax id: name = id.Identifier.Text; break;
            case GenericNameSyntax g: name = g.Identifier.Text; break;
            case MemberAccessExpressionSyntax ma: name = ma.Name.Identifier.Text; recv = ma.Expression; break;
            default: name = "?"; break;
        }
        var args = inv.ArgumentList.Arguments;
        foreach (var oa in args)
            if (oa.RefKindKeyword.IsKind(SyntaxKind.OutKeyword) || oa.RefKindKeyword.IsKind(SyntaxKind.RefKeyword))
            {
                string on = oa.Expression is IdentifierNameSyntax oi ? oi.Identifier.Text : oa.Expression is DeclarationExpressionSyntax de && de.Designation is SingleVariableDesignationSyntax sv ? sv.Identifier.Text : null;
                if (on != null) Assign(st, on, new VSym(on), Line(inv));
            }
        string rt = recv?.ToString() ?? "";
        bool staticRecv = recv is IdentifierNameSyntax rid && !st.Env.ContainsKey(rid.Identifier.Text) && (f.Cls == null || !f.Cls.Members.Contains(rid.Identifier.Text)) || recv is PredefinedTypeSyntax;
        int ln = Line(inv);

        // nội tuyến method cùng class
        if ((recv == null || recv is ThisExpressionSyntax) && f.Cls != null && f.Cls.Methods.TryGetValue(name, out var cands))
        {
            var m = cands.FirstOrDefault(c => c.ParameterList.Parameters.Count == args.Count) ?? cands.FirstOrDefault(c => c.ParameterList.Parameters.Count >= args.Count && c.ParameterList.Parameters.Skip(args.Count).All(p => p.Default != null));
            if (m != null && f.Depth < MaxDepth && !f.Stack.Contains(name + "/" + m.SpanStart))
            {
                var bind = new Dictionary<string, V>(StringComparer.Ordinal);
                var ps = m.ParameterList.Parameters;
                for (int i = 0; i < ps.Count; i++)
                {
                    ArgumentSyntax a = null;
                    var named = args.FirstOrDefault(x => x.NameColon?.Name.Identifier.Text == ps[i].Identifier.Text);
                    if (named != null) a = named; else if (i < args.Count && args[i].NameColon == null) a = args[i];
                    bind[ps[i].Identifier.Text] = a != null ? Eval(a.Expression, st, f) : ps[i].Default != null ? Eval(ps[i].Default.Value, st, f) : new VSym(ps[i].Identifier.Text);
                }
                var savedUsed = used;
                var sub = new Frame { Cls = f.Cls, Method = name, Depth = f.Depth + 1, Via = (f.Via != null ? f.Via + " > " : "") + name + "@" + Line(m), Stack = f.Stack.Add(name + "/" + m.SpanStart) };
                var sst = new State { Env = new Dictionary<string, Entry>(StringComparer.Ordinal), Path = st.Path, Loops = st.Loops };
                foreach (var kv in bind) sst.Env[kv.Key] = new Entry(kv.Value, ImmutableSortedSet.CreateRange(savedUsed).Add(ln));
                // trường lớp đã gán trong hook hiện tại được truyền xuống
                foreach (var kv in st.Env) if (kv.Key.StartsWith("this.", StringComparison.Ordinal)) sst.Env[kv.Key] = kv.Value;
                sub.BasePathLen = st.Path.Count;
                if (m.Body != null) Exec(m.Body, sst, sub);
                else if (m.ExpressionBody != null) { used = new(); sub.Returns.Add((sst.Path, Eval(m.ExpressionBody.Expression, sst, sub))); }
                used = savedUsed;
                foreach (var kv in sst.Env) if (kv.Key.StartsWith("this.", StringComparison.Ordinal)) st.Env[kv.Key] = kv.Value;
                if (sub.Returns.Count == 0) return VConst.Null;
                if (sub.Returns.Count == 1) return sub.Returns[0].Val;
                var alts = sub.Returns.Select(r => ((Cond)C.And(r.Path.Skip(st.Path.Count).Select(p => p.C)), r.Val)).ToList();
                var cv = Ar.Cases(alts.Select((a, i) => (i == alts.Count - 1 ? null : a.Item1, a.Item2)).ToList());
                return cv is VOpaque ? new VOpaque("ret:" + name) : cv;
            }
        }

        switch (name)
        {
            case "ArrayCellId" when args.Count == 1: return ArrayCellIdOf(Arg(inv, 0, st, f));
            case "ColumnLettersToIndex" when args.Count == 1:
                {
                    var v = Arg(inv, 0, st, f); var s = StrConst(v);
                    return s != null ? new VConst(Fmt.LettersToIndex(s)) : VLin.Atom("colIndex(" + R.Expr(v, new Flags()) + ")");
                }
            case "ColumnIndexToLetters" when args.Count == 1:
                {
                    var v = Arg(inv, 0, st, f);
                    return LettersOf(v);
                }
            case "ToString" or "ToStringOrDefault" or "Trim" or "ToUpper" or "ToLower" or "TrimEnd" or "TrimStart" when recv != null && !staticRecv:
                {
                    var v = Eval(recv, st, f);
                    if (name == "ToString" && args.Count > 0) { var fv = Arg(inv, 0, st, f); if (v is VConst k && k.IsNum && fv is VConst) return new VSym("format(" + Fmt.Num(k.Val) + ")"); return v; }
                    v = Ar.ToStr(v);
                    if (v is VConst ks && ks.Val is string s) return new VConst(name switch { "Trim" => s.Trim(), "ToUpper" => s.ToUpperInvariant(), "ToLower" => s.ToLowerInvariant(), "TrimEnd" => s.TrimEnd(), "TrimStart" => s.TrimStart(), _ => s });
                    if (name is "ToUpper" or "ToLower") return Ar.ToStr(Fn(name == "ToUpper" ? "upper" : "lower", v));
                    return v;
                }
            case "Format" when staticRecv && (rt == "string" || rt == "String"):
                {
                    var fmt = Arg(inv, 0, st, f);
                    var vals = Enumerable.Range(1, args.Count - 1).Select(i => Ar.ToStr(Arg(inv, i, st, f))).ToList();
                    if (vals.Count == 1 && vals[0] is VArr va) vals = va.Items.Select(Ar.ToStr).ToList();
                    return Ar.Lift(fmt, new VConst(0L), (x, _) => FormatV(x, vals));
                }
            case "Concat" when staticRecv && (rt == "string" || rt == "String"):
                return VCat.Make(Enumerable.Range(0, args.Count).Select(i => Ar.ToStr(Arg(inv, i, st, f))));
            case "Replace" when recv != null && args.Count == 2 && !staticRecv:
                return Ar.Replace(Ar.ToStr(Eval(recv, st, f)), Arg(inv, 0, st, f), Ar.ToStr(Arg(inv, 1, st, f)));
            case "Equals" when recv != null && args.Count >= 1 && !staticRecv:
                return new VCond(C.Cmp("==", Eval(recv, st, f), Arg(inv, 0, st, f)));
            case "Equals" when staticRecv && args.Count >= 2:
                return new VCond(C.Cmp("==", Arg(inv, 0, st, f), Arg(inv, 1, st, f)));
            case "IsNullOrEmpty" or "IsNullOrWhiteSpace" when staticRecv:
                {
                    var v = Arg(inv, 0, st, f);
                    if (v is VConst k) return new VConst(k.Val == null || (k.Val is string s && (name == "IsNullOrEmpty" ? s.Length == 0 : s.Trim().Length == 0)));
                    return new VCond(new CAtom("isEmpty", v));
                }
            case "Contains" or "StartsWith" or "EndsWith" when recv != null && args.Count == 1 && !staticRecv:
                {
                    var a = Eval(recv, st, f); var b = Arg(inv, 0, st, f);
                    if (StrConst(a) is string sa && StrConst(b) is string sb) return new VConst(name == "Contains" ? sa.Contains(sb, StringComparison.Ordinal) : name == "StartsWith" ? sa.StartsWith(sb, StringComparison.Ordinal) : sa.EndsWith(sb, StringComparison.Ordinal));
                    if (a is VArr arr) return new VCond(C.Or(arr.Items.Select(i => C.Cmp("==", i, b))));
                    return new VCond(new CAtom(name.ToLowerInvariant(), a, b));
                }
            case "Parse" or "ToInt32" or "ToInt64" or "ToInt16" or "ConvertToInt" or "ConvertToLong" or "ToDouble" or "ToDecimal" or "ConvertToDecimal" or "ConvertToDouble" when args.Count >= 1:
                {
                    var v = Arg(inv, 0, st, f);
                    bool isInt = !(name.Contains("Double") || name.Contains("Decimal")) && rt != "double" && rt != "decimal";
                    if (v is VConst k)
                    {
                        if (k.Val is string s) { if (isInt && long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var l)) return new VConst(l); if (double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return isInt ? new VConst((long)d) : new VConst(d); }
                        if (k.IsNum) return k;
                    }
                    if (v is VSym y) return VLin.Atom(y.Text);
                    return v;
                }
            case "Find" or "FirstOrDefault" or "First" or "SingleOrDefault" or "Single" when recv != null && args.Count == 1:
                {
                    var coll = Eval(recv, st, f);
                    var lam = Arg(inv, 0, st, f) as VLambda;
                    var p = lam != null ? PathOfLambda(lam.L, st, f) : null;
                    if (p != null) return new VCellInfo(p);
                    Opaque(rt.Length > 0 ? "<collection>." + name : name, inv, "value");
                    return new VOpaque(name);
                }
            case "GetValueCustom" when args.Count >= 3:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var r = Arg(inv, 1, st, f); var c = Arg(inv, 2, st, f);
                    if (sh != null) return Member(new VCell(sh, "cell", Ar.Add(r, new VConst(-1L)), Ar.Add(c, new VConst(-1L))), "value");
                    break;
                }
            case "GetObjectValue" when args.Count == 2:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var id = ArrayCellIdOf(Arg(inv, 1, st, f));
                    if (sh != null && id is VArr ar && ar.Items.Count == 2) return Member(new VCell(sh, "cell", ar.Items[1], ar.Items[0]), "value");
                    break;
                }
            case "ChangeValueCell" when args.Count >= 6:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var r = Arg(inv, 1, st, f); var c = Arg(inv, 2, st, f);
                    var lk = Arg(inv, 3, st, f); var fm = Arg(inv, 4, st, f); var val = Arg(inv, 5, st, f);
                    if (sh == null) break;
                    var cell = new VCell(sh, "cell", r, c);
                    var e1 = NewEffect(st, f, ln); e1.Target = cell; e1.Params["via"] = "ChangeValueCell";
                    if (fm is VConst fk && (fk.Val == null || (fk.Val is string fs && fs.Length == 0))) { e1.Op = "setValue"; e1.Val = val; e1.Params["resetFormula"] = true; }
                    else if (StrConst(fm) != null || fm is VCat) { e1.Op = "setFormula"; e1.Val = fm; }
                    else { e1.Op = "setFormula"; e1.Val = fm; e1.Params["elseValue"] = R.Expr(val, new Flags()); e1.Params["note"] = "nếu công thức rỗng thì đặt giá trị"; }
                    Effects.Add(e1);
                    AddLock(st, f, ln, cell, lk, "ChangeValueCell");
                    return VConst.Null;
                }
            case "LockUnLockCol" when args.Count >= 4:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var r = Arg(inv, 1, st, f); var c = Arg(inv, 2, st, f); var lk = Arg(inv, 3, st, f); var del = args.Count > 4 ? Arg(inv, 4, st, f) : new VConst(false);
                    if (sh == null) break;
                    var cell = new VCell(sh, "cell", Ar.Add(r, new VConst(-1L)), Ar.Add(c, new VConst(-1L)));
                    AddLock(st, f, ln, cell, lk, "LockUnLockCol", del);
                    return VConst.Null;
                }
            case "LockUnLockCols" when args.Count >= 4:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var r = Arg(inv, 1, st, f); var cs = Arg(inv, 2, st, f); var lk = Arg(inv, 3, st, f); var del = args.Count > 4 ? Arg(inv, 4, st, f) : new VConst(false);
                    if (sh == null || cs is not VArr arr) break;
                    foreach (var c in arr.Items) AddLock(st, f, ln, new VCell(sh, "cell", Ar.Add(r, new VConst(-1L)), Ar.Add(c, new VConst(-1L))), lk, "LockUnLockCols", del);
                    return VConst.Null;
                }
            case "changFormatCell" when args.Count == 3:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var fmt = Arg(inv, 1, st, f); var cur = Arg(inv, 2, st, f);
                    var s = StrConst(fmt);
                    if (sh == null || s == null) break;
                    foreach (var part in s.Split("###"))
                    {
                        if (ArrayCellIdOf(new VConst(part)) is not VArr ar) continue;
                        var e1 = NewEffect(st, f, ln); e1.Op = "changeFormat"; e1.Target = new VCell(sh, "cell", ar.Items[1], ar.Items[0]);
                        e1.Params["via"] = "changFormatCell"; e1.Params["decimals"] = "2 nếu " + R.Expr(cur, new Flags()) + " == \"USD\", ngược lại 0";
                        e1.Val = cur;
                        Effects.Add(e1);
                    }
                    return VConst.Null;
                }
            case "ResetValue" when recv != null && args.Count == 0:
                {
                    var t = Eval(recv, st, f);
                    if (t is VCell) { var e1 = NewEffect(st, f, ln); e1.Op = "setValue"; e1.Target = t; e1.Val = VConst.Null; e1.Params["reset"] = true; Effects.Add(e1); return VConst.Null; }
                    break;
                }
            case "ResetFormula" when recv != null && args.Count == 0:
                {
                    var t = Eval(recv, st, f);
                    if (t is VCell) { var e1 = NewEffect(st, f, ln); e1.Op = "setFormula"; e1.Target = t; e1.Val = new VConst(""); e1.Params["reset"] = true; Effects.Add(e1); return VConst.Null; }
                    break;
                }
            case "Add" when recv is IdentifierNameSyntax lid && st.Env.TryGetValue(lid.Identifier.Text, out var le) && le.Val is VArr la && args.Count == 1:
                {
                    var items = new List<V>(la.Items) { Arg(inv, 0, st, f) };
                    Assign(st, lid.Identifier.Text, new VArr(items), ln);
                    return VConst.Null;
                }
            case "Split" when recv != null && !staticRecv:
                {
                    var s = StrConst(Eval(recv, st, f)); var sep = args.Count > 0 ? Arg(inv, 0, st, f) : null;
                    string sp = sep is VConst sk && sk.Val is string ss ? ss : sep is VArr sa && sa.Items.Count == 1 && StrConst(sa.Items[0]) is string s1 ? s1 : null;
                    if (s != null && sp != null && sp.Length > 0) return new VArr(s.Split(sp).Select(x => (V)new VConst(x)).ToList());
                    return Fn("split", Eval(recv, st, f), sep ?? VConst.Null);
                }
            case "Substring" when recv != null && !staticRecv:
                {
                    var s = StrConst(Eval(recv, st, f)); var a0 = Arg(inv, 0, st, f); var a1 = args.Count > 1 ? Arg(inv, 1, st, f) : null;
                    if (s != null && a0 is VConst k0 && k0.Val is long i0 && i0 >= 0 && i0 <= s.Length)
                    {
                        if (a1 == null) return new VConst(s.Substring((int)i0));
                        if (a1 is VConst k1 && k1.Val is long i1 && i1 >= 0 && i0 + i1 <= s.Length) return new VConst(s.Substring((int)i0, (int)i1));
                    }
                    return Fn("substring", Eval(recv, st, f), a0, a1 ?? VConst.Null);
                }
            case "ToBoolean" or "ConvertToBool" or "ToBoolOrDefault" when args.Count >= 1 || recv != null:
                {
                    var v = args.Count >= 1 && (staticRecv || recv == null) ? Arg(inv, 0, st, f) : Eval(recv, st, f);
                    if (v is VConst k && k.Val is string bs) return new VConst(bs.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) || bs.Trim() == "1");
                    return new VCond(C.Of(v));
                }
            case "isNullString" when args.Count == 1:
                {
                    var v = Arg(inv, 0, st, f);
                    if (v is VConst k) return new VConst(k.Val == null || (k.Val is string s0 && s0.Trim().Length == 0));
                    return new VCond(new CAtom("isEmpty", v));
                }
            case "IndexOf" or "LastIndexOf" when recv != null && !staticRecv && args.Count >= 1:
                return Fn(name == "IndexOf" ? "indexOf" : "lastIndexOf", Eval(recv, st, f), Arg(inv, 0, st, f));
            case "Round" or "Abs" or "Max" or "Min" or "Floor" or "Ceiling" when rt == "Math":
                {
                    var vs = Enumerable.Range(0, args.Count).Select(i => Arg(inv, i, st, f)).ToArray();
                    return Fn(name.ToLowerInvariant(), vs);
                }
            case "Log" or "Show" or "WriteLine":
                foreach (var a in args) Eval(a.Expression, st, f);
                return VConst.Null;
        }

        // không giải được
        string qn; V rv = null;
        if (recv == null) qn = name;
        else if (staticRecv) qn = rt + "." + name;
        else
        {
            rv = Eval(recv, st, f);
            qn = rv switch { VSheet => "<sheet>.", VCell => "<cell>.", VSpread => "<spread>.", VSheets => "<sheets>.", VCells => "<cells>.", _ => "<obj>." } + name;
        }
        var argv = args.Select(a => a.Expression is LambdaExpressionSyntax ? new VOpaque("lambda") : Eval(a.Expression, st, f)).ToList();
        Opaque(qn, inv, isStmt ? "statement" : "value");
        bool localRecv = rv is VArr || rv is VConst || (rv is VSym rsy && rsy.Text.StartsWith("new ", StringComparison.Ordinal));
        if (isStmt && !localRecv)
        {
            var e1 = NewEffect(st, f, ln); e1.Op = "opaque"; e1.Target = null;
            e1.Params["call"] = qn;
            e1.Params["args"] = argv.Select(a => { var s = R.Expr(a, new Flags()); return s.Length > 160 ? s.Substring(0, 157) + "..." : s; }).ToList();
            Effects.Add(e1);
        }
        return new VOpaque(qn);
    }

    void AddLock(State st, Frame f, int ln, VCell cell, V lk, string via, V del = null)
    {
        var e2 = NewEffect(st, f, ln); e2.Target = cell; e2.Params["via"] = via;
        if (lk is VConst k && k.Val is bool b) e2.Op = b ? "lock" : "unlock";
        else { e2.Op = "lock"; e2.Params["locked"] = R.Expr(lk, new Flags()); e2.Val = lk; }
        if (del != null && !(del is VConst dk && dk.Val is bool db && !db)) e2.Params["clearData"] = R.Expr(del, new Flags());
        Effects.Add(e2);
    }

    static V Fn(string fn, params V[] a)
    {
        var fl = new Flags(); var t = fn + "(" + string.Join(", ", a.Select(x => R.Expr(x, fl))) + ")";
        return fl.Opaque ? new VOpaque(fn) : new VSym(t);
    }

    static V LettersOf(V v)
    {
        if (v is VConst k && k.Val is long n) return n >= 1 ? new VConst(Fmt.Letters(n)) : new VOpaque("letters<1");
        if (v is VCases cs) return Ar.Cases(cs.Alts.Select(a => (a.When, LettersOf(a.Val))).ToList());
        var fl = new Flags(); var t = R.Expr(v, fl);
        return fl.Opaque ? new VOpaque("letters") : VCat.Make(new V[] { new VSym("letters(" + t + ")") });
    }

    static V FormatV(V fmt, List<V> vals)
    {
        if (fmt is VCat fc) return VCat.Make(fc.Parts.Select(p => p is VConst pk && pk.Val is string ps ? FormatV(new VConst(ps), vals) : p));
        var s = StrConst(fmt);
        if (s == null) return new VOpaque("Format");
        var parts = new List<V>();
        int i = 0; var sb = new System.Text.StringBuilder();
        while (i < s.Length)
        {
            char ch = s[i];
            if (ch == '{' && i + 1 < s.Length && s[i + 1] == '{') { sb.Append('{'); i += 2; continue; }
            if (ch == '}' && i + 1 < s.Length && s[i + 1] == '}') { sb.Append('}'); i += 2; continue; }
            if (ch == '{')
            {
                int j = s.IndexOf('}', i);
                if (j < 0) return new VOpaque("Format");
                var spec = s.Substring(i + 1, j - i - 1).Split(':', ',')[0];
                if (!int.TryParse(spec, out var ai) || ai >= vals.Count) return new VOpaque("Format");
                parts.Add(new VConst(sb.ToString())); sb.Clear(); parts.Add(vals[ai]); i = j + 1; continue;
            }
            sb.Append(ch); i++;
        }
        parts.Add(new VConst(sb.ToString()));
        return VCat.Make(parts);
    }

    string PathOfLambda(LambdaExpressionSyntax l, State st, Frame f)
    {
        ExpressionSyntax body = l switch { SimpleLambdaExpressionSyntax s => s.ExpressionBody, ParenthesizedLambdaExpressionSyntax p => p.ExpressionBody, _ => null };
        if (body == null) return null;
        // x.Path.Equals(S) | x.Path == S
        ExpressionSyntax other = null;
        if (body is InvocationExpressionSyntax iv && iv.Expression is MemberAccessExpressionSyntax m && m.Name.Identifier.Text == "Equals" && m.Expression is MemberAccessExpressionSyntax pm && pm.Name.Identifier.Text == "Path" && iv.ArgumentList.Arguments.Count == 1)
            other = iv.ArgumentList.Arguments[0].Expression;
        else if (body is BinaryExpressionSyntax be && be.IsKind(SyntaxKind.EqualsExpression) && be.Left is MemberAccessExpressionSyntax pm2 && pm2.Name.Identifier.Text == "Path")
            other = be.Right;
        if (other == null) return null;
        var v = Eval(other, st, f);
        return StrConst(v);
    }
}
