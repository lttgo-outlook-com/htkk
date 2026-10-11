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
/// Mô tả vòng lặp; Desc được bổ sung sau khi phân tích thân (biến cảm ứng: giá trị đầu, bước)
public sealed class LoopInfo { public string Desc; public readonly int Line; public LoopInfo(string d, int l) { Desc = d; Line = l; } }
/// Ký hiệu lúc chạy có tên (biến cục bộ không truy ngược được thành biểu thức) kèm nguồn
public sealed class SymInfo { public string Name, Origin, Desc; public int Line; }

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

/// Một lần gọi method cùng class: đã nội tuyến hay chưa, lý do, khoảng effect/dropped sinh ra trong lần gọi
public sealed class CallRec { public int Line; public string Name, Via, Reason; public int Depth; public bool Followed; public int EffBefore, EffAfter, DropBefore, DropAfter; }

public sealed class Interp
{
    public readonly List<Effect> Effects = new();
    public readonly List<CallRec> Calls = new();
    // phiên bản đọc ô: mỗi lệnh ghi (trực tiếp, qua method nội tuyến, hoặc lời gọi có thể ghi) tăng phiên bản; hai lần đọc khác phiên bản không coi là cùng giá trị
    readonly Dictionary<string, long> cellVer = new(StringComparer.Ordinal);
    long wildVer, anyVer, fieldVer;
    readonly Dictionary<V, List<Effect>> pendingBtn = new(ReferenceEqualityComparer.Instance);
    static string CellKey(VCell c, out bool isConst)
    {
        isConst = c.Kind == "cell" && c.S.Index is VConst && c.R1 is VConst && c.C1 is VConst;
        return isConst ? R.Expr(c.S.Index, new Flags()) + "!" + R.Expr(c.R1, new Flags()) + "," + R.Expr(c.C1, new Flags()) : null;
    }
    void Wild() { wildVer = anyVer = Mk.Next(); }
    void BumpCell(V t)
    {
        if (t is VCell c) { var k = CellKey(c, out var isConst); if (isConst) { cellVer[k] = anyVer = Mk.Next(); return; } }
        Wild();
    }
    long ReadVer(VCell c)
    {
        var k = CellKey(c, out var isConst);
        return isConst ? Math.Max(cellVer.GetValueOrDefault(k), wildVer) : anyVer;
    }
    void AddEffect(Effect e)
    {
        Effects.Add(e);
        if (e.Target is VCell) BumpCell(e.Target);
        else if (e.Op is "setTabColor" or "setTabProp" or "setButtonText" or "setModelDefault") { }
        else if (e.Op != "opaque" || (e.Params.TryGetValue("touchesSheet", out var ts) && ts is true)) Wild();
        e.Params.Remove("touchesSheet");
    }
    static bool SheetLike(V v) => v is VSheet || v is VSpread || v is VSheets || v is VCells || v is VCell || (v is VCases cs && cs.Alts.Any(a => SheetLike(a.Val))) || (v is VArr va && va.Items.Any(SheetLike));
    static readonly HashSet<string> PureCalls = new(StringComparer.Ordinal) { "ToString", "ToStringOrDefault", "ToIntOrDefault", "ToBoolOrDefault", "Trim", "TrimEnd", "TrimStart", "ToUpper", "ToLower", "Equals", "Contains", "ContainsKey", "TryGetValue", "StartsWith", "EndsWith", "IndexOf", "LastIndexOf", "Substring", "Replace", "Split", "Format", "Concat", "Parse", "TryParse", "ToInt32", "ToInt64", "ToInt16", "ToDouble", "ToDecimal", "ToBoolean", "ConvertToInt", "ConvertToLong", "ConvertToDouble", "ConvertToDecimal", "ConvertToDouble1", "ConvertToDecimal1", "ConvertToBool", "ConvertToString", "IsNullOrEmpty", "IsNullOrWhiteSpace", "isNullString", "ArrayCellId", "ColumnLettersToIndex", "ColumnIndexToLetters", "GetValueCustom", "GetObjectValue", "Log", "Round", "Abs", "Max", "Min", "Floor", "Ceiling", "Add", "Count", "Any", "ToList", "ToArray", "Where", "Select", "FirstOrDefault", "First", "Find", "PadLeft", "PadRight", "GetMessage", "ConvertToDateTime", "ToDateTime", "Exists", "Show", "WriteLine", "Sum", "Length", "GetLength", "AddDays", "AddMonths", "ParseExact", "Compare", "DaysInMonth" };
    /// thân vòng lặp có thể ghi ô (gán thuộc tính, lời gọi không thuần): đọc ô trong vòng lặp không coi là cùng giá trị với đọc trước vòng lặp
    static bool MayWrite(SyntaxNode body) => body.DescendantNodes().Any(n =>
        (n is AssignmentExpressionSyntax a && (a.Left is MemberAccessExpressionSyntax || a.Left is ElementAccessExpressionSyntax ea0 && ea0.Expression is MemberAccessExpressionSyntax)) ||
        (n is InvocationExpressionSyntax iv && !PureCalls.Contains(iv.Expression switch { IdentifierNameSyntax i => i.Identifier.Text, GenericNameSyntax g => g.Identifier.Text, MemberAccessExpressionSyntax m => m.Name.Identifier.Text, _ => "?" })));
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
    public readonly SortedDictionary<string, List<SymInfo>> Syms = new(StringComparer.Ordinal);
    public SymInfo RegSym(string name, string origin, int line, string desc)
    {
        if (!Syms.TryGetValue(name, out var l)) Syms[name] = l = new();
        var e = l.FirstOrDefault(x => x.Origin == origin && x.Line == line);
        if (e != null) return e;
        e = new SymInfo { Name = name, Origin = origin, Line = line, Desc = desc };
        l.Add(e); return e;
    }
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
        foreach (var kv in args)
        {
            st.Env[kv.Key] = new Entry(kv.Value, ImmutableSortedSet<int>.Empty);
            if ((kv.Value is VSym ps && ps.Text == kv.Key) || (kv.Value is VLin pl && pl.IsSingleAtom(out var pa) && pa == kv.Key) || (kv.Value is VSheet psh && psh.Index is VSym pss && pss.Text == kv.Key)) RegSym(kv.Key, "param", Line(m), "tham số của method " + MethodName(m) + ", giá trị do nơi gọi truyền vào");
        }
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
                    V val = v.Initializer != null ? Eval(v.Initializer.Value, st, f) : UninitSym(v.Identifier.Text, Line(v));
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
                if (us.Declaration != null) foreach (var v in us.Declaration.Variables) { used = new(); Assign(st, v.Identifier.Text, v.Initializer != null ? Eval(v.Initializer.Value, st, f) : UninitSym(v.Identifier.Text, Line(v)), Line(v)); }
                Exec(us.Statement, st, f); break;
            case ForStatementSyntax fs:
                {
                    var declInit = new Dictionary<string, V>(StringComparer.Ordinal);
                    if (fs.Declaration != null) foreach (var v in fs.Declaration.Variables) { used = new(); declInit[v.Identifier.Text] = v.Initializer != null ? Eval(v.Initializer.Value, st, f) : UninitSym(v.Identifier.Text, Line(v)); }
                    foreach (var ie in fs.Initializers) { used = new(); ExecExpr(ie, st, f); }
                    var assigned = AssignedIn(fs);
                    if (TryUnroll(fs, st, f)) break;
                    var init0 = Snap(st, assigned);
                    foreach (var kv in declInit) init0[kv.Key] = kv.Value;
                    EnterLoop(st, assigned, Line(fs), fs);
                    if (fs.Declaration != null) foreach (var v in fs.Declaration.Variables) { st.Env[v.Identifier.Text] = new Entry(VLin.Atom(v.Identifier.Text), ImmutableSortedSet.Create(Line(v))); RegSym(v.Identifier.Text, "loop", Line(fs), "biến lặp của vòng lặp dòng " + Line(fs)); }
                    used = new();
                    string cond = fs.Condition != null ? R.Cond(C.Of(Eval(fs.Condition, st, f)), new Flags()) : "true";
                    Loop(st, f, fs.Statement, "for", "while " + cond, Line(fs), assigned, init0, fs.Incrementors, declInit.Keys.ToList());
                    break;
                }
            case ForEachStatementSyntax fe:
                {
                    used = new();
                    var coll = Eval(fe.Expression, st, f);
                    var assigned = AssignedIn(fe);
                    var init0 = Snap(st, assigned);
                    EnterLoop(st, assigned, Line(fe), fe);
                    string ct = R.Expr(coll, new Flags());
                    st.Env[fe.Identifier.Text] = new Entry(new VSym(fe.Identifier.Text), ImmutableSortedSet.Create(Line(fe)));
                    RegSym(fe.Identifier.Text, "foreach", Line(fe), "phần tử đang duyệt của " + ct);
                    Loop(st, f, fe.Statement, "foreach " + fe.Identifier.Text + " in " + ct, null, Line(fe), assigned, init0, null, null);
                    break;
                }
            case WhileStatementSyntax ws:
                {
                    var assigned = AssignedIn(ws);
                    var init0 = Snap(st, assigned);
                    EnterLoop(st, assigned, Line(ws), ws);
                    used = new();
                    Loop(st, f, ws.Statement, "while", R.Cond(C.Of(Eval(ws.Condition, st, f)), new Flags()), Line(ws), assigned, init0, null, null);
                    break;
                }
            case DoStatementSyntax ds:
                {
                    var assigned = AssignedIn(ds);
                    var init0 = Snap(st, assigned);
                    EnterLoop(st, assigned, Line(ds), ds);
                    used = new();
                    Loop(st, f, ds.Statement, "do", "while " + R.Cond(C.Of(Eval(ds.Condition, st, f)), new Flags()), Line(ds), assigned, init0, null, null);
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
        if (vals.Count == 0) return false; // không kết luận vòng lặp 0 lần, để phân tích ký hiệu
        foreach (var v in vals)
        {
            if (st.Dead) break;
            st.Env[name] = new Entry(v, ImmutableSortedSet.Create(Line(fs)));
            Exec(fs.Statement, st, f);
        }
        return true;
    }

    static Dictionary<string, V> Snap(State st, HashSet<string> assigned)
    {
        var d = new Dictionary<string, V>(StringComparer.Ordinal);
        foreach (var n in assigned) if (st.Env.TryGetValue(n, out var e)) d[n] = e.Val;
        return d;
    }

    void EnterLoop(State st, HashSet<string> assigned, int line, SyntaxNode loopNode = null)
    {
        if (loopNode != null && MayWrite(loopNode)) { Wild(); fieldVer = Mk.Next(); }
        foreach (var n in assigned.OrderBy(x => x, StringComparer.Ordinal))
        {
            st.Env[n] = new Entry(VLin.Atom(n), ImmutableSortedSet<int>.Empty);
            RegSym(n, "loop", line, "giá trị đầu mỗi vòng của vòng lặp dòng " + line);
        }
    }

    /// head: "for"/"while"/"do"/"foreach ..."; cond: điều kiện lặp. Mô tả cuối: head + biến cảm ứng (from, step) + điều kiện.
    void Loop(State st, Frame f, StatementSyntax body, string head, string cond, int line, HashSet<string> assigned, Dictionary<string, V> init0, SeparatedSyntaxList<ExpressionSyntax>? incs, List<string> declared)
    {
        var li = new LoopInfo(head + (cond != null ? " " + cond : ""), line);
        var bs = st.Clone(); bs.Loops = bs.Loops.Add(li);
        Exec(body, bs, f);
        if (incs != null) { bs.Dead = false; foreach (var inc in incs.Value) { used = new(); ExecExpr(inc, bs, f); } }
        // biến cảm ứng: giá trị đầu và bước hằng
        var ind = new List<string>();
        foreach (var n in assigned.OrderBy(x => (declared != null && declared.Contains(x)) ? 0 : 1).ThenBy(x => x, StringComparer.Ordinal))
        {
            V init = init0.GetValueOrDefault(n);
            string it = init != null ? R.Expr(init, new Flags()) : "?";
            if (!bs.Env.TryGetValue(n, out var be)) continue;
            var step = StepOf(n, be.Val);
            string bodyTxt = R.Expr(be.Val, new Flags());
            if (step is VConst sc && sc.Val is long sl && sl != 0) ind.Add(n + " from " + it + " step " + (sl > 0 ? "+" : "") + sl);
            foreach (var si in Syms.GetValueOrDefault(n) ?? new())
                if (si.Origin == "loop" && si.Line == line)
                    si.Desc = "biến mang qua vòng lặp dòng " + line + ": giá trị trước vòng " + it + ", cuối mỗi vòng " + n + " := " + bodyTxt;
        }
        // mô tả có số dòng (head@dòng) để ký hiệu biến lặp gắn đúng vòng lặp
        string hd = head.Contains(' ') ? head.Substring(0, head.IndexOf(' ')) + "@" + line + head.Substring(head.IndexOf(' ')) : head + "@" + line;
        string desc = head == "for"
            ? hd + (ind.Count > 0 ? " " + string.Join(", ", ind) : "") + (cond != null ? " " + cond : "")
            : hd + (cond != null ? " " + cond : "") + (ind.Count > 0 ? " [" + string.Join(", ", ind) + "]" : "");
        li.Desc = desc;
        var mutated = MutatedIn(body);
        foreach (var n in assigned)
        {
            if (declared != null && declared.Contains(n)) continue;
            V init = init0.GetValueOrDefault(n), fin = null;
            // tập hợp thêm/bớt phần tử trong vòng lặp ký hiệu: nội dung sau vòng lặp là giá trị lúc chạy (không giữ giá trị trước vòng)
            if (mutated.Contains(n)) { RegSym(n, "loop-collection", line, "tập hợp được thêm/bớt phần tử trong vòng lặp dòng " + line + ", nội dung lúc chạy"); st.Env[n] = new Entry(new VSym(n + Mk.Tag(Mk.Next())), ImmutableSortedSet.Create(line)); continue; }
            if (init != null && bs.Env.TryGetValue(n, out var be)) fin = Accum(n, init, be.Val, desc);
            if (fin == null) RegSym("after_loop", "after_loop", line, "giá trị sau vòng lặp dòng " + line + " khi không xác định được giá trị trước vòng");
            st.Env[n] = new Entry(fin ?? new VSym("after_loop(" + n + ", " + desc + ")"), ImmutableSortedSet.Create(line));
        }
    }

    static bool Mentions(V v, string n) { var fl = new Flags(); R.Expr(v, fl); return fl.Notes.Contains(n); }
    static V StepOf(string n, V v)
    {
        if (v.Key == VLin.Atom(n).Key) return new VConst(0L);
        if (v is VLin l && l.T.TryGetValue(n, out var c) && c == 1)
        {
            var t = new SortedDictionary<string, long>(l.T, StringComparer.Ordinal); t.Remove(n);
            var r = VLin.Make(t, l.C);
            return Mentions(r, n) ? null : r;
        }
        if (v is VSym y && y.Text.StartsWith("(" + n + "+", StringComparison.Ordinal) && y.Text.EndsWith(")"))
        {
            var rest = y.Text.Substring(n.Length + 2, y.Text.Length - n.Length - 3);
            return rest.Contains(n) ? null : new VSym(rest, y.HasOpaque);
        }
        return null;
    }
    static bool IsZero(V v) => v is VConst k && k.IsNum && Convert.ToDouble(k.Val, CultureInfo.InvariantCulture) == 0;
    static bool HasOpq(V v) { var q = new Flags(); R.Expr(v, q); return q.Opaque; }
    // biến cục bộ bị gán trong vòng lặp ký hiệu: biểu diễn tích lũy trung tính, ghi đủ vòng lặp, giá trị đầu, bộ lọc, bước
    static V Accum(string n, V init, V body, string desc)
    {
        if (body.Key == VLin.Atom(n).Key) return init;
        string from = R.Expr(init, new Flags());
        bool opq = HasOpq(init);
        V step = StepOf(n, body);
        string where = null;
        if (step == null && body is VCases cs)
        {
            var alts = new List<(Cond W, V S)>();
            foreach (var (w, v) in cs.Alts) { var s0 = StepOf(n, v); if (s0 == null) { alts = null; break; } alts.Add((w, s0)); }
            if (alts != null)
            {
                // tách bộ lọc: nhánh bước 0 là "không cộng"; còn đúng 1 nhánh khác 0 thì ghi where + step
                var nz = alts.Where(a => !IsZero(a.S)).ToList();
                if (nz.Count == 1 && alts.Count > 1)
                {
                    var prior = new List<Cond>(); Cond wsel = null;
                    foreach (var a in alts)
                    {
                        var cw = a.W ?? C.Not(C.Or(prior));
                        if (ReferenceEquals(a.S, nz[0].S)) { wsel = C.And(C.Not(C.Or(prior)), cw); break; }
                        if (a.W != null) prior.Add(a.W);
                    }
                    var wf = new Flags(); where = R.Cond(wsel, wf); opq |= wf.Opaque;
                    step = nz[0].S;
                }
                else step = Ar.Cases(alts.Select(a => (a.W, a.S)).ToList());
            }
        }
        if (step != null)
        {
            var fl = new Flags(); var txt = "sum_over(" + desc + ", from " + from + (where != null ? ", where " + where : "") + ", step " + R.Expr(step, fl) + ")";
            return new VSym(txt, fl.Opaque || opq);
        }
        var fl2 = new Flags(); var bt = R.Expr(body, fl2);
        if (!Mentions(body, n)) return new VSym("last_in_loop(" + desc + ", from " + from + ", last " + bt + ")", fl2.Opaque || opq);
        return new VSym("accum(" + n + " := " + bt + ", from " + from + ", " + desc + ")", fl2.Opaque || opq);
    }

    static HashSet<string> MutatedIn(SyntaxNode n)
    {
        var h = new HashSet<string>(StringComparer.Ordinal);
        foreach (var a in n.DescendantNodes())
            if (a is InvocationExpressionSyntax iv && iv.Expression is MemberAccessExpressionSyntax im && im.Expression is IdentifierNameSyntax id4 && im.Name.Identifier.Text is "Add" or "Insert" or "Remove" or "RemoveAt" or "Clear" or "AddRange" or "RemoveAll" or "Reverse" or "Sort") h.Add(id4.Identifier.Text);
        return h;
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
            if (v is VOpaque) { v = VLin.Atom("opaque(branches)" + Mk.Tag(Mk.Next(), k.StartsWith("this.", StringComparison.Ordinal) ? k.Substring(5) : k)); RegSym(k.StartsWith("this.", StringComparison.Ordinal) ? k.Substring(5) : k, "merge", line, "giá trị phụ thuộc nhiều hơn " + Ar.MaxAlts + " nhánh rẽ tại dòng " + line); }
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
                {
                    var ofl = new Flags(); var nv = new VSym(osy.Text.Substring(0, osy.Text.Length - 1) + ", " + ma.Name.Identifier.Text + "=" + R.Expr(rhs, ofl) + ")", osy.HasOpaque || ofl.Opaque);
                    var pl = pendingBtn.TryGetValue(osy, out var pl0) ? new List<Effect>(pl0) : new List<Effect>();
                    if (ma.Name.Identifier.Text == "Text" && osy.Text.StartsWith("celltype(Button", StringComparison.Ordinal))
                    {
                        // chữ trên nút: effect tại dòng gán Text; ô đích điền khi kiểu ô được gán vào ô
                        var eb = NewEffect(st, f, ln); eb.Op = "setButtonText"; eb.Target = null; eb.Val = rhs;
                        eb.Params["text"] = R.Expr(rhs, new Flags()); eb.Params["note"] = "kiểu ô nút chưa gán vào ô nào trong phạm vi phân tích";
                        AddEffect(eb); pl.Add(eb);
                    }
                    if (pl.Count > 0) pendingBtn[nv] = pl;
                    Assign(st, oid.Identifier.Text, nv, ln);
                    return rhs;
                }
            case MemberAccessExpressionSyntax ma:
                {
                    string prop = ma.Name.Identifier.Text;
                    if (ma.Expression is ElementAccessExpressionSyntax tea && tea.Expression is MemberAccessExpressionSyntax tm && tm.Name.Identifier.Text == "TabStrip")
                    {
                        var ti = tea.ArgumentList.Arguments.Count == 1 ? Eval(tea.ArgumentList.Arguments[0].Expression, st, f) : new VOpaque("tab-index");
                        var et = NewEffect(st, f, ln); et.Op = prop == "BackColor" ? "setTabColor" : "setTabProp"; et.Target = null; et.Val = rhs;
                        et.Params["tab"] = ti is VConst tk ? tk.Val : R.Expr(ti, new Flags());
                        if (prop != "BackColor") et.Params["prop"] = prop;
                        et.Params[prop == "BackColor" ? "color" : "value"] = R.Expr(rhs, new Flags());
                        et.Params["note"] = "màu/thuộc tính thẻ sheet (tab) theo chỉ số sheet";
                        AddEffect(et);
                        return rhs;
                    }
                    if (prop == "Text" && ma.Expression is IdentifierNameSyntax bid && bid.Identifier.Text.Contains("ButtonCellType", StringComparison.OrdinalIgnoreCase))
                    {
                        var eb = NewEffect(st, f, ln); eb.Op = "setButtonText"; eb.Target = null; eb.Val = rhs;
                        eb.Params["text"] = R.Expr(rhs, new Flags()); eb.Params["note"] = "kiểu ô nút lấy lúc chạy, ô đích chưa xác định";
                        AddEffect(eb);
                        return rhs;
                    }
                    bool cellSyntax = ma.Expression is ElementAccessExpressionSyntax eas && eas.Expression is MemberAccessExpressionSyntax em && em.Name.Identifier.Text is "Cells" or "Rows" or "Columns";
                    bool known = prop == "Formula" || prop == "Value" || prop == "Text" || prop == "Locked" || FormatProps.Contains(prop);
                    bool cellVar = ma.Expression is IdentifierNameSyntax cvid && st.Env.TryGetValue(cvid.Identifier.Text, out var cve) && cve.Val is VCell;
                    if (known || cellSyntax || cellVar)
                    {
                        var tgt = Eval(ma.Expression, st, f);
                        if (pendingThrow) return rhs;
                        if (tgt is VCell || prop == "Formula" || cellSyntax)
                        {
                            var ef = NewEffect(st, f, ln);
                            ef.Target = tgt is VCell ? tgt : new VOpaque("target:" + Kind(tgt));
                            string ck = tgt is VCell tc ? tc.Kind : null;
                            switch (prop)
                            {
                                case "Formula": ef.Op = "setFormula"; ef.Val = rhs; ef.Point = PointId(file, a); break;
                                case "Value" or "Text": ef.Op = "setValue"; ef.Val = rhs; ef.Params["prop"] = prop.ToLowerInvariant(); break;
                                case "Locked":
                                    if (rhs is VConst k && k.Val is bool b) ef.Op = b ? "lock" : "unlock";
                                    else { ef.Op = "lock"; ef.Params["locked"] = R.Expr(rhs, new Flags()); ef.Val = rhs; }
                                    break;
                                case "Visible":
                                    ef.Op = ck == "row" ? "setRowVisible" : ck == "col" ? "setColumnVisible" : "setCellVisible";
                                    ef.Params["visible"] = rhs is VConst kv && kv.Val is bool bv ? bv : R.Expr(rhs, new Flags()); ef.Val = rhs; break;
                                case "Width": ef.Op = "setColumnWidth"; ef.Params["width"] = R.Expr(rhs, new Flags()); ef.Val = rhs; break;
                                case "Height": ef.Op = "setRowHeight"; ef.Params["height"] = R.Expr(rhs, new Flags()); ef.Val = rhs; break;
                                case "Border": ef.Op = "setBorder"; ef.Params["border"] = R.Expr(rhs, new Flags()); break;
                                case "Note" or "NoteStyle": ef.Op = "setNote"; ef.Params["prop"] = prop; ef.Params["value"] = R.Expr(rhs, new Flags()); ef.Val = rhs; break;
                                case "DefaultValue" when tgt is not VCell:
                                    ef.Op = "setModelDefault"; ef.Target = null;
                                    ef.Params["model"] = R.Expr(tgt, new Flags()); ef.Params["value"] = R.Expr(rhs, new Flags());
                                    ef.Params["note"] = "giá trị mặc định trong mô hình dòng mẫu của bảng, không phải ô bảng tính";
                                    break;
                                default: ef.Op = "changeFormat"; ef.Params["prop"] = prop; ef.Params["value"] = R.Expr(rhs, new Flags()); break;
                            }
                            AddEffect(ef);
                            // kiểu ô nút có chữ chờ gán: tạo setButtonText cho ô này (đường đi tại lệnh gán kiểu ô)
                            if (prop == "CellType" && tgt is VCell && pendingBtn.TryGetValue(rhs, out var pend))
                                foreach (var eb in pend.ToList())
                                {
                                    var nb = NewEffect(st, f, eb.Line); nb.Op = "setButtonText"; nb.Target = tgt; nb.Val = eb.Val; nb.Via = eb.Via;
                                    nb.Params["text"] = eb.Params["text"]; nb.Params["cellTypeAt"] = ln;
                                    Effects.Add(nb);
                                    if (eb.Target == null) Effects.Remove(eb);
                                }
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
                    else { Assign(st, aid.Identifier.Text, VLin.Atom(aid.Identifier.Text), ln); RegSym(aid.Identifier.Text, "array", ln, "mảng sau khi gán phần tử theo chỉ số lúc chạy tại dòng " + ln); }
                    return rhs;
                }
        }
        return rhs;
    }

    static string Kind(V v) => v switch { VSym => "runtime", VOpaque => "opaque", _ => v.GetType().Name.Substring(1).ToLowerInvariant() };

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
        if (at != null) RegSym(n, "unknown", Line(at), "tên không có giá trị trong phạm vi phân tích (khai báo ở phạm vi khác hoặc kiểu tĩnh)");
        return new VSym(n);
    }

    public V Field(string n, ClassCtx cls)
    {
        if (cls.FieldType.TryGetValue(n, out var ft)) { if (ft == "FpSpread") return VSpread.I; if (ft == "SheetView") return new VSheet(new VSym("field(" + n + ")")); }
        if (cls.Mutable.Contains(n)) return new VSym("field(" + n + ")" + (fieldVer > 0 ? Mk.Tag(fieldVer) : ""));
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
        // giá trị có đọc ô: không lưu đệm (phiên bản ô thay đổi theo lệnh ghi)
        if (!R.Expr(v, new Flags()).Contains("cell(")) cls.Cache[n] = v;
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
                    return new VSym("newArray()");
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

    public V Member(V b, string n)
    {
        switch (b)
        {
            case VSpread:
                if (n == "Sheets") return VSheets.I;
                if (n == "ActiveSheet") return new VSheet(new VSym("activeSheet()"));
                if (n == "ActiveSheetIndex") return VLin.Atom("activeSheet()");
                return new VSym("spread()." + n);
            case VSheets: return new VSym("sheets()." + n);
            case VSheet s:
                if (n == "Cells") return new VCells(s, "cells");
                if (n == "Rows") return new VCells(s, "rows");
                if (n == "Columns") return new VCells(s, "cols");
                if (n == "FpSpread" || n == "Parent") return VSpread.I;
                if (n == "SheetName") return new VSym("sheetName(" + R.Sheet(s, new Flags()) + ")");
                return new VSym("sheetOf(" + R.Sheet(s, new Flags()) + ")." + n);
            case VCell c:
                {
                    var fl = new Flags(); string rf = R.CellRef(c, fl);
                    string pre = R.SheetPrefix(c.S, new Flags());
                    long ver = ReadVer(c);
                    return new VSym("cell(" + pre + rf + ")." + n.ToLowerInvariant() + (ver > 0 ? Mk.Tag(ver) : ""), fl.Opaque);
                }
            case VCellInfo ci:
                if (n == "CellTemplate" || n == "Location") return ci;
                if (n == "RowIndex") return VLin.Atom("rowIndex(" + ci.Path + ")");
                if (n == "ColIndex") return VLin.Atom("colIndex(" + ci.Path + ")");
                return new VSym("cellInfo(" + ci.Path + ")." + n);
            case VConst k when k.Val is string s2 && n == "Length": return new VConst((long)s2.Length);
            case VArr a when n == "Length" || n == "Count": return new VConst((long)a.Items.Count);
            case VOpaque:
                // điều hướng cấu trúc ô/bảng giữ nguyên; chỉ số dòng/cột viết dạng hàm trung tính
                if (n == "CellTemplate" || n == "Location") return b;
                if (n == "RowIndex") return VLin.Atom("rowIndex(" + R.Expr(b, new Flags()) + ")");
                if (n == "ColIndex") return VLin.Atom("colIndex(" + R.Expr(b, new Flags()) + ")");
                return new VSym(R.Expr(b, new Flags()) + "." + n, true);
            case VCases cs: return Ar.Cases(cs.Alts.Select(x => (x.When, Member(x.Val, n))).ToList());
            case VSym y:
                if (n == "Location" || n == "CellTemplate" || n == "Value" && y.Text.EndsWith(".value")) return y;
                if (n == "Cells") return new VSym(y.Text + ".cells");
                return new VSym(y.Text + "." + n);
            default:
                {
                    var fl = new Flags(); var t = R.Expr(b, fl);
                    if (b is VLin bl && !bl.IsSingleAtom(out _)) t = "(" + t + ")";
                    return new VSym(t + "." + n, fl.Opaque);
                }
        }
    }

    V EvalElement(ElementAccessExpressionSyntax ea, State st, Frame f)
    {
        var b = Eval(ea.Expression, st, f);
        var idx = ea.ArgumentList.Arguments.Select(a => Eval(a.Expression, st, f)).ToList();
        string hint = ea.Expression is IdentifierNameSyntax hid ? hid.Identifier.Text : null;
        if (hint != null && b is VArr ha && ha.Items.Count > Ar.MaxAlts && idx.Count == 1 && idx[0] is not VConst)
            RegSym(hint, "array", Line(ea), "mảng cục bộ " + ha.Items.Count + " phần tử, truy cập theo chỉ số lúc chạy");
        return Element(b, idx, hint);
    }
    V Element(V b, List<V> idx, string hint = null)
    {
        switch (b)
        {
            case VArr a when idx.Count == 1:
                if (idx[0] is VConst k && k.Val is long i) { if (i >= 0 && i < a.Items.Count) return a.Items[(int)i]; pendingThrow = true; return new VOpaque("index-out-of-range"); }
                if (a.Items.Count > 0 && a.Items.Count <= Ar.MaxAlts)
                    return Ar.Cases(a.Items.Select((it, j) => (j == a.Items.Count - 1 ? null : C.Cmp("==", idx[0], new VConst((long)j)), it)).ToList());
                {
                    var fa = new Flags(); var it2 = R.Expr(idx[0], fa);
                    return hint != null ? new VSym(hint + "[" + it2 + "]", fa.Opaque) : new VOpaque("array-index");
                }
            case VSheets when idx.Count == 1: return new VSheet(idx[0]);
            case VCells c:
                if (c.Kind == "cells" && idx.Count == 2) return new VCell(c.S, "cell", idx[0], idx[1]);
                if (c.Kind == "cells" && idx.Count == 4) return new VCell(c.S, "range", idx[0], idx[1], Ar.Add(Ar.Add(idx[0], idx[2]), new VConst(-1L)), Ar.Add(Ar.Add(idx[1], idx[3]), new VConst(-1L)));
                if (c.Kind == "rows" && idx.Count == 1) return new VCell(c.S, "row", idx[0], null);
                if (c.Kind == "cols" && idx.Count == 1) return new VCell(c.S, "col", null, idx[0]);
                return new VOpaque("cells-index");
            case VCases cs:
                {
                    // ngoại lệ chắc chắn chỉ khi mọi nhánh đều ngoài mảng
                    bool saved = pendingThrow, allThrow = true; var alts = new List<(Cond, V)>();
                    foreach (var x in cs.Alts) { pendingThrow = false; alts.Add((x.When, Element(x.Val, idx, hint))); allThrow &= pendingThrow; }
                    pendingThrow = saved || allThrow;
                    return Ar.Cases(alts);
                }
            case VOpaque: { var fo = new Flags(); return new VSym(R.Expr(b, fo) + "[" + string.Join(",", idx.Select(i => R.Expr(i, fo))) + "]", true); }
            case VConst ks when ks.Val is string s && idx.Count == 1 && idx[0] is VConst ki && ki.Val is long ii: return ii >= 0 && ii < s.Length ? new VConst(s[(int)ii].ToString()) : new VOpaque("index-out-of-range");
            default:
                {
                    var fl = new Flags(); var t = R.Expr(b, fl) + "[" + string.Join(",", idx.Select(i => R.Expr(i, fl))) + "]";
                    return new VSym(t, fl.Opaque);
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
                if (on != null) { Assign(st, on, new VSym(on + Mk.Tag(Mk.Next())), Line(inv)); RegSym(on, "out", Line(inv), "nhận qua đối số out/ref của lời gọi " + CallName(inv) + " tại dòng " + Line(inv)); }
            }
        string rt = recv?.ToString() ?? "";
        bool staticRecv = recv is IdentifierNameSyntax rid && !st.Env.ContainsKey(rid.Identifier.Text) && (f.Cls == null || !f.Cls.Members.Contains(rid.Identifier.Text)) || recv is PredefinedTypeSyntax;
        int ln = Line(inv);

        // nội tuyến method cùng class
        if ((recv == null || recv is ThisExpressionSyntax) && f.Cls != null && f.Cls.Methods.TryGetValue(name, out var cands))
        {
            var m = cands.FirstOrDefault(c => c.ParameterList.Parameters.Count == args.Count) ?? cands.FirstOrDefault(c => c.ParameterList.Parameters.Count >= args.Count && c.ParameterList.Parameters.Skip(args.Count).All(p => p.Default != null));
            var crec = new CallRec { Line = ln, Name = name, Depth = f.Depth + 1, EffBefore = Effects.Count, DropBefore = Dropped.Count };
            Calls.Add(crec);
            crec.Via = (f.Via != null ? f.Via + " > " : "") + name + "@" + ln;
            crec.Reason = m == null ? "không khớp chữ ký method cùng class (số đối số " + args.Count + ")" : f.Depth >= MaxDepth ? "vượt giới hạn nội tuyến " + MaxDepth + " cấp" : f.Stack.Contains(name + "/" + m.SpanStart) ? "đệ quy (method đang được nội tuyến trên ngăn gọi)" : null;
            if (crec.Reason != null && m != null)
            {
                // không theo được: lời gọi có thể ghi ô và trường lớp bất kỳ
                Wild(); fieldVer = Mk.Next();
                RegSym(name, "method", ln, "method cùng class không nội tuyến tại dòng " + ln + ": " + crec.Reason);
                crec.EffAfter = Effects.Count; crec.DropAfter = Dropped.Count;
                foreach (var a in args) Eval(a.Expression, st, f);
                if (isStmt) { var eo = NewEffect(st, f, ln); eo.Op = "opaque"; eo.Target = null; eo.Params["call"] = name; eo.Params["reason"] = crec.Reason; AddEffect(eo); }
                Opaque(name, inv, isStmt ? "statement" : "value");
                crec.EffAfter = Effects.Count;
                return new VOpaque("call", name);
            }
            if (m != null) crec.Followed = true;
            if (m != null)
            {
                var bind = new Dictionary<string, V>(StringComparer.Ordinal);
                var ps = m.ParameterList.Parameters;
                for (int i = 0; i < ps.Count; i++)
                {
                    ArgumentSyntax a = null;
                    var named = args.FirstOrDefault(x => x.NameColon?.Name.Identifier.Text == ps[i].Identifier.Text);
                    if (named != null) a = named; else if (i < args.Count && args[i].NameColon == null) a = args[i];
                    bind[ps[i].Identifier.Text] = a != null ? Eval(a.Expression, st, f) : ps[i].Default != null ? Eval(ps[i].Default.Value, st, f) : ParamSym(ps[i].Identifier.Text, m);
                }
                var savedUsed = used;
                var sub = new Frame { Cls = f.Cls, Method = name, Depth = f.Depth + 1, Via = (f.Via != null ? f.Via + " > " : "") + name + "@" + ln, Stack = f.Stack.Add(name + "/" + m.SpanStart) };
                crec.Via = sub.Via;
                var sst = new State { Env = new Dictionary<string, Entry>(StringComparer.Ordinal), Path = st.Path, Loops = st.Loops };
                foreach (var kv in bind) sst.Env[kv.Key] = new Entry(kv.Value, ImmutableSortedSet.CreateRange(savedUsed).Add(ln));
                // trường lớp đã gán trong hook hiện tại được truyền xuống
                foreach (var kv in st.Env) if (kv.Key.StartsWith("this.", StringComparison.Ordinal)) sst.Env[kv.Key] = kv.Value;
                sub.BasePathLen = st.Path.Count;
                if (m.Body != null) Exec(m.Body, sst, sub);
                else if (m.ExpressionBody != null) { used = new(); sub.Returns.Add((sst.Path, Eval(m.ExpressionBody.Expression, sst, sub))); }
                used = savedUsed;
                crec.EffAfter = Effects.Count; crec.DropAfter = Dropped.Count;
                foreach (var kv in sst.Env) if (kv.Key.StartsWith("this.", StringComparison.Ordinal)) st.Env[kv.Key] = kv.Value;
                // method có kiểu trả về nhưng không ghi nhận được lệnh return nào (ngoại lệ chắc chắn, vòng lặp): không giả định null
                if (sub.Returns.Count == 0)
                {
                    if (m is MethodDeclarationSyntax rmd && rmd.ReturnType.ToString() != "void") { RegSym(name, "method-result", ln, "giá trị trả về của method cùng class gọi tại dòng " + ln + " không xác định (không điểm kết thúc có giá trị nào tới được)"); return new VOpaque("result", name); }
                    return VConst.Null;
                }
                if (sub.Returns.Count == 1) return sub.Returns[0].Val;
                var alts = sub.Returns.Select(r => ((Cond)C.And(r.Path.Skip(st.Path.Count).Select(p => p.C)), r.Val)).ToList();
                var cv = Ar.Cases(alts.Select((a, i) => (i == alts.Count - 1 ? null : a.Item1, a.Item2)).ToList());
                if (cv is VOpaque) { RegSym(name, "method-result", ln, "giá trị trả về của method cùng class gọi tại dòng " + ln + " có quá " + Ar.MaxAlts + " nhánh"); return new VOpaque("result", name); }
                return cv;
            }
        }

        switch (name)
        {
            case "ArrayCellId" when args.Count == 1: return ArrayCellIdOf(Arg(inv, 0, st, f));
            case "ToIntOrDefault" or "ToLongOrDefault" or "ToDoubleOrDefault" or "ToDecimalOrDefault" when recv != null && !staticRecv:
                {
                    // phương thức mở rộng: toán hạng là đối tượng nhận
                    var v = Eval(recv, st, f);
                    bool isInt = name is "ToIntOrDefault" or "ToLongOrDefault";
                    if (v is VConst k && k.IsNum) return isInt && k.Val is double dk ? new VConst((long)dk) : k;
                    if (v is VConst ks && ks.Val is string ss) { if (isInt && long.TryParse(ss.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var li)) return new VConst(li); if (!isInt && double.TryParse(ss.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var di)) return new VConst(di); }
                    return VLin.Atom(((VSym)Fn(isInt ? "int" : "num", v)).Text);
                }
            case "Split" when staticRecv && rt == "Regex" && args.Count >= 2:
                return Fn("split", Arg(inv, 0, st, f), Arg(inv, 1, st, f));
            case "checkLockUnLockCells" when args.Count >= 2:
                {
                    // GeneralFunctions.checkLockUnLockCells: khóa/mở ô theo bảng cấu hình ở sheet cuối (đã đọc mã nguồn hàm)
                    var e1 = NewEffect(st, f, ln); e1.Op = "lock"; e1.Target = null;
                    e1.Params["via"] = "checkLockUnLockCells";
                    e1.Params["enabled"] = "ô BA1 của sheet cuối == \"1\"";
                    e1.Params["targets"] = "ô đích theo bảng cấu hình ở sheet cuối, mỗi dòng cấu hình: cột AK xác định sheet và bảng đích, cột AM liệt kê ô đích (nối bằng &)";
                    e1.Params["locked"] = "khóa khi (bảng động) kết quả công thức cột AL áp cho từng dòng dữ liệu == \"0\", hoặc (ô tĩnh) giá trị cột AL == 0, ngược lại mở khóa";
                    e1.Params["clearData"] = "ô cột AN của dòng cấu hình == \"1\"";
                    if (args.Count > 2) e1.Params["rowCheck"] = R.Expr(Arg(inv, 2, st, f), new Flags());
                    if (args.Count > 3) e1.Params["tableCheck"] = R.Expr(Arg(inv, 3, st, f), new Flags());
                    e1.Params["note"] = "ô đích theo cấu hình lúc chạy";
                    AddEffect(e1);
                    return VConst.Null;
                }
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
                    if (name == "ToString" && args.Count > 0) { var fv = Arg(inv, 0, st, f); return VCat.Make(new[] { Fn("format", v, fv) }); }
                    v = Ar.ToStr(v);
                    var targs = Enumerable.Range(0, args.Count).Select(i => Arg(inv, i, st, f)).ToList();
                    if (name.StartsWith("Trim", StringComparison.Ordinal) && targs.Count > 0)
                    {
                        var chs = targs.SelectMany(a => a is VArr aa ? aa.Items : new List<V> { a }).Select(StrConst).ToList();
                        if (v is VConst kt && kt.Val is string st0 && chs.All(c => c != null && c.Length == 1))
                        { var ca = chs.Select(c => c[0]).ToArray(); return new VConst(name == "Trim" ? st0.Trim(ca) : name == "TrimEnd" ? st0.TrimEnd(ca) : st0.TrimStart(ca)); }
                        return VCat.Make(new[] { Fn(name == "Trim" ? "trim" : name == "TrimEnd" ? "trimEnd" : "trimStart", new[] { v }.Concat(targs).ToArray()) });
                    }
                    if (v is VConst ks && ks.Val is string s) return new VConst(name switch { "Trim" => s.Trim(), "ToUpper" => s.ToUpperInvariant(), "ToLower" => s.ToLowerInvariant(), "TrimEnd" => s.TrimEnd(), "TrimStart" => s.TrimStart(), _ => s });
                    if (name is "ToUpper" or "ToLower") return Ar.ToStr(Fn(name == "ToUpper" ? "upper" : "lower", v));
                    if (name is "Trim" or "TrimEnd" or "TrimStart") return VCat.Make(new[] { Fn(name == "Trim" ? "trim" : name == "TrimEnd" ? "trimEnd" : "trimStart", v) });
                    return v;
                }
            case "Format" when staticRecv && (rt == "string" || rt == "String"):
                {
                    var fmt = Arg(inv, 0, st, f);
                    var vals = Enumerable.Range(1, args.Count - 1).Select(i => Ar.ToStr(Arg(inv, i, st, f))).ToList();
                    if (vals.Count == 1 && vals[0] is VArr va) vals = va.Items.Select(Ar.ToStr).ToList();
                    return Ar.Lift(fmt, new VConst(0L), (x, _) => FormatV(x, vals));
                }
            case "PadLeft" or "PadRight" when recv != null && !staticRecv && args.Count >= 1:
                {
                    var v = Ar.ToStr(Eval(recv, st, f)); var w = Arg(inv, 0, st, f); var pc = args.Count > 1 ? Arg(inv, 1, st, f) : new VConst(" ");
                    if (v is VConst kv && kv.Val is string sv && w is VConst kw && kw.Val is long wn && wn >= 0 && wn < 1000 && StrConst(pc) is string ps && ps.Length == 1)
                        return new VConst(name == "PadLeft" ? sv.PadLeft((int)wn, ps[0]) : sv.PadRight((int)wn, ps[0]));
                    return VCat.Make(new[] { Fn(name == "PadLeft" ? "padLeft" : "padRight", v, w, pc) });
                }
            case "TryGetValue" when recv != null && !staticRecv && args.Count == 2:
                {
                    var key = Arg(inv, 0, st, f);
                    string dn = MapName(recv);
                    if (!dn.Contains('.') && (f.Cls == null || !f.Cls.Members.Contains(dn))) RegSym(dn, "local-map", ln, "bảng tra (từ điển) cục bộ, nội dung nạp lúc chạy");
                    var oa = args[1];
                    string on = oa.Expression is IdentifierNameSyntax oi ? oi.Identifier.Text : oa.Expression is DeclarationExpressionSyntax de && de.Designation is SingleVariableDesignationSyntax sv ? sv.Identifier.Text : null;
                    if (on != null) RegSym(on, "out", ln, "").Desc = "giá trị tra từ bảng " + dn + " theo khóa " + R.Expr(key, new Flags()) + " tại dòng " + ln + ", không có khóa thì giá trị mặc định";
                    return new VCond(new CAtom("hasKey", new VSym(dn), key));
                }
            case "ContainsKey" when recv != null && !staticRecv && args.Count == 1:
                {
                    var key = Arg(inv, 0, st, f);
                    string dn = MapName(recv);
                    if (!dn.Contains('.') && (f.Cls == null || !f.Cls.Members.Contains(dn))) RegSym(dn, "local-map", ln, "bảng tra (từ điển) cục bộ, nội dung nạp lúc chạy");
                    return new VCond(new CAtom("hasKey", new VSym(dn), key));
                }
            case "ConvertToDateTime" or "ToDateTime" when args.Count >= 1:
                return Fn("toDate", Enumerable.Range(0, args.Count).Select(i => Arg(inv, i, st, f)).ToArray());
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
            case "Parse" or "ToInt32" or "ToInt64" or "ToInt16" or "ConvertToInt" or "ConvertToLong" or "ToDouble" or "ToDecimal" or "ConvertToDecimal" or "ConvertToDouble" or "ConvertToDouble1" or "ConvertToDecimal1" when args.Count >= 1:
                {
                    var v = Arg(inv, 0, st, f);
                    if (v is VCat vcat && vcat.Parts.Count == 1 && vcat.Parts[0] is not VConst) v = vcat.Parts[0];
                    bool isInt = !(name.Contains("Double") || name.Contains("Decimal")) && rt != "double" && rt != "decimal";
                    if (v is VConst k)
                    {
                        if (k.Val is string s) { if (isInt && long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var l)) return new VConst(l); if (double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return isInt ? new VConst((long)d) : new VConst(d); }
                        if (k.IsNum) return k;
                    }
                    if (v is VLin vl && (!isInt || !vl.T.Keys.Any(a => a.Contains("cell(") || a.Contains("num(")))) return v;
                    var fv = Fn(isInt ? "int" : "num", v);
                    return VLin.Atom(((VSym)fv).Text);
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
                    AddEffect(e1);
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
                    if (sh != null && s == null && fmt is VCat fcat)
                    {
                        var pieces = new List<List<V>> { new() };
                        foreach (var p in fcat.Parts)
                        {
                            if (p is VConst pk && pk.Val is string ps)
                            {
                                var seg = ps.Split("###");
                                for (int i = 0; i < seg.Length; i++) { if (i > 0) pieces.Add(new()); if (seg[i].Length > 0) pieces[^1].Add(new VConst(seg[i])); }
                            }
                            else pieces[^1].Add(p);
                        }
                        foreach (var pc in pieces)
                        {
                            if (pc.Count == 0) continue;
                            V tgt = null;
                            if (pc.Count == 1 && pc[0] is VConst pc0) { if (ArrayCellIdOf(pc0) is VArr ar0) tgt = new VCell(sh, "cell", ar0.Items[1], ar0.Items[0]); }
                            else if (pc.Count == 2 && pc[0] is VConst c0 && c0.Val is string s0 && s0.EndsWith("_") && Fmt.LettersToIndex(s0.TrimEnd('_')) > 0)
                                tgt = new VCell(sh, "cell", Ar.Add(Ar.ToLin(pc[1]) ?? pc[1], new VConst(-1L)), new VConst(Fmt.LettersToIndex(s0.TrimEnd('_')) - 1));
                            var e1 = NewEffect(st, f, ln); e1.Op = "changeFormat"; e1.Target = tgt ?? new VOpaque("format-template");
                            e1.Params["via"] = "changFormatCell"; e1.Params["format"] = R.Template(VCat.Make(pc), new Flags());
                            e1.Params["decimals"] = "2 nếu " + R.Expr(cur, new Flags()) + " == \"USD\", ngược lại 0";
                            AddEffect(e1);
                        }
                        return VConst.Null;
                    }
                    if (sh == null || s == null) break;
                    foreach (var part in s.Split("###"))
                    {
                        if (ArrayCellIdOf(new VConst(part)) is not VArr ar) continue;
                        var e1 = NewEffect(st, f, ln); e1.Op = "changeFormat"; e1.Target = new VCell(sh, "cell", ar.Items[1], ar.Items[0]);
                        e1.Params["via"] = "changFormatCell"; e1.Params["decimals"] = "2 nếu " + R.Expr(cur, new Flags()) + " == \"USD\", ngược lại 0";
                        e1.Val = cur;
                        AddEffect(e1);
                    }
                    return VConst.Null;
                }
            case string rn when rn.StartsWith("Reset", StringComparison.Ordinal) && rn.Length > 5 && recv != null && args.Count == 0 && !staticRecv:
                {
                    var t = Eval(recv, st, f);
                    if (t is VCell)
                    {
                        var e1 = NewEffect(st, f, ln); e1.Target = t; e1.Params["reset"] = true;
                        if (rn == "ResetValue") { e1.Op = "setValue"; e1.Val = VConst.Null; }
                        else if (rn == "ResetFormula") { e1.Op = "setFormula"; e1.Val = new VConst(""); }
                        else if (rn == "ResetText") { e1.Op = "setValue"; e1.Val = new VConst(""); e1.Params["prop"] = "text"; }
                        else { e1.Op = "resetFormat"; e1.Params["prop"] = rn.Substring(5); }
                        AddEffect(e1); return VConst.Null;
                    }
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
                    return a1 == null ? Fn("substring", Eval(recv, st, f), a0) : Fn("substring", Eval(recv, st, f), a0, a1);
                }
            case "ToBoolean" or "ConvertToBool" or "ToBoolOrDefault" when args.Count >= 1 || recv != null:
                {
                    var v = args.Count >= 1 && (staticRecv || recv == null) ? Arg(inv, 0, st, f) : Eval(recv, st, f);
                    if (v is VConst k && k.Val is string bs) return new VConst(bs.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) || bs.Trim() == "1");
                    if (v is VConst kb && kb.Val is bool) return v;
                    if (v is VCond) return v;
                    return new VCond(new CAtom("bool", v));
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
        bool classLocal = (recv == null || recv is ThisExpressionSyntax) && f.Cls != null && f.Cls.Methods.ContainsKey(name);
        if (classLocal) { RegSym(name, "method", ln, "method cùng class không nội tuyến tại dòng " + ln + ": không khớp chữ ký (số đối số " + args.Count + ")"); Wild(); fieldVer = Mk.Next(); }
        bool localRecv = rv is VArr || rv is VConst || (rv is VSym rsy && rsy.Text.StartsWith("new ", StringComparison.Ordinal));
        bool touches = SheetLike(rv) || argv.Any(SheetLike) || argv.Any(a => a is VSym ts0 && ts0.Text == "this");
        if (isStmt && !localRecv)
        {
            var e1 = NewEffect(st, f, ln); e1.Op = "opaque"; e1.Target = null;
            e1.Params["call"] = qn;
            e1.Params["args"] = argv.Select(a => R.Expr(a, new Flags())).ToList();
            e1.Params["touchesSheet"] = touches;
            AddEffect(e1);
        }
        else if (touches) Wild();
        return classLocal ? new VOpaque("call", name) : new VOpaque(qn);
    }

    void AddLock(State st, Frame f, int ln, VCell cell, V lk, string via, V del = null)
    {
        var e2 = NewEffect(st, f, ln); e2.Target = cell; e2.Params["via"] = via;
        if (lk is VConst k && k.Val is bool b) e2.Op = b ? "lock" : "unlock";
        else { e2.Op = "lock"; e2.Params["locked"] = R.Expr(lk, new Flags()); e2.Val = lk; }
        if (del != null && !(del is VConst dk && dk.Val is bool db && !db)) e2.Params["clearData"] = R.Expr(del, new Flags());
        AddEffect(e2);
    }

    static V Fn(string fn, params V[] a) => Ar.Fn(fn, a);
    /// tên bảng tra: định danh hoặc chuỗi định danh (Lớp.TrườngTĩnh), bỏ this.
    static string MapName(ExpressionSyntax e)
    {
        var parts = new List<string>();
        while (e is MemberAccessExpressionSyntax ma) { parts.Insert(0, ma.Name.Identifier.Text); e = ma.Expression; }
        if (e is IdentifierNameSyntax id) parts.Insert(0, id.Identifier.Text);
        else if (e is not ThisExpressionSyntax) return "map";
        return parts.Count == 0 ? "map" : string.Join(".", parts);
    }
    static string CallName(InvocationExpressionSyntax inv) => inv.Expression is MemberAccessExpressionSyntax ma ? (ma.Expression is IdentifierNameSyntax id ? id.Identifier.Text + "." : "") + ma.Name.Identifier.Text : inv.Expression is IdentifierNameSyntax i2 ? i2.Identifier.Text : "?";
    V UninitSym(string n, int line) { RegSym(n, "uninit", line, "khai báo tại dòng " + line + " không có giá trị đầu"); return new VSym(n); }
    V ParamSym(string n, BaseMethodDeclarationSyntax m) { RegSym(n, "param", Line(m), "tham số của method " + MethodName(m) + " không được truyền"); return new VSym(n); }

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
                var full = s.Substring(i + 1, j - i - 1); var spec = full.Split(':', ',')[0];
                if (!int.TryParse(spec, out var ai) || ai >= vals.Count) return new VOpaque("Format");
                parts.Add(new VConst(sb.ToString())); sb.Clear();
                parts.Add(full.Length > spec.Length ? Ar.Fn("format", vals[ai], new VConst(full.Substring(spec.Length))) : vals[ai]); i = j + 1; continue;
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
