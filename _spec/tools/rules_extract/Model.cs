// Mô hình giá trị ký hiệu và điều kiện trung tính. Chỉ dùng cho phân tích văn bản cú pháp; không chạy mã decompile.
using System.Globalization;
using System.Text;

namespace RulesExtract;

public sealed class Flags { public bool Sym; public bool Opaque; public readonly SortedSet<string> Notes = new(StringComparer.Ordinal); }

public abstract class V
{
    string _key;
    public string Key => _key ??= MakeKey();
    protected abstract string MakeKey();
}

public sealed class VConst : V
{
    public readonly object Val; // string, long, double, bool, null
    public VConst(object v) { Val = v; }
    public static readonly VConst Null = new(null);
    public bool IsStr => Val is string;
    public bool IsNum => Val is long || Val is double;
    protected override string MakeKey() => "k:" + (Val == null ? "null" : Val.GetType().Name + ":" + Fmt.Num(Val));
}

/// Biểu thức số tuyến tính: C + Σ coef*atom. Atom là tên placeholder trung tính.
public sealed class VLin : V
{
    public readonly SortedDictionary<string, long> T;
    public readonly long C;
    VLin(SortedDictionary<string, long> t, long c) { T = t; C = c; }
    public static V Make(SortedDictionary<string, long> t, long c)
    {
        var t2 = new SortedDictionary<string, long>(StringComparer.Ordinal);
        foreach (var kv in t) if (kv.Value != 0) t2[kv.Key] = kv.Value;
        if (t2.Count == 0) return new VConst(c);
        return new VLin(t2, c);
    }
    public static V Atom(string a) => Make(new SortedDictionary<string, long>(StringComparer.Ordinal) { [a] = 1 }, 0);
    public bool IsSingleAtom(out string atom)
    {
        atom = null;
        if (T.Count == 1 && C == 0) { var kv = T.First(); if (kv.Value == 1) { atom = kv.Key; return true; } }
        return false;
    }
    protected override string MakeKey() => "lin:" + string.Join(",", T.Select(kv => kv.Key + "*" + kv.Value)) + "+" + C;
}

/// Ký hiệu lúc chạy; HasOpaque: văn bản có chứa phần opaque(...) nên bản ghi phải ở trạng thái unresolved
public sealed class VSym : V { public readonly string Text; public readonly bool HasOpaque; public VSym(string t, bool op = false) { Text = t; HasOpaque = op; } protected override string MakeKey() => "sym:" + Text; }
/// Giá trị chưa giải. What chỉ mô tả loại (tên lời gọi API hoặc loại phân tích); SymName (nếu có) là tên biến/method C# đưa vào symbols.
/// Mỗi lần đánh giá có Id riêng: hai lần đánh giá khác nhau không bao giờ coi là cùng giá trị.
public sealed class VOpaque : V
{
    public readonly string What, SymName; public readonly long Id; public readonly List<V> Args;
    public VOpaque(string w, string sym = null) { What = w; SymName = sym; Id = Mk.Next(); }
    /// nhãn kèm đối số: opaque(Tên, đối số 1, ...)
    public VOpaque(string w, List<V> args) { What = w; Args = args; Id = Mk.Next(); }
    protected override string MakeKey() => "op:" + What + "#" + Id;
}
/// Đánh dấu ẩn trong văn bản ký hiệu: phân biệt danh tính (phiên bản ô sau lệnh ghi, lần đánh giá opaque, biến out) mà không đổi phần hiển thị.
/// Dạng \u001F<số>[:tên]\u001F; bị xóa khi ghi JSON; tên (nếu có) đi vào symbols.
public static class Mk
{
    public const char M = '\u001F';
    static long seq;
    public static long Next() => ++seq;
    public static string Tag(long id, string name = null) => M + id.ToString(CultureInfo.InvariantCulture) + (name != null ? ":" + name : "") + M;
    static readonly System.Text.RegularExpressions.Regex Rx = new("\u001F[^\u001F]*\u001F", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
    static readonly System.Text.RegularExpressions.Regex RxName = new("\u001F\\d+:([A-Za-z_]\\w*)\u001F", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
    public static string Strip(string s) => s != null && s.IndexOf(M) >= 0 ? Rx.Replace(s, "") : s;
    /// dấu ký hiệu có tên: tên + \u001Fs<id>\u001F (id của mục symbols trong Interp); đổi thành tên hoặc tên@dòng khi ghi bản ghi
    public static string SymTag(long id) => M + "s" + id.ToString(CultureInfo.InvariantCulture) + M;
    public static readonly System.Text.RegularExpressions.Regex RxSym = new("([A-Za-z_]\\w*)\u001Fs(\\d+)\u001F", System.Text.RegularExpressions.RegexOptions.CultureInvariant);
    public static bool HasSym(string s) => s != null && s.IndexOf(M) >= 0 && RxSym.IsMatch(s);
    /// bỏ mọi dấu ẩn trừ dấu ký hiệu có tên
    public static string StripKeepSym(string s) => s != null && s.IndexOf(M) >= 0 ? Rx.Replace(s, m => m.Value.Length > 2 && m.Value[1] == 's' && m.Value.Skip(2).Take(m.Value.Length - 3).All(char.IsDigit) ? m.Value : "") : s;
    public static IEnumerable<string> Names(string s) => s.IndexOf(M) < 0 ? Enumerable.Empty<string>() : RxName.Matches(s).Select(m => m.Groups[1].Value);
}
public sealed class VArr : V { public readonly List<V> Items; public VArr(List<V> i) { Items = i; } protected override string MakeKey() => "arr[" + string.Join(";", Items.Select(i => i.Key)) + "]"; }
public sealed class VCat : V
{
    public readonly List<V> Parts;
    VCat(List<V> p) { Parts = p; }
    public static V Make(IEnumerable<V> parts)
    {
        var res = new List<V>();
        foreach (var p0 in parts)
        {
            var ps = p0 is VCat c ? c.Parts : new List<V> { p0 };
            foreach (var p in ps)
            {
                V q = p is VConst k && k.Val != null && !(k.Val is string) ? new VConst(Fmt.Num(k.Val)) : p;
                if (q is VConst kn && kn.Val == null) q = new VConst("");
                if (q is VConst ks && ks.Val is string s && res.Count > 0 && res[^1] is VConst kp && kp.Val is string sp)
                    res[^1] = new VConst(sp + s);
                else if (!(q is VConst ke && ke.Val is string se && se.Length == 0)) res.Add(q);
            }
        }
        if (res.Count == 0) return new VConst("");
        if (res.Count == 1 && res[0] is VConst kk && kk.Val is string) return res[0];
        return new VCat(res);
    }
    protected override string MakeKey() => "cat(" + string.Join("|", Parts.Select(p => p.Key)) + ")";
}
public sealed class VCases : V
{
    public readonly List<(Cond When, V Val)> Alts; // When == null: còn lại
    public VCases(List<(Cond, V)> a) { Alts = a; }
    protected override string MakeKey() => "cases(" + string.Join("|", Alts.Select(a => (a.When?.Key ?? "else") + "=>" + a.Val.Key)) + ")";
}
public sealed class VCond : V { public readonly Cond C; public VCond(Cond c) { C = c; } protected override string MakeKey() => "cond:" + C.Key; }
public sealed class VSpread : V { public static readonly VSpread I = new(); protected override string MakeKey() => "spread"; }
public sealed class VSheets : V { public static readonly VSheets I = new(); protected override string MakeKey() => "sheets"; }
public sealed class VSheet : V { public readonly V Index; public VSheet(V i) { Index = i; } protected override string MakeKey() => "sheet(" + Index.Key + ")"; }
public sealed class VCells : V { public readonly VSheet S; public readonly string Kind; public VCells(VSheet s, string k) { S = s; Kind = k; } protected override string MakeKey() => Kind + "coll(" + S.Key + ")"; }
public sealed class VCell : V
{
    public readonly VSheet S; public readonly string Kind; public readonly V R1, C1, R2, C2; // 0-based
    public VCell(VSheet s, string kind, V r1, V c1, V r2 = null, V c2 = null) { S = s; Kind = kind; R1 = r1; C1 = c1; R2 = r2; C2 = c2; }
    protected override string MakeKey() => "cell(" + S.Key + "," + Kind + "," + R1?.Key + "," + C1?.Key + "," + R2?.Key + "," + C2?.Key + ")";
}
public sealed class VCellInfo : V { public readonly string Path; public VCellInfo(string p) { Path = p; } protected override string MakeKey() => "ci:" + Path; }
/// hàm vô danh: Text là dạng trung tính fn(tham số: thân) (bỏ kiểu, ép kiểu; biến bắt được có giá trị hằng thì thay giá trị)
public sealed class VLambda : V { public readonly Microsoft.CodeAnalysis.CSharp.Syntax.LambdaExpressionSyntax L; public readonly string Text; public VLambda(Microsoft.CodeAnalysis.CSharp.Syntax.LambdaExpressionSyntax l, string text) { L = l; Text = text; } protected override string MakeKey() => "lambda@" + L.SpanStart; }

public abstract class Cond
{
    string _key;
    public string Key => _key ??= MakeKey();
    protected abstract string MakeKey();
}
public sealed class CConst : Cond { public readonly bool B; CConst(bool b) { B = b; } public static readonly CConst T = new(true), F = new(false); protected override string MakeKey() => B ? "T" : "F"; }
public sealed class CAnd : Cond { public readonly List<Cond> Xs; public CAnd(List<Cond> x) { Xs = x; } protected override string MakeKey() => "and(" + string.Join(",", Xs.Select(x => x.Key)) + ")"; }
public sealed class COr : Cond { public readonly List<Cond> Xs; public COr(List<Cond> x) { Xs = x; } protected override string MakeKey() => "or(" + string.Join(",", Xs.Select(x => x.Key)) + ")"; }
public sealed class CNot : Cond { public readonly Cond X; public CNot(Cond x) { X = x; } protected override string MakeKey() => "not(" + X.Key + ")"; }
public sealed class CCmp : Cond { public readonly string Op; public readonly V A, B; public CCmp(string op, V a, V b) { Op = op; A = a; B = b; } protected override string MakeKey() => "cmp(" + Op + "," + A.Key + "," + B.Key + ")"; }
public sealed class CAtom : Cond { public readonly string Fn; public readonly V[] Args; public CAtom(string fn, params V[] a) { Fn = fn; Args = a; } protected override string MakeKey() => "atom(" + Fn + "," + string.Join(",", Args.Select(a => a.Key)) + ")"; }

public static class Fmt
{
    public static string Num(object v) => v switch
    {
        long l => l.ToString(CultureInfo.InvariantCulture),
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        bool b => b ? "true" : "false",
        string s => s,
        null => "null",
        _ => v.ToString()
    };
    public static string Letters(long v)
    {
        string t = "";
        while (--v >= 0) { t = (char)('A' + v % 26) + t; v /= 26; }
        return t;
    }
    public static long LettersToIndex(string s)
    {
        long n = 0;
        foreach (char ch in s) { if (ch < 'A' || ch > 'Z') return -1; n = n * 26 + (ch - 'A' + 1); }
        return n;
    }
}

public static class C
{
    public static Cond And(params Cond[] xs) => And((IEnumerable<Cond>)xs);
    public static Cond And(IEnumerable<Cond> xs)
    {
        var l = new List<Cond>(); var seen = new HashSet<string>();
        foreach (var x in xs)
        {
            if (x is CConst k) { if (!k.B) return CConst.F; continue; }
            foreach (var y in x is CAnd a ? a.Xs : new List<Cond> { x }) if (seen.Add(y.Key)) l.Add(y);
        }
        if (l.Count == 0) return CConst.T;
        return l.Count == 1 ? l[0] : new CAnd(l);
    }
    public static Cond Or(params Cond[] xs) => Or((IEnumerable<Cond>)xs);
    public static Cond Or(IEnumerable<Cond> xs)
    {
        var l = new List<Cond>(); var seen = new HashSet<string>();
        foreach (var x in xs)
        {
            if (x is CConst k) { if (k.B) return CConst.T; continue; }
            foreach (var y in x is COr o ? o.Xs : new List<Cond> { x }) if (seen.Add(y.Key)) l.Add(y);
        }
        if (l.Count == 0) return CConst.F;
        return l.Count == 1 ? l[0] : new COr(l);
    }
    static readonly Dictionary<string, string> Inv = new() { ["=="] = "!=", ["!="] = "==", ["<"] = ">=", [">="] = "<", [">"] = "<=", ["<="] = ">" };
    static readonly Dictionary<string, string> Mir = new() { ["=="] = "==", ["!="] = "!=", ["<"] = ">", [">="] = "<=", [">"] = "<", ["<="] = ">=" };
    public static Cond Not(Cond x) => x switch
    {
        CConst k => k.B ? CConst.F : CConst.T,
        CNot n => n.X,
        CCmp c => new CCmp(Inv[c.Op], c.A, c.B),
        COr o => And(o.Xs.Select(Not)),
        CAnd a => Or(a.Xs.Select(Not)),
        _ => new CNot(x)
    };
    public static Cond Cmp(string op, V a, V b)
    {
        if (a is VCases ca) return LiftCmp(op, ca, b, false);
        if (b is VCases cb) return LiftCmp(op, cb, a, true);
        if (a is VConst x && b is VConst y && !(a is null))
        {
            var r = ConstCmp(op, x.Val, y.Val);
            if (r.HasValue) return r.Value ? CConst.T : CConst.F;
        }
        // chuẩn hóa: đưa trigger.row/trigger.col (hệ số 1) về một vế riêng
        static string TrigAtom(V v) => v is VLin l ? l.T.Where(kv => kv.Value == 1 && (kv.Key == R.TRow || kv.Key == R.TCol)).Select(kv => kv.Key).FirstOrDefault() : null;
        if (TrigAtom(a) == null && TrigAtom(b) != null && Ar.ToLin(a) != null) { (a, b) = (b, a); op = Mir[op]; }
        if (TrigAtom(a) is string ta && Ar.ToLin(b) != null && !(a is VLin la0 && la0.IsSingleAtom(out _)))
        {
            var rest = Ar.Arith("-", a, VLin.Atom(ta));
            if (Ar.ToLin(rest) != null) { a = VLin.Atom(ta); b = Ar.Arith("-", b, rest); }
        }
        if (a is VConst && b is VLin) { (a, b) = (b, a); op = Mir[op]; }
        if (a is VLin la && b is VConst kb && kb.Val is long lb && la.T.Count == 1 && la.T.First().Value == 1 && la.C != 0)
        { a = VLin.Atom(la.T.First().Key); b = new VConst(lb - la.C); }
        if (a is VConst && !(b is VConst)) { (a, b) = (b, a); op = Mir[op]; }
        return new CCmp(op, a, b);
    }
    static Cond LiftCmp(string op, VCases cs, V other, bool swapped)
    {
        var ors = new List<Cond>(); var prior = new List<Cond>();
        foreach (var (w, v) in cs.Alts)
        {
            var when = w ?? Not(Or(prior));
            var cmp = swapped ? Cmp(op, other, v) : Cmp(op, v, other);
            ors.Add(And(when, cmp));
            if (w != null) prior.Add(w);
        }
        return Or(ors);
    }
    static bool? ConstCmp(string op, object x, object y)
    {
        if ((x is long || x is double) && (y is long || y is double))
        {
            double a = Convert.ToDouble(x, CultureInfo.InvariantCulture), b = Convert.ToDouble(y, CultureInfo.InvariantCulture);
            return op switch { "==" => a == b, "!=" => a != b, "<" => a < b, "<=" => a <= b, ">" => a > b, ">=" => a >= b, _ => null };
        }
        if (x is string || y is string || x == null || y == null || x is bool || y is bool)
        {
            if (op == "==") return Equals(x, y);
            if (op == "!=") return !Equals(x, y);
        }
        return null;
    }
    public static Cond Of(V v) => v switch
    {
        VCond c => c.C,
        VConst k when k.Val is bool b => b ? CConst.T : CConst.F,
        VCases cs => Or(cs.Alts.Select((a, i) => And(a.When ?? Not(Or(cs.Alts.Take(i).Where(p => p.When != null).Select(p => p.When))), Of(a.Val)))),
        _ => new CAtom("is", v)
    };
}

/// Hiển thị trung tính
public static class R
{
    public const string TRow = "trigger.row", TCol = "trigger.col";
    public static string Expr(V v, Flags f)
    {
        switch (v)
        {
            case VConst k:
                if (k.Val is string s) return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t") + "\"";
                return Fmt.Num(k.Val);
            case VLin l:
                {
                    f.Sym = true;
                    var sb = new StringBuilder();
                    foreach (var kv in l.T)
                    {
                        string t = kv.Value == 1 ? kv.Key : kv.Value == -1 ? "-" + kv.Key : kv.Value + "*" + kv.Key;
                        if (sb.Length > 0 && !t.StartsWith("-")) sb.Append('+');
                        sb.Append(t);
                        f.Notes.Add(kv.Key);
                        if (kv.Key.Contains("opaque(")) f.Opaque = true;
                    }
                    if (l.C > 0) sb.Append('+').Append(l.C); else if (l.C < 0) sb.Append(l.C);
                    return sb.ToString();
                }
            case VSym y: f.Sym = true; if (y.HasOpaque) f.Opaque = true; f.Notes.Add(y.Text); return y.Text;
            case VOpaque o: f.Opaque = true; f.Notes.Add("opaque:" + o.What); return "opaque(" + o.What + (o.Args is { Count: > 0 } ? ", " + string.Join(", ", o.Args.Select(a => a == null ? "null" : Expr(a, f))) : "") + ")" + Mk.Tag(o.Id, o.SymName);
            case VArr a: return "[" + string.Join(", ", a.Items.Select(i => Expr(i, f))) + "]";
            case VCat c: return "\"" + Template(c, f) + "\"";
            case VCases cs: return "cases(" + string.Join(" | ", cs.Alts.Select(a => (a.When == null ? "else" : Cond(a.When, f)) + " -> " + Expr(a.Val, f))) + ")";
            case VCond cc: return Cond(cc.C, f);
            case VCell cl: return "cell(" + SheetPrefix(cl.S, f) + CellRef(cl, f) + ")";
            case VSheet sh: return "sheet(" + Expr(sh.Index, f) + ")";
            case VCellInfo ci: f.Sym = true; return "cellInfo(" + ci.Path + ")";
            case VSpread: return "spread()";
            case VSheets: return "sheets()";
            case VCells cs2: return "cells(" + Expr(cs2.S.Index, f) + ")";
            case VLambda la: return la.Text;
            default: f.Opaque = true; return "opaque(?)";
        }
    }
    /// Chuỗi công thức/giá trị: hằng ghi nguyên, phần lúc chạy đặt trong {…}
    public static string Template(V v, Flags f)
    {
        switch (v)
        {
            case VConst k: return k.Val is string s ? s : Fmt.Num(k.Val);
            case VCat c: return string.Concat(c.Parts.Select(p => p is VConst pk ? Fmt.Num(pk.Val) : "{" + Expr(p, f) + "}"));
            default: return "{" + Expr(v, f) + "}";
        }
    }
    public static string ColPart(V c0, Flags f)
    {
        var c1 = Ar.Add(c0, new VConst(1L));
        if (c1 is VConst k && k.Val is long n) { if (n >= 1) return Fmt.Letters(n); f.Opaque = true; f.Notes.Add("invalid column " + n); return "#COL" + n; }
        return "{col(" + Expr(c1, f) + ")}";
    }
    public static string RowPart(V r0, Flags f)
    {
        var r1 = Ar.Add(r0, new VConst(1L));
        if (r1 is VConst k && k.Val is long n) { if (n >= 1) return n.ToString(CultureInfo.InvariantCulture); f.Opaque = true; f.Notes.Add("invalid row " + n); return "#ROW" + n; }
        return "{" + Expr(r1, f) + "}";
    }
    public static string CellRef(VCell c, Flags f) => c.Kind switch
    {
        "cell" => ColPart(c.C1, f) + RowPart(c.R1, f),
        "range" => ColPart(c.C1, f) + RowPart(c.R1, f) + ":" + ColPart(c.C2, f) + RowPart(c.R2, f),
        "row" => RowPart(c.R1, f) + ":" + RowPart(c.R1, f),
        "col" => ColPart(c.C1, f) + ":" + ColPart(c.C1, f),
        _ => "?"
    };
    public static string Sheet(VSheet s, Flags f) => s.Index is VConst k ? Fmt.Num(k.Val) : Expr(s.Index, f);
    /// Tiền tố sheet tường minh: S<n>! khi biết chỉ số, S[biểu thức]! khi chỉ số lúc chạy (activeSheet, tham số)
    public static string SheetPrefix(VSheet s, Flags f) => s.Index is VConst k && k.Val is long n ? "S" + n + "!" : "S[" + Expr(s.Index, f) + "]!";
    static int Prec(Cond c) => c switch { COr => 1, CAnd => 2, _ => 3 };
    public static string Cond(Cond c, Flags f)
    {
        switch (c)
        {
            case CConst k: return k.B ? "true" : "false";
            case CAnd a: return string.Join(" && ", a.Xs.Select(x => Prec(x) < 2 ? "(" + Cond(x, f) + ")" : Cond(x, f)));
            case COr o: return string.Join(" || ", o.Xs.Select(x => Cond(x, f)));
            case CNot n: return "!(" + Cond(n.X, f) + ")";
            case CCmp m:
                {
                    string a = Expr(m.A, f), b = Expr(m.B, f);
                    if (m.A is VLin la && la.IsSingleAtom(out var at) && at == TCol && m.B is VConst kb && kb.Val is long cn && cn >= 1) b = Fmt.Letters(cn);
                    return a + " " + m.Op + " " + b;
                }
            case CAtom t:
                if (t.Fn == "is") return Expr(t.Args[0], f);
                return t.Fn + "(" + string.Join(", ", t.Args.Select(x => Expr(x, f))) + ")";
        }
        return "?";
    }
}

/// Phép toán trên giá trị ký hiệu
public static class Ar
{
    public const int MaxAlts = 8;
    public static V ToLin(V v) => v switch
    {
        VConst k when k.Val is long => v,
        VLin => v,
        VSym s => VLin.Atom(s.Text),
        VConst k when k.Val is string s && long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) => new VConst(n),
        _ => null
    };
    public static V Lift(V a, V b, Func<V, V, V> op)
    {
        if (a is VCases ca)
        {
            if (ca.Alts.Count > MaxAlts) return new VOpaque("too-many-cases");
            return Cases(ca.Alts.Select(x => (x.When, Lift(x.Val, b, op))).ToList());
        }
        if (b is VCases cb)
        {
            if (cb.Alts.Count > MaxAlts) return new VOpaque("too-many-cases");
            return Cases(cb.Alts.Select(x => (x.When, Lift(a, x.Val, op))).ToList());
        }
        return op(a, b);
    }
    public static V Cases(List<(Cond When, V Val)> alts)
    {
        // gộp, phẳng hóa
        var flat = new List<(Cond, V)>();
        var prior = new List<Cond>();
        foreach (var (w0, v) in alts)
        {
            if (w0 is CConst k && !k.B) continue;
            var w = w0;
            if (v is VCases inner)
            {
                var ip = new List<Cond>();
                foreach (var (iw, iv) in inner.Alts)
                {
                    var ww = iw ?? C.Not(C.Or(ip));
                    flat.Add((w == null ? (iw == null && ip.Count == 0 ? null : ww) : C.And(w, ww), iv));
                    if (iw != null) ip.Add(iw);
                }
            }
            else flat.Add((w, v));
            if (w == null || (w is CConst kt && kt.B)) break;
        }
        if (flat.Count == 0) return new VOpaque("no-case");
        if (flat.All(x => x.Item2.Key == flat[0].Item2.Key)) return flat[0].Item2;
        if (flat.Count > MaxAlts) return new VOpaque("too-many-cases");
        if (flat[^1].Item1 != null) flat[^1] = (null, flat[^1].Item2); // nhánh cuối là phần còn lại
        return new VCases(flat);
    }
    public static V Add(V a, V b) => Lift(a, b, (x, y) => Arith("+", x, y));
    public static V Arith(string op, V a, V b)
    {
        if (a is VConst x && b is VConst y && x.IsNum && y.IsNum)
        {
            if (x.Val is long p && y.Val is long q)
                return op switch { "+" => new VConst(p + q), "-" => new VConst(p - q), "*" => new VConst(p * q), "/" => q == 0 ? new VOpaque("div0") : new VConst(p / q), "%" => q == 0 ? new VOpaque("div0") : new VConst(p % q), _ => new VOpaque("op" + op) };
            double d1 = Convert.ToDouble(x.Val, CultureInfo.InvariantCulture), d2 = Convert.ToDouble(y.Val, CultureInfo.InvariantCulture);
            return op switch { "+" => new VConst(d1 + d2), "-" => new VConst(d1 - d2), "*" => new VConst(d1 * d2), "/" => new VConst(d1 / d2), _ => new VOpaque("op" + op) };
        }
        var la = ToLin(a); var lb = ToLin(b);
        if (la != null && lb != null && (op == "+" || op == "-"))
        {
            var t = new SortedDictionary<string, long>(StringComparer.Ordinal); long c = 0; long sg = op == "+" ? 1 : -1;
            void acc(V v, long s) { if (v is VConst k) c += s * (long)k.Val; else foreach (var kv in ((VLin)v).T) { t.TryGetValue(kv.Key, out var o); t[kv.Key] = o + s * kv.Value; } if (v is VLin l2) c += s * l2.C; }
            acc(la, 1); acc(lb, sg);
            return VLin.Make(t, c);
        }
        if (la != null && lb != null && op == "*" && (la is VConst || lb is VConst))
        {
            var (k, l) = la is VConst ? ((long)((VConst)la).Val, lb) : ((long)((VConst)lb).Val, la);
            if (l is VConst kk) return new VConst(k * (long)kk.Val);
            var ll = (VLin)l;
            return VLin.Make(new SortedDictionary<string, long>(ll.T.ToDictionary(kv => kv.Key, kv => kv.Value * k), StringComparer.Ordinal), ll.C * k);
        }
        var f = new Flags();
        string txt = "(" + R.Expr(a, f) + op + R.Expr(b, f) + ")";
        return new VSym(txt, f.Opaque);
    }
    public static bool IsStringy(V v) => v is VCat || (v is VConst k && k.Val is string) || (v is VCases cs && cs.Alts.Any(a => IsStringy(a.Val)));
    public static V Concat(V a, V b) => Lift(a, b, (x, y) => VCat.Make(new[] { x, y }));
    /// Replace: hằng thì thay trực tiếp; phần ký hiệu bọc replace(x, a, b). Mẫu 1 ký tự thay từng phần (không vắt qua ranh giới);
    /// mẫu dài hơn bọc cả chuỗi để không sai khi mẫu vắt qua ranh giới hằng/ký hiệu.
    public static V Replace(V s, V a, V b)
    {
        return Lift(s, b, (x, y) =>
        {
            string from = a is VConst ka && ka.Val is string fs ? fs : null;
            if (from == null || from.Length == 0) return Fn("replace", x, a, y);
            if (x is VConst kx && kx.Val is string sx && y is VConst ky && ky.Val is string sy) return new VConst(sx.Replace(from, sy));
            if (x is VConst kn && kn.IsNum && y is VConst ky2 && ky2.Val is string sy2) return new VConst(Fmt.Num(kn.Val).Replace(from, sy2));
            var parts = x is VCat c ? c.Parts : new List<V> { x };
            // mẫu nhiều ký tự chỉ thay từng phần khi không thể vắt qua ranh giới: phần ký hiệu đều là số và mẫu không chứa ký tự số
            bool safe = from.Length == 1 || SafeSplit(parts, from);
            if (!safe && parts.Any(p => !(p is VConst))) return VCat.Make(new[] { Fn("replace", x, a, y) });
            var res = new List<V>();
            foreach (var p in parts)
            {
                if (p is VConst k && k.Val is string str)
                {
                    var segs = str.Split(from);
                    for (int i = 0; i < segs.Length; i++) { if (i > 0) res.Add(y); res.Add(new VConst(segs[i])); }
                }
                else if (p is VConst kn2 && kn2.IsNum && y is VConst ky3 && ky3.Val is string sy3) res.Add(new VConst(Fmt.Num(kn2.Val).Replace(from, sy3)));
                else if (p is VLin && from.Length > 1 && safe) res.Add(p); // số không chứa mẫu (mẫu có ký tự không phải số)
                else res.Add(Fn("replace", p, a, y));
            }
            return VCat.Make(res);
        });
    }
    static bool NumCh(char ch) => "0123456789+-.,Ee".IndexOf(ch) >= 0;
    /// Thay từng phần an toàn khi mọi phần ký hiệu là số (VLin), mẫu có ký tự không phải số ở đầu và cuối,
    /// và không có phần hằng đứng trước kết thúc bằng đoạn đầu của mẫu tới một ký tự số (trường hợp mẫu bao trọn phần số)
    static bool SafeSplit(List<V> parts, string from)
    {
        if (parts.Any(p => !(p is VConst || p is VLin))) return false;
        if (!from.Any(NumCh)) return true;
        if (!parts.Any(p => p is VLin)) return true;
        if (NumCh(from[0]) || NumCh(from[^1])) return false;
        for (int i = 0; i < parts.Count; i++)
        {
            if (parts[i] is not VLin || i == 0) continue;
            if (parts[i - 1] is not VConst pc || pc.Val is not string ps) return false;
            for (int j = 1; j < from.Length; j++) if (NumCh(from[j]) && ps.EndsWith(from.Substring(0, j), StringComparison.Ordinal)) return false;
        }
        return true;
    }
    /// Hàm trung tính trên ký hiệu: giữ cấu trúc; nếu đối số có phần opaque thì đánh dấu HasOpaque (không bỏ phép)
    public static V Fn(string fn, params V[] a)
    {
        var fl = new Flags(); var t = fn + "(" + string.Join(", ", a.Select(x => x == null ? "null" : R.Expr(x, fl))) + ")";
        return new VSym(t, fl.Opaque);
    }
    public static V ToStr(V v) => v switch
    {
        VConst k when k.IsNum => new VConst(Fmt.Num(k.Val)),
        VConst k when k.Val is bool b => new VConst(b ? "True" : "False"),
        VCases cs => Cases(cs.Alts.Select(a => (a.When, ToStr(a.Val))).ToList()),
        VLin or VSym => VCat.Make(new[] { v }),
        _ => v
    };
}
