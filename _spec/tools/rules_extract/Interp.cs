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
public sealed class SymInfo { public string Name, Origin, Desc; public int Line; public long Id; public string Ref => Name + Mk.SymTag(Id); }

public sealed class State
{
    public Dictionary<string, Entry> Env;
    public ImmutableList<PathItem> Path;
    public ImmutableList<LoopInfo> Loops;
    public bool Dead, Threw;
    // phiên bản ô theo đường đi, chỉ do lệnh đổi giá trị ô: ô hằng (CV, khóa vùng một ô), vùng có chỉ số lúc chạy (RV: lệnh ghi theo
    // cột/hàng chỉ phủ cột/hàng đó, '*' = bất kỳ), lệnh/lời gọi có thể ghi ô bất kỳ (WildV)
    public ImmutableDictionary<string, long> CV = ImmutableDictionary.Create<string, long>(StringComparer.Ordinal), RV = ImmutableDictionary.Create<string, long>(StringComparer.Ordinal);
    public long WildV;
    public State Clone() => new() { Env = new Dictionary<string, Entry>(Env, StringComparer.Ordinal), Path = Path, Loops = Loops, Dead = Dead, Threw = Threw, CV = CV, RV = RV, WildV = WildV };
    public void CopyVers(State o) { CV = o.CV; RV = o.RV; WildV = o.WildV; }
}

public sealed class Frame
{
    public ClassCtx Cls; public string Method; public int Depth; public string Via;
    public List<(ImmutableList<PathItem> Path, V Val)> Returns = new();
    public ImmutableHashSet<string> Stack = ImmutableHashSet<string>.Empty;
    public int BasePathLen;
    public List<State> Exits; // phiên bản ô tại các lệnh return của method nội tuyến
    public Dictionary<string, string> DlgSkip; // biến nhận kết quả hộp thoại trong nhánh hằng không chạy -> nguyên nhân
}

public sealed class Effect
{
    public string Op; public V Target; public V Val; public SortedDictionary<string, object> Params = new(StringComparer.Ordinal);
    public ImmutableList<PathItem> Path; public ImmutableList<LoopInfo> Loops; public int Line; public string Via; public string Point;
    public ImmutableSortedSet<int> Lines;
}

/// Một lần gọi method cùng class: đã nội tuyến hay chưa, lý do, khoảng effect/dropped sinh ra trong lần gọi
public sealed class CallRec { public int Line; public string Name, Via, Reason; public int Depth; public bool Followed; public int EffBefore, EffAfter, DropBefore, DropAfter; }

/// Bảng cấu hình khóa/mở ô ở sheet cuối của mẫu .xls (lock_config.json do lock_config.py sinh)
public sealed class LockRow { public int Row, Sheet, RefSheet, ControlRow; public string ControlCol, AL, ALn, AN, Skip, Error; public bool ANformula, AMempty; public object ALvalue; public List<(int Row, string Col)> Targets = new(); }
public sealed class LockCfg { public string Template, BA1, Stop, Error; public int LastSheet; public readonly List<LockRow> Rows = new(); }

public sealed class Interp
{
    public LockCfg Lock;
    public readonly List<Effect> Effects = new();
    public readonly List<CallRec> Calls = new();
    // phiên bản đọc ô: mỗi lệnh ghi (trực tiếp, qua method nội tuyến, hoặc lời gọi có thể ghi) tăng phiên bản; hai lần đọc khác phiên bản không coi là cùng giá trị
    long fieldVer;
    readonly Dictionary<long, (int Line, string Desc)> verInfo = new();
    readonly Dictionary<string, Dictionary<long, string>> verLabel = new(StringComparer.Ordinal);
    /// "cell(S0!L8)#74" -> mô tả sự kiện tạo phiên bản (lệnh ghi, lời gọi có thể ghi, hợp nhánh)
    public readonly Dictionary<string, string> VerRefs = new(StringComparer.Ordinal);
    readonly Dictionary<string, (int Line, string Type, string Code, int Len)> dialogForms = new(StringComparer.Ordinal);
    /// lời gọi chỉ tác động giao diện (bố cục, cuộn, chọn ô): không ghi ô, không tạo phiên bản ô
    public static readonly HashSet<string> UiCalls = new(StringComparer.Ordinal) { "SuspendLayout", "ResumeLayout", "SetActiveCell", "Focus", "ShowCell", "ShowActiveCell", "Invalidate", "Select", "BringToFront", "SetViewportTopRow", "SetViewportLeftColumn", "SetActiveSheet", "SetNodeBackGroundPink", "SetRowHeight", "SetColumnWidth", "Refresh", "Update", "PerformLayout", "ScrollTo", "SetFocus" };
    /// lời gọi không đổi giá trị ô (đọc ô, ghi chú, xóa kiểm tra, màu nền, khóa, ghi log): đã đối chiếu thân method trong _decompiled
    /// (GeneralFunctions/HTKKFile/ConvertHelpers: không gán Value/Text/Formula, không ResetValue/Clear/chèn hàng); không tạo phiên bản ô
    public static readonly HashSet<string> NoWriteCalls = new(StringComparer.Ordinal) { "RemoveValidate", "RemoveValidateCell", "GetMessage", "GetRowIndexInSViewByCompares", "ConvertCellValueToText", "UpdateFromSViewToSInfo", "GetValue", "ChangeLockedCell", "GetPreferredHeight", "SetNodeBackGroundWhite", "SetNote", "ResetNote", "GetCurrentMethod", "Log", "LogErrorBigData" };
    /// lời gọi đổi giá trị ô hoặc dịch ô (chèn/xóa hàng, xóa vùng, helper ghi giá trị đã đọc thân): phiên bản kind "ghi"
    public static readonly HashSet<string> WriteCalls = new(StringComparer.Ordinal) { "AddRows", "RemoveRows", "AddRow", "RemoveRow", "InsertRows", "DeleteRows", "InsertTableRow", "RemoveDataRows", "ClearRange", "ClearData", "ClearSelection", "ClearAll", "InsertNumberCell", "InsertCombobox", "resetOtherCheckboxStatic", "ChangeValueCell", "SetValue", "SetText", "SetFormula", "CopyRange", "MoveRange", "SwapRange", "ResetValue", "ResetText", "ResetFormula" };
    /// thao tác danh sách cục bộ trong thân vòng lặp (không phải Rows/Columns): không ghi ô
    static readonly HashSet<string> ListOps = new(StringComparer.Ordinal) { "Remove", "RemoveAt", "Insert", "Clear", "FindIndex", "FindAll", "FromArgb", "FromHtml", "GetType", "Value" };
    /// phân loại lời gọi: null = không ghi ô; "ghi" = đổi giá trị hoặc dịch ô; "co_the_ghi" = chưa rõ có ghi ô (bảo thủ)
    public static string CallWriteKind(string qn)
    {
        string nm = qn.Split('.').Last();
        if (UiCalls.Contains(nm) || NoWriteCalls.Contains(nm) || nm.StartsWith("SetNodeBackGround", StringComparison.Ordinal) || nm.StartsWith("RemoveValidate", StringComparison.Ordinal)) return null;
        if (WriteCalls.Contains(nm)) return "ghi";
        if (nm.StartsWith("Reset", StringComparison.Ordinal) && !nm.Contains('_')) return null; // ResetBackColor, ResetLocked...: trả định dạng
        if (qn.StartsWith("<cells>.", StringComparison.Ordinal) && nm is "Add" or "Remove" or "Insert" or "Clear") return "ghi";
        return "co_the_ghi";
    }
    /// thân vòng lặp (kể cả method cùng class được gọi, bắc cầu) có lệnh đổi giá trị ô ("ghi") hoặc lời gọi có thể ghi ("co_the_ghi" + tên)
    (string Kind, string[] Calls) LoopWrites(SyntaxNode node, ClassCtx cls)
    {
        bool w = false; var calls = new SortedSet<string>(StringComparer.Ordinal);
        var seen = new HashSet<SyntaxNode>(); var todo = new Stack<SyntaxNode>();
        // chỉ phần chạy lặp lại: thân, điều kiện, bước; biểu thức tập duyệt của foreach chạy một lần trước vòng
        IEnumerable<SyntaxNode> parts = node switch
        {
            ForEachStatementSyntax fe0 => new SyntaxNode[] { fe0.Statement },
            ForStatementSyntax fs0 => new SyntaxNode[] { fs0.Statement, fs0.Condition }.Concat(fs0.Incrementors),
            WhileStatementSyntax ws0 => new SyntaxNode[] { ws0.Statement, ws0.Condition },
            DoStatementSyntax ds0 => new SyntaxNode[] { ds0.Statement, ds0.Condition },
            _ => new[] { node }
        };
        foreach (var pn in parts) if (pn != null) todo.Push(pn);
        while (todo.Count > 0)
        {
            var b = todo.Pop(); if (!seen.Add(b)) continue;
            foreach (var n in b.DescendantNodesAndSelf())
            {
                if (n is AssignmentExpressionSyntax a && a.Left is MemberAccessExpressionSyntax lm && lm.Name.Identifier.Text is "Value" or "Text" or "Formula") { w = true; continue; }
                if (n is not InvocationExpressionSyntax iv) continue;
                string nm = iv.Expression switch { IdentifierNameSyntax i => i.Identifier.Text, GenericNameSyntax g => g.Identifier.Text, MemberAccessExpressionSyntax m => m.Name.Identifier.Text, _ => "?" };
                bool local = iv.Expression is IdentifierNameSyntax || (iv.Expression is MemberAccessExpressionSyntax tm && tm.Expression is ThisExpressionSyntax);
                if (local && cls != null && cls.Methods.TryGetValue(nm, out var ms)) { foreach (var m in ms) todo.Push(m); continue; }
                if (iv.Expression is MemberAccessExpressionSyntax rm && rm.Expression is MemberAccessExpressionSyntax rm2 && rm2.Name.Identifier.Text is "Rows" or "Columns" && nm is "Add" or "Remove" or "Insert" or "Clear") { w = true; continue; }
                if (PureCalls.Contains(nm) || ListOps.Contains(nm)) continue;
                var k = CallWriteKind(nm);
                if (k == "ghi") w = true; else if (k == "co_the_ghi") calls.Add(nm);
            }
        }
        return (w ? "ghi" : calls.Count > 0 ? "co_the_ghi" : null, calls.ToArray());
    }
    /// phiên bản trước vòng và phiên bản đầu vòng của các vòng lặp đang phân tích (EnterLoop đẩy, Loop lấy ra)
    readonly Stack<(long, long)> loopVers = new();
    /// đích của break/continue: vòng lặp (Sw = false) hoặc switch (Sw = true); lưu trạng thái tại lệnh nhảy để hợp ở cuối thân vòng
    /// (continue), sau vòng hoặc sau nhánh switch (break). Câu lệnh sau lệnh nhảy trong cùng khối không chạy trên đường đó.
    sealed class JumpCtx { public bool Sw; public readonly List<State> Cont = new(), Brk = new(); }
    readonly List<JumpCtx> jumps = new();
    /// ký hiệu đầu vòng của biến khai báo kiểu chuỗi/đối tượng
    readonly HashSet<string> strAtoms = new(StringComparer.Ordinal);
    /// hợp các trạng thái còn sống vào target: đường đi chung là tiền tố chung, phần sau tiền tố là điều kiện của từng trạng thái
    void Join(State target, List<State> states, int line)
    {
        var live = states.Where(s => !s.Dead).ToList();
        if (live.Count == 0) { target.Dead = true; return; }
        int n = live.Min(s => s.Path.Count), p = 0;
        while (p < n && live.All(s => ReferenceEquals(s.Path[p], live[0].Path[p]))) p++;
        target.Path = live[0].Path.GetRange(0, p); target.Dead = false; target.Threw = false;
        Merge(target, live.Select(s => (C.And(s.Path.Skip(p).Select(x => x.C)), s)).ToList(), line);
    }
    /// phiên bản đầu vòng của các vòng đang phân tích: hợp nhánh trong thân vòng không lấy dòng đầu vòng làm dòng ghi gốc
    /// (trừ khi chỉ còn nguồn đó), để dòng đầu vòng không theo phiên bản ra ngoài vòng
    readonly HashSet<long> activeLoop = new();
    List<long> SrcVers(IEnumerable<long> vs)
    {
        var src = vs.Where(v => v > 0 && verWrites.ContainsKey(v)).Distinct().ToList();
        var nl = src.Where(v => !activeLoop.Contains(v)).ToList();
        return nl.Count > 0 ? nl : src;
    }
    /// dòng lệnh ghi (hoặc vòng lặp có ghi) gốc của mỗi phiên bản; phiên bản hợp nhánh lấy hợp các dòng gốc của nhánh
    readonly Dictionary<long, SortedSet<int>> verWrites = new();
    /// loại phiên bản: ghi | co_the_ghi, kèm tên lời gọi; hợp nhánh là co_the_ghi nếu một nguồn là co_the_ghi
    readonly Dictionary<long, (string Kind, SortedSet<string> Calls)> verKind = new();
    /// "cell(S0!L8)#74" -> loại, dòng lệnh ghi gốc, tên lời gọi
    public readonly Dictionary<string, (string Kind, int[] Writes, string[] Calls)> VerMeta = new(StringComparer.Ordinal);
    /// phiên bản gốc (lệnh ghi hoặc lời gọi) của mỗi phiên bản: theo dòng ghi (verOrig, như verWrites) và theo loại (verOrigK);
    /// vùng ghi của phiên bản gốc (null = có thể ghi ô bất kỳ). Dòng ghi hiển thị cho một ô chỉ lấy phiên bản gốc có vùng giao ô đó.
    readonly Dictionary<long, SortedSet<long>> verOrig = new(), verOrigK = new();
    readonly Dictionary<long, Rgn> verRg = new();
    /// số ô có phiên bản mà không phiên bản gốc nào giao vùng ô (giữ mọi phiên bản gốc)
    public static int RgnFallback;
    /// chỗ đặt danh sách dòng ghi trong mô tả phiên bản hợp nhánh, điền khi gắn nhãn theo ô
    const string WPh = "\u0001W";
    long NewVer(int line, string desc, IEnumerable<long> from = null, string kind = "ghi", IEnumerable<string> calls = null, Rgn rgn = null)
    {
        var id = Mk.Next(); verInfo[id] = (line, desc);
        var w = new SortedSet<int>(); var cs = new SortedSet<string>(StringComparer.Ordinal); string k = kind;
        var og = new SortedSet<long>(); var ogk = new SortedSet<long>();
        if (from == null) { w.Add(line); if (calls != null) cs.UnionWith(calls); og.Add(id); ogk.Add(id); verRg[id] = rgn; }
        else
        {
            k = "ghi";
            var fl = from.ToList();
            foreach (var v in fl)
                if (v > 0 && verKind.TryGetValue(v, out var vk)) { if (vk.Kind == "co_the_ghi") k = "co_the_ghi"; cs.UnionWith(vk.Calls); ogk.UnionWith(verOrigK[v]); }
            foreach (var v in SrcVers(fl)) { w.UnionWith(verWrites[v]); og.UnionWith(verOrig[v]); }
        }
        verWrites[id] = w; verKind[id] = (k, cs); verOrig[id] = og; verOrigK[id] = ogk; return id;
    }
    readonly Dictionary<V, List<Effect>> pendingBtn = new(ReferenceEqualityComparer.Instance);
    /// vùng ô (sheet, dòng, cột 0-based); thành phần lúc chạy là bất kỳ
    /// RS/CS: phần ký hiệu của chỉ số hàng/cột tuyến tính (cùng phần ký hiệu, khác hằng RO/CO thì rời nhau)
    public sealed record Rgn(string S, long R1, long R2, long C1, long C2, string RS = null, long RO = 0, string CS = null, long CO = 0);
    readonly Dictionary<string, Rgn> rgnOf = new(StringComparer.Ordinal);
    static long? Lc(V v) => v is VConst k && k.Val is long l ? l : null;
    string RegionKey(VCell c, out bool exact)
    {
        string s = c.S.Index is VConst sk ? R.Expr(sk, new Flags()) : null;
        long? r1 = null, r2 = null, c1 = null, c2 = null;
        string rs = null, cs = null; long ro = 0, co = 0;
        static string Ls(V v, out long off) { off = 0; if (v is VLin l) { off = l.C; return string.Join(",", l.T.Select(kv => kv.Key + "*" + kv.Value.ToString(CultureInfo.InvariantCulture))); } return null; }
        switch (c.Kind)
        {
            case "cell": r1 = r2 = Lc(c.R1); c1 = c2 = Lc(c.C1); if (r1 == null) rs = Ls(c.R1, out ro); if (c1 == null) cs = Ls(c.C1, out co); break;
            case "range":
                r1 = Lc(c.R1); r2 = Lc(c.R2); c1 = Lc(c.C1); c2 = Lc(c.C2);
                if (r1 == null || r2 == null) r1 = r2 = null;
                if (c1 == null || c2 == null) c1 = c2 = null;
                break;
            case "row": r1 = r2 = Lc(c.R1); if (r1 == null) rs = Ls(c.R1, out ro); break;
            case "col": c1 = c2 = Lc(c.C1); if (c1 == null) cs = Ls(c.C1, out co); break;
        }
        exact = s != null && c.Kind == "cell" && r1 != null && c1 != null;
        static string P(long? x) => x == null ? "*" : x.Value.ToString(CultureInfo.InvariantCulture);
        static string Q(string sy, long o) => sy == null ? "" : "~" + sy + "@" + o.ToString(CultureInfo.InvariantCulture);
        var key = (s ?? "*") + "|" + P(r1) + "|" + P(r2) + Q(rs, ro) + "|" + P(c1) + "|" + P(c2) + Q(cs, co);
        if (!rgnOf.ContainsKey(key)) rgnOf[key] = new Rgn(s, r1 ?? long.MinValue, r2 ?? long.MaxValue, c1 ?? long.MinValue, c2 ?? long.MaxValue, rs, ro, cs, co);
        return key;
    }
    static bool Overlap(Rgn a, Rgn b) => (a.S == null || b.S == null || a.S == b.S) && a.R1 <= b.R2 && b.R1 <= a.R2 && a.C1 <= b.C2 && b.C1 <= a.C2
        && !(a.RS != null && a.RS == b.RS && a.RO != b.RO) && !(a.CS != null && a.CS == b.CS && a.CO != b.CO);
    void Wild(State st, int line, string desc, string kind, IEnumerable<string> calls) { st.WildV = NewVer(line, desc, null, kind, calls); }
    /// lệnh đổi giá trị ô: phiên bản mới chỉ cho vùng được ghi (ô hằng, hoặc cột/hàng/sheet khi chỉ số lúc chạy)
    void WriteCell(State st, VCell c, int line, string desc, string kind = "ghi", IEnumerable<string> calls = null)
    {
        var key = RegionKey(c, out var exact); var v = NewVer(line, desc, null, kind, calls, rgnOf[key]);
        if (exact) st.CV = st.CV.SetItem(key, v); else st.RV = st.RV.SetItem(key, v);
    }
    /// phiên bản hiệu lực của vùng trong một trạng thái: lớn nhất trong các lệnh ghi có vùng giao với vùng đọc
    long EffVer(State st, string key, bool exact)
    {
        var g = rgnOf[key]; long v = st.WildV;
        if (exact) v = Math.Max(v, st.CV.GetValueOrDefault(key));
        else foreach (var kv in st.CV) if (kv.Value > v && Overlap(g, rgnOf[kv.Key])) v = kv.Value;
        foreach (var kv in st.RV) if (kv.Value > v && Overlap(g, rgnOf[kv.Key])) v = kv.Value;
        return v;
    }
    long ReadVer(State st, VCell c) { var k = RegionKey(c, out var exact); return EffVer(st, k, exact); }
    /// nhãn phiên bản hiển thị #<dòng sự kiện tạo phiên bản>; cùng ô cùng dòng nhiều phiên bản thì #<dòng>.2, .3 theo thứ tự gặp
    string VerTag(VCell cell, string cellText, long ver)
    {
        if (ver <= 0) return "";
        if (!verLabel.TryGetValue(cellText, out var m)) verLabel[cellText] = m = new();
        if (m.TryGetValue(ver, out var lab)) return lab;
        (int Line, string Desc) vi = verInfo.TryGetValue(ver, out var v0) ? v0 : (0, "?");
        string b = "#" + vi.Line.ToString(CultureInfo.InvariantCulture); lab = b; int n = 1;
        while (m.ContainsValue(lab)) lab = b + "." + (++n).ToString(CultureInfo.InvariantCulture);
        m[ver] = lab;
        var rf = "cell(" + cellText + ")" + lab;
        (string Kind, SortedSet<string> Calls) vk = verKind.TryGetValue(ver, out var vk0) ? vk0 : ("co_the_ghi", new SortedSet<string>(StringComparer.Ordinal));
        int[] wr = verWrites.TryGetValue(ver, out var vw) ? vw.ToArray() : Array.Empty<int>();
        string desc = vi.Desc;
        if (verOrig.TryGetValue(ver, out var og))
        {
            // dòng ghi, loại, lời gọi theo vùng ô đang đọc: chỉ phiên bản gốc có vùng ghi giao ô này
            var g = rgnOf[RegionKey(cell, out _)];
            bool Hit(long o) => verRg.GetValueOrDefault(o) is not Rgn r || Overlap(r, g);
            var ogf = og.Where(Hit).ToList();
            if (ogf.Count == 0 && og.Count > 0) { RgnFallback++; ogf = og.ToList(); }
            wr = ogf.SelectMany(o => verWrites[o]).Distinct().OrderBy(x => x).ToArray();
            var okf = verOrigK[ver].Where(Hit).ToList(); if (okf.Count == 0) okf = verOrigK[ver].ToList();
            if (okf.Count > 0)
            {
                var cs2 = new SortedSet<string>(okf.SelectMany(o => verKind[o].Calls), StringComparer.Ordinal);
                vk = (okf.Any(o => verKind[o].Kind == "co_the_ghi") ? "co_the_ghi" : "ghi", cs2);
            }
            desc = desc.Replace(WPh, string.Join(", ", wr));
        }
        VerRefs[rf] = desc;
        VerMeta[rf] = (vk.Kind, wr, vk.Calls.ToArray());
        return lab;
    }
    string WriteLines(IEnumerable<long> vs) => string.Join(", ", SrcVers(vs).SelectMany(v => verWrites[v]).Distinct().OrderBy(x => x));
    /// hợp phiên bản ô của các nhánh còn sống: khác nhau thì tạo phiên bản mới tại dòng hợp nhánh.
    /// Thứ tự: lời gọi có thể ghi (WildV), vùng (RV), ô hằng (CV); mỗi khóa so phiên bản hiệu lực của từng nhánh (kể cả vùng phủ nó)
    void MergeVers(State st, List<State> live, int line)
    {
        if (live.Count == 0) return;
        var wv = live.Select(s => s.WildV).ToList();
        long wild = wv.All(v => v == wv[0]) ? wv[0] : NewVer(line, "hợp nhánh dòng " + line + ": trên một số nhánh có lời gọi có thể ghi ô (dòng " + WPh + ")", wv);
        // mục vùng được ghi trong các nhánh (giá trị khác nhau giữa các nhánh): mỗi mục là một lệnh ghi cuối có thể có cho phần ô
        // giao với nó, nên phiên bản hợp nhánh của một khóa lấy nguồn từ mọi mục như vậy giao khóa (không chỉ phiên bản lớn nhất)
        var div = new List<(string K, long V)>();
        foreach (var s in live)
            foreach (var kv in s.RV.Concat(s.CV))
                if (!live.All(o => (o.RV.TryGetValue(kv.Key, out var x) || o.CV.TryGetValue(kv.Key, out x)) && x == kv.Value)) div.Add((kv.Key, kv.Value));
        ImmutableDictionary<string, long> MergeMap(Func<State, ImmutableDictionary<string, long>> sel, bool exact)
        {
            var keys = new SortedSet<string>(live.SelectMany(s => sel(s).Keys), StringComparer.Ordinal);
            var bld = ImmutableDictionary.CreateBuilder<string, long>(StringComparer.Ordinal);
            foreach (var k in keys)
            {
                var raw = live.Select(s => sel(s).GetValueOrDefault(k)).ToList();
                if (raw.All(v => v == raw[0]) && live.All(s => s.WildV == live[0].WildV) && live.All(s => s.RV == live[0].RV)) { bld[k] = raw[0]; continue; }
                var vs = live.Select(s => EffVer(s, k, exact)).ToList();
                if (!vs.All(v => v == vs[0])) { var gk = rgnOf[k]; foreach (var (dk, dv) in div) if (Overlap(gk, rgnOf[dk])) vs.Add(dv); }
                bld[k] = vs.Take(live.Count).All(v => v == vs[0]) ? vs[0] : NewVer(line, "hợp nhánh dòng " + line + ": ô được ghi trên một số nhánh (dòng " + WPh + ")", vs);
            }
            return bld.ToImmutable();
        }
        var rv = MergeMap(s => s.RV, false);
        var cv = MergeMap(s => s.CV, true);
        st.WildV = wild; st.RV = rv; st.CV = cv;
    }
    void AddEffect(Effect e, State st)
    {
        Effects.Add(e);
        string d = e.Op + " dòng " + e.Line + (e.Via != null ? " qua " + e.Via : "");
        // chỉ lệnh đổi giá trị ô tạo phiên bản; khóa/mở, định dạng, Visible/Height/Width, ghi chú, màu thẻ, chữ nút: không
        if (e.Op is "setValue" or "setFormula")
        {
            if (e.Target is VCell tc) WriteCell(st, tc, e.Line, "sau lệnh " + d);
            else Wild(st, e.Line, "sau lệnh " + d + " (ô đích chưa xác định)", "ghi", null);
        }
        else if (e.Op == "opaque" && e.Params.TryGetValue("touchesSheet", out var ts) && ts is true && e.Params.TryGetValue("call", out var cu) && cu is string cus && CallWriteKind(cus) is string wk)
            Wild(st, e.Line, wk == "ghi" ? "sau lệnh " + cus + " dòng " + e.Line + (e.Via != null ? " qua " + e.Via : "") + " (đổi giá trị hoặc dịch ô)" : "sau lời gọi chưa giải " + cus + " dòng " + e.Line + " (có thể ghi ô)", wk, new[] { cus });
        e.Params.Remove("touchesSheet");
    }
    /// nguyên nhân điều kiện hằng dựa trên biến kết quả hộp thoại khi hộp thoại không hiện (nhánh gán kết quả bị bỏ vì điều kiện hằng)
    string DialogCause(ExpressionSyntax cond, State st, Frame f)
    {
        if (f.DlgSkip == null) return null;
        foreach (var id in cond.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>())
            if (f.DlgSkip.TryGetValue(id.Identifier.Text, out var cause))
            {
                var iv = st.Env.TryGetValue(id.Identifier.Text, out var en) && en.Val is VConst ik ? R.Expr(ik, new Flags()) : null;
                return cause + (iv != null ? ", giá trị khởi tạo " + iv : "");
            }
        return null;
    }
    void NoteSkippedDialog(StatementSyntax skipped, ExpressionSyntax cond, State st, Frame f)
    {
        if (skipped == null) return;
        foreach (var a in skipped.DescendantNodesAndSelf().OfType<AssignmentExpressionSyntax>())
        {
            if (a.Left is not IdentifierNameSyntax lid) continue;
            bool dlg = a.Right.DescendantNodesAndSelf().OfType<InvocationExpressionSyntax>().Any(iv => iv.Expression is MemberAccessExpressionSyntax m && (m.Name.Identifier.Text == "ShowDialog" || (m.Name.Identifier.Text == "Show" && m.Expression.ToString() == "MessageBox")));
            if (!dlg) continue;
            var parts = cond.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>().Select(i => i.Identifier.Text).Distinct()
                .Where(n => st.Env.TryGetValue(n, out var en) && en.Val is VConst).Select(n => n + " = " + R.Expr(st.Env[n].Val, new Flags())).ToList();
            f.DlgSkip ??= new Dictionary<string, string>(StringComparer.Ordinal);
            f.DlgSkip[lid.Identifier.Text] = "hộp thoại không hiện vì " + (parts.Count > 0 ? string.Join(", ", parts) : "điều kiện dòng " + Line(a) + " luôn sai");
        }
    }
    /// method cùng class có ghi ô (do Program cấp), dùng khi ghi dropped cho nhánh hằng
    public Func<string, bool> IsWriterMethod;
    public readonly List<(int Line, string Via, string Reason)> ConstDrops = new();
    void DropConst(SyntaxNode s, Frame f, string why)
    {
        if (s == null) return;
        foreach (var n in s.DescendantNodesAndSelf())
        {
            int l0 = 0;
            if (n is AssignmentExpressionSyntax a && a.IsKind(SyntaxKind.SimpleAssignmentExpression) && a.Left is MemberAccessExpressionSyntax lm
                && ((lm.Expression is ElementAccessExpressionSyntax ea && ea.Expression is MemberAccessExpressionSyntax em && em.Name.Identifier.Text is "Cells" or "Rows" or "Columns" or "TabStrip")
                    || (lm.Name.Identifier.Text == "Text" && lm.Expression is IdentifierNameSyntax bi && bi.Identifier.Text.Contains("ButtonCellType", StringComparison.OrdinalIgnoreCase)))) l0 = Line(a);
            else if (n is InvocationExpressionSyntax iv)
            {
                string nm = iv.Expression switch { IdentifierNameSyntax i => i.Identifier.Text, MemberAccessExpressionSyntax m => m.Name.Identifier.Text, _ => "" };
                bool local = iv.Expression is IdentifierNameSyntax || (iv.Expression is MemberAccessExpressionSyntax tm && tm.Expression is ThisExpressionSyntax);
                if (nm is "checkLockUnLockCells" or "ChangeValueCell" or "LockUnLockCol" or "LockUnLockCols" or "changFormatCell"
                    || (nm.StartsWith("Reset", StringComparison.Ordinal) && iv.Expression is MemberAccessExpressionSyntax rm && rm.Expression is ElementAccessExpressionSyntax)
                    || (local && IsWriterMethod != null && IsWriterMethod(nm))) l0 = Line(iv);
            }
            if (l0 == 0) continue;
            string reason = "nhánh hằng sai: " + why;
            Dropped.Add((l0, reason));
            if (f.Depth > 0) ConstDrops.Add((l0, f.Via, reason));
        }
    }
    public readonly Dictionary<long, SymInfo> SymById = new();
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
        var e = l.FirstOrDefault(x => x.Origin == origin && (x.Line == line || origin == "unknown"));
        if (e != null) return e;
        e = new SymInfo { Name = name, Origin = origin, Line = line, Desc = desc, Id = Mk.Next() };
        SymById[e.Id] = e;
        l.Add(e); return e;
    }
    bool pendingThrow; string pendingWhat;
    void CheckThrow(State st, int line, string what)
    {
        if (!pendingThrow) return;
        pendingThrow = false; Dropped.Add((line, "ngoại lệ chắc chắn: " + (pendingWhat ?? what))); pendingWhat = null; st.Dead = true; st.Threw = true;
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
            V pv = kv.Value;
            if ((kv.Value is VSym ps && ps.Text == kv.Key) || (kv.Value is VLin pl && pl.IsSingleAtom(out var pa) && pa == kv.Key) || (kv.Value is VSheet psh && psh.Index is VSym pss && pss.Text == kv.Key))
            {
                var si = RegSym(kv.Key, "param", Line(m), "tham số của method " + MethodName(m) + ", giá trị do nơi gọi truyền vào");
                pv = kv.Value is VSym ? new VSym(si.Ref) : kv.Value is VLin ? VLin.Atom(si.Ref) : new VSheet(new VSym(si.Ref));
            }
            else if (kv.Value is VSym pvs && pvs.Text is "tk" or "ttc" or "period" or "header")
            {
                // tên quy ước cho tham số thông tin tờ khai của hook (tk, ttc, period, header)
                var pt = m.ParameterList.Parameters.FirstOrDefault(x => x.Identifier.Text == kv.Key)?.Type?.ToString() ?? "?";
                pv = new VSym(RegSym(pvs.Text, "hook-param", Line(m), "tham số " + kv.Key + " (" + pt + ") của method " + MethodName(m)).Ref);
            }
            st.Env[kv.Key] = new Entry(pv, ImmutableSortedSet<int>.Empty);
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
                    Assign(st, v.Identifier.Text, NamedCollection(v.Identifier.Text, val, Line(v)), Line(v));
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
                    if (c is CConst k)
                    {
                        if (k.B) es2.Dead = true; else ts.Dead = true; ts.Path = st.Path; es2.Path = st.Path;
                        var skipped = k.B ? ifs.Else?.Statement : ifs.Statement;
                        string why = "điều kiện dòng " + ln + " luôn " + (k.B ? "đúng, nhánh else không chạy" : "sai");
                        if (DialogCause(ifs.Condition, st, f) is string dcause) why += " (" + dcause + ")";
                        NoteSkippedDialog(skipped, ifs.Condition, st, f);
                        DropConst(skipped, f, why);
                    }
                    if (!ts.Dead) Exec(ifs.Statement, ts, f);
                    if (!es2.Dead && ifs.Else != null) Exec(ifs.Else.Statement, es2, f);
                    Merge(st, new List<(Cond, State)> { (c, ts), (C.Not(c), es2) }, ln);
                    break;
                }
            case ReturnStatementSyntax rs:
                if (rs.Expression != null) { used = new(); var v = Eval(rs.Expression, st, f); f.Returns.Add((st.Path, v)); }
                if (!st.Dead) f.Exits?.Add(new State { CV = st.CV, RV = st.RV, WildV = st.WildV });
                st.Dead = true; break;
            case ThrowStatementSyntax: st.Dead = true; st.Threw = true; break;
            case TryStatementSyntax ts2:
                Exec(ts2.Block, st, f);
                if (st.Threw) { st.Dead = false; st.Threw = false; }
                // finally chạy cả khi khối try kết thúc bằng return/break/continue; sau finally đường đó vẫn đã thoát
                if (ts2.Finally != null) { bool wd = st.Dead; st.Dead = false; Exec(ts2.Finally.Block, st, f); if (wd) st.Dead = true; }
                break;
            case BreakStatementSyntax:
                if (jumps.Count > 0) { jumps[^1].Brk.Add(st.Clone()); st.Dead = true; }
                break;
            case ContinueStatementSyntax:
                for (int i = jumps.Count - 1; i >= 0; i--)
                    if (!jumps[i].Sw) { jumps[i].Cont.Add(st.Clone()); st.Dead = true; break; }
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
                    var atoms = EnterLoop(st, assigned, Line(fs), fs, f);
                    if (fs.Declaration != null) foreach (var v in fs.Declaration.Variables) { var si = RegSym(v.Identifier.Text, "loop", Line(fs), "biến lặp của vòng lặp dòng " + Line(fs)); atoms[v.Identifier.Text] = si.Ref; st.Env[v.Identifier.Text] = new Entry(VLin.Atom(si.Ref), ImmutableSortedSet.Create(Line(v))); }
                    used = new();
                    string cond = fs.Condition != null ? R.Cond(C.Of(Eval(fs.Condition, st, f)), new Flags()) : "true";
                    Loop(st, f, fs.Statement, "for", "while " + cond, Line(fs), assigned, init0, fs.Incrementors, declInit.Keys.ToList(), atoms);
                    break;
                }
            case ForEachStatementSyntax fe:
                {
                    used = new();
                    var coll = Eval(fe.Expression, st, f);
                    var assigned = AssignedIn(fe);
                    var init0 = Snap(st, assigned);
                    var atoms = EnterLoop(st, assigned, Line(fe), fe, f);
                    string ct = R.Expr(coll, new Flags());
                    var fsi = RegSym(fe.Identifier.Text, "foreach", Line(fe), "phần tử đang duyệt của " + ct);
                    st.Env[fe.Identifier.Text] = new Entry(new VSym(fsi.Ref), ImmutableSortedSet.Create(Line(fe)));
                    Loop(st, f, fe.Statement, "foreach " + fsi.Ref + " in " + ct, null, Line(fe), assigned, init0, null, null, atoms);
                    break;
                }
            case WhileStatementSyntax ws:
                {
                    var assigned = AssignedIn(ws);
                    var init0 = Snap(st, assigned);
                    var atoms = EnterLoop(st, assigned, Line(ws), ws, f);
                    used = new();
                    Loop(st, f, ws.Statement, "while", R.Cond(C.Of(Eval(ws.Condition, st, f)), new Flags()), Line(ws), assigned, init0, null, null, atoms);
                    break;
                }
            case DoStatementSyntax ds:
                {
                    var assigned = AssignedIn(ds);
                    var init0 = Snap(st, assigned);
                    var atoms = EnterLoop(st, assigned, Line(ds), ds, f);
                    used = new();
                    Loop(st, f, ds.Statement, "do", "while " + R.Cond(C.Of(Eval(ds.Condition, st, f)), new Flags()), Line(ds), assigned, init0, null, null, atoms);
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
                        if (c is CConst k0 && !k0.B) { bs.Dead = true; foreach (var x in sec.Statements) DropConst(x, f, "nhãn case dòng " + Line(sec) + " không khớp giá trị switch dòng " + ln); }
                        bs = SwitchSection(sec, bs, f);
                        branches.Add((c, bs));
                    }
                    var dc = C.Not(C.Or(labels));
                    var ds2 = st.Clone(); ds2.Path = ds2.Path.Add(new PathItem(dc, def != null ? Line(def) : ln));
                    if (dc is CConst k1 && !k1.B) { ds2.Dead = true; if (def != null) foreach (var x in def.Statements) DropConst(x, f, "nhánh default dòng " + Line(def) + " không chạy vì giá trị switch dòng " + ln + " luôn khớp một nhãn case"); }
                    if (def != null) ds2 = SwitchSection(def, ds2, f);
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

    /// một nhánh switch: break thoát nhánh; trạng thái sau nhánh là hợp các đường tới break và đường chạy hết nhánh
    State SwitchSection(SwitchSectionSyntax sec, State bs, Frame f)
    {
        if (bs.Dead) return bs;
        var jc = new JumpCtx { Sw = true }; jumps.Add(jc);
        foreach (var x in sec.Statements) { if (bs.Dead) break; Exec(x, bs, f); }
        jumps.RemoveAt(jumps.Count - 1);
        if (jc.Brk.Count == 0) return bs;
        var all = new List<State>(jc.Brk); if (!bs.Dead) all.Add(bs);
        var t = bs.Clone(); Join(t, all, Line(sec)); return t;
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
        var brk = new List<State>();
        foreach (var v in vals)
        {
            if (st.Dead) break;
            st.Env[name] = new Entry(v, ImmutableSortedSet.Create(Line(fs)));
            var jc = new JumpCtx(); jumps.Add(jc);
            Exec(fs.Statement, st, f);
            jumps.RemoveAt(jumps.Count - 1);
            if (jc.Cont.Count > 0)
            {
                // continue: vòng sau chạy từ hợp các đường tới continue và đường chạy hết thân
                var all = new List<State>(jc.Cont); if (!st.Dead) all.Add(st);
                var t = st.Clone(); Join(t, all, Line(fs));
                st.Env = t.Env; st.CopyVers(t); st.Path = t.Path; st.Dead = t.Dead; st.Threw = t.Threw;
            }
            brk.AddRange(jc.Brk);
        }
        if (brk.Count > 0)
        {
            var all = new List<State>(brk); if (!st.Dead) all.Add(st);
            var t = st.Clone(); Join(t, all, Line(fs));
            st.Env = t.Env; st.CopyVers(t); st.Path = t.Path; st.Dead = t.Dead; st.Threw = t.Threw;
        }
        return true;
    }

    static Dictionary<string, V> Snap(State st, HashSet<string> assigned)
    {
        var d = new Dictionary<string, V>(StringComparer.Ordinal);
        foreach (var n in assigned) if (st.Env.TryGetValue(n, out var e)) d[n] = e.Val;
        return d;
    }

    Dictionary<string, string> EnterLoop(State st, HashSet<string> assigned, int line, SyntaxNode loopNode, Frame f)
    {
        long pre = st.WildV;
        if (loopNode != null)
        {
            if (MayWrite(loopNode)) fieldVer = Mk.Next();
            // phiên bản đầu vòng: chỉ khi thân vòng có lệnh đổi giá trị ô hoặc lời gọi có thể ghi; chỉ có hiệu lực trong thân vòng (xem Loop)
            var lw = LoopWrites(loopNode, f?.Cls);
            if (lw.Kind != null) { Wild(st, line, "vòng lặp dòng " + line + (lw.Kind == "ghi" ? " có lệnh ghi giá trị ô" : " có lời gọi có thể ghi ô (" + string.Join(", ", lw.Calls) + ")"), lw.Kind, lw.Calls); activeLoop.Add(st.WildV); }
        }
        loopVers.Push((pre, st.WildV));
        var atoms = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var n in assigned.OrderBy(x => x, StringComparer.Ordinal))
        {
            var si = RegSym(n, "loop", line, "giá trị đầu mỗi vòng của vòng lặp dòng " + line);
            atoms[n] = si.Ref;
            // biến chuỗi/đối tượng mang qua vòng: phép chuyển số trên nó không được bỏ (int(text) khác text)
            if (loopNode != null && loopNode.Ancestors().OfType<BaseMethodDeclarationSyntax>().FirstOrDefault() is { } md
                && md.DescendantNodes().OfType<VariableDeclaratorSyntax>().FirstOrDefault(d => d.Identifier.Text == n)?.Parent is VariableDeclarationSyntax vd
                && vd.Type.ToString() is "string" or "object" or "String" or "Object") strAtoms.Add(si.Ref);
            st.Env[n] = new Entry(VLin.Atom(si.Ref), ImmutableSortedSet<int>.Empty);
        }
        return atoms;
    }

    /// head: "for"/"while"/"do"/"foreach ..."; cond: điều kiện lặp. Mô tả cuối: head + biến cảm ứng (from, step) + điều kiện.
    void Loop(State st, Frame f, StatementSyntax body, string head, string cond, int line, HashSet<string> assigned, Dictionary<string, V> init0, SeparatedSyntaxList<ExpressionSyntax>? incs, List<string> declared, Dictionary<string, string> atoms)
    {
        string A(string n) => atoms != null && atoms.TryGetValue(n, out var a) ? a : n;
        var li = new LoopInfo(head + (cond != null ? " " + cond : ""), line);
        var lp = loopVers.Count > 0 ? loopVers.Pop() : (st.WildV, st.WildV);
        var bs = st.Clone(); bs.Loops = bs.Loops.Add(li);
        var b0 = bs.Clone(); var jc = new JumpCtx(); jumps.Add(jc);
        Exec(body, bs, f);
        jumps.RemoveAt(jumps.Count - 1);
        // continue: cuối thân vòng là hợp các đường tới continue và đường chạy hết thân
        if (jc.Cont.Count > 0) { var all = new List<State>(jc.Cont); if (!bs.Dead) all.Add(bs); Join(b0, all, line); bs = b0; }
        if (incs != null) { bs.Dead = false; foreach (var inc in incs.Value) { used = new(); ExecExpr(inc, bs, f); } }
        // vòng lặp chạy 0 hoặc nhiều lần: hợp phiên bản ô trước vòng và sau thân vòng; phiên bản đầu vòng chỉ dùng trong thân vòng,
        // sau vòng ô giữ phiên bản trước vòng hoặc phiên bản của lệnh ghi trong thân (theo vùng ghi)
        var pre0 = st.Clone();
        if (pre0.WildV == lp.Item2) pre0.WildV = lp.Item1;
        if (bs.WildV == lp.Item2) bs.WildV = lp.Item1;
        var outs = new List<State> { pre0, bs };
        // break: phiên bản ô tại lệnh break cũng ra sau vòng
        foreach (var b in jc.Brk) { var bc = b.Clone(); if (bc.WildV == lp.Item2) bc.WildV = lp.Item1; outs.Add(bc); }
        MergeVers(st, outs, line);
        activeLoop.Remove(lp.Item2);
        // biến cảm ứng: giá trị đầu và bước hằng
        var ind = new List<string>();
        foreach (var n in assigned.OrderBy(x => (declared != null && declared.Contains(x)) ? 0 : 1).ThenBy(x => x, StringComparer.Ordinal))
        {
            V init = init0.GetValueOrDefault(n);
            string it = init != null ? R.Expr(init, new Flags()) : "?";
            if (!bs.Env.TryGetValue(n, out var be)) continue;
            var step = StepOf(A(n), be.Val);
            string bodyTxt = R.Expr(be.Val, new Flags());
            if (step is VConst sc && sc.Val is long sl && sl != 0) ind.Add(A(n) + " from " + it + " step " + (sl > 0 ? "+" : "") + sl);
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
            if (mutated.Contains(n)) { var sc2 = RegSym(n, "loop-collection", line, "tập hợp được thêm/bớt phần tử trong vòng lặp dòng " + line + ", nội dung lúc chạy"); st.Env[n] = new Entry(new VSym(sc2.Ref + Mk.Tag(Mk.Next())), ImmutableSortedSet.Create(line)); continue; }
            if (init != null && bs.Env.TryGetValue(n, out var be)) fin = Accum(A(n), init, be.Val, desc);
            string al = null;
            if (fin == null) al = RegSym("after_loop", "after_loop", line, "giá trị sau vòng lặp dòng " + line + " khi không xác định được giá trị trước vòng").Ref;
            st.Env[n] = new Entry(fin ?? new VSym(al + "(" + A(n) + ", " + desc + ")"), ImmutableSortedSet.Create(line));
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
            st.Env = live[0].S.Env; st.CopyVers(live[0].S);
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
            if (v is VOpaque) { var sm = RegSym(k.StartsWith("this.", StringComparison.Ordinal) ? k.Substring(5) : k, "merge", line, "giá trị phụ thuộc nhiều hơn " + Ar.MaxAlts + " nhánh rẽ tại dòng " + line); v = VLin.Atom("opaque(branches, " + sm.Ref + ")" + Mk.Tag(Mk.Next())); }
            env[k] = new Entry(v, lines.Add(line));
        }
        st.Env = env;
        MergeVers(st, live.Select(b => b.S).ToList(), line);
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
                    Assign(st, n, n.StartsWith("this.", StringComparison.Ordinal) ? rhs : NamedCollection(n, rhs, ln), ln); return rhs;
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
                        AddEffect(eb, st); pl.Add(eb);
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
                        AddEffect(et, st);
                        return rhs;
                    }
                    if (prop == "Text" && ma.Expression is IdentifierNameSyntax bid && bid.Identifier.Text.Contains("ButtonCellType", StringComparison.OrdinalIgnoreCase))
                    {
                        var eb = NewEffect(st, f, ln); eb.Op = "setButtonText"; eb.Target = null; eb.Val = rhs;
                        eb.Params["text"] = R.Expr(rhs, new Flags()); eb.Params["note"] = "kiểu ô nút lấy lúc chạy, ô đích chưa xác định";
                        AddEffect(eb, st);
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
                            AddEffect(ef, st);
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
                    else { var sa = RegSym(aid.Identifier.Text, "array", ln, "mảng sau khi gán phần tử theo chỉ số lúc chạy tại dòng " + ln); Assign(st, aid.Identifier.Text, VLin.Atom(sa.Ref), ln); }
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
        if (f.Cls != null && f.Cls.Members.Contains(n)) return Field(n, f.Cls, st);
        if (at != null) return new VSym(RegSym(n, "unknown", Line(at), "tên không có giá trị trong phạm vi phân tích (khai báo ở phạm vi khác hoặc kiểu tĩnh)").Ref);
        return new VSym(n);
    }

    public V Field(string n, ClassCtx cls, State vs = null)
    {
        if (cls.FieldType.TryGetValue(n, out var ft)) { if (ft == "FpSpread") return VSpread.I; if (ft == "SheetView") return new VSheet(new VSym("field(" + n + ")")); }
        if (cls.Mutable.Contains(n)) return new VSym("field(" + n + ")" + (fieldVer > 0 ? Mk.Tag(fieldVer) : ""));
        if (cls.Cache.TryGetValue(n, out var v)) return v;
        if (!cls.Evaluating.Add(n)) return new VSym("field(" + n + ")");
        var saved = used; used = new();
        var st = new State { Env = new(StringComparer.Ordinal), Path = ImmutableList<PathItem>.Empty, Loops = ImmutableList<LoopInfo>.Empty };
        if (vs != null) st.CopyVers(vs);
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
        // ký hiệu có tên trong dấu ẩn đăng ký ở Interp hiện tại: không lưu đệm để Interp sau đăng ký lại symbols
        var ftxt = R.Expr(v, new Flags());
        if (!ftxt.Contains("cell(") && !Mk.Names(ftxt).Any() && !Mk.HasSym(ftxt)) cls.Cache[n] = v;
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
                    var ocArgs = oc.ArgumentList != null ? oc.ArgumentList.Arguments.Select(a => Eval(a.Expression, st, f)).ToList() : new List<V>();
                    var tn = oc.Type.ToString(); int lt = tn.IndexOf('<'); if (lt > 0) tn = tn.Substring(0, lt);
                    if ((tn == "List" || tn == "HashSet") && oc.Initializer == null && (oc.ArgumentList == null || oc.ArgumentList.Arguments.Count == 0)) return new VArr(new List<V>());
                    tn = tn.Split('.').Last();
                    if (tn.EndsWith("CellType", StringComparison.Ordinal))
                    {
                        var props = oc.Initializer?.Expressions.OfType<AssignmentExpressionSyntax>().Select(a => a.Left.ToString().Trim() + "=" + R.Expr(Eval(a.Right, st, f), new Flags())) ?? Enumerable.Empty<string>();
                        return new VSym("celltype(" + string.Join(", ", new[] { tn.Substring(0, tn.Length - 8) }.Concat(props)) + ")");
                    }
                    if (tn.StartsWith("frm", StringComparison.Ordinal))
                    {
                        // biểu mẫu hộp thoại: mỗi lần tạo là một đối tượng riêng; Yes/No/DialogResult đọc sau ShowDialog là lựa chọn của người dùng
                        var t0 = "obj(" + tn + ")"; var tf = t0 + Mk.Tag(Mk.Next());
                        dialogForms[tf] = (Line(oc), tn, ocArgs.Count > 0 ? StrConst(ocArgs[0]) : null, t0.Length);
                        return new VSym(tf);
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
            case LambdaExpressionSyntax le: return new VLambda(le, LamText(le, st, f, 0));
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
        return Member(b, n, st);
    }

    public V Member(V b, string n, State st)
    {
        // số phần tử / độ dài viết dạng hàm trung tính count(x) / length(x)
        if (n is "Count" or "Length" && b is not VArr && b is not VCases && !(b is VConst kc0 && (kc0.Val is string || kc0.Val == null)))
        {
            var fc = new Flags(); var t0 = R.Expr(b, fc);
            return new VSym((n == "Count" ? "count(" : "length(") + t0 + ")", fc.Opaque);
        }
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
                    long ver = ReadVer(st, c);
                    return new VSym("cell(" + pre + rf + ")" + VerTag(c, pre + rf, ver) + "." + n.ToLowerInvariant(), fl.Opaque);
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
            case VConst kn0 when kn0.Val == null:
                // truy cập thành viên trên null: ngoại lệ chắc chắn ở câu lệnh này
                pendingThrow = true; pendingWhat = "truy cập thành viên " + n + " trên giá trị null";
                return new VOpaque("null-deref");
            case VCases cs:
                {
                    // nhánh có giá trị null không đi tiếp được (ngoại lệ): bỏ nhánh đó, giữ các nhánh còn lại
                    var alts = cs.Alts.Where(x => !(x.Val is VConst kz && kz.Val == null)).ToList();
                    if (alts.Count == 0) alts = cs.Alts;
                    if (alts.Count < cs.Alts.Count && alts.Count > 0 && alts[^1].When != null) alts[^1] = (null, alts[^1].Val);
                    return Ar.Cases(alts.Select(x => (x.When, Member(x.Val, n, st))).ToList());
                }
            case VSym y when (n == "Yes" || n == "No" || n == "DialogResult") && dialogForms.TryGetValue(y.Text, out var dfm):
                {
                    var si = RegSym("user_choice", "dialog", dfm.Line, "nút người dùng bấm trên biểu mẫu hộp thoại" + (dfm.Code != null ? " mã \"" + dfm.Code + "\"" : "") + " tạo tại dòng " + dfm.Line + ", chọn lúc chạy, .yes đúng khi người dùng chọn Có, .no khi chọn Không" + (n == "DialogResult" ? ", .result theo mã DialogResult: 1 OK, 2 Cancel, 6 Yes, 7 No" : ""));
                    var txt = si.Ref + "(form" + (dfm.Code != null ? " \"" + dfm.Code + "\"" : "") + ", dòng " + dfm.Line + ")." + (n == "Yes" ? "yes" : n == "No" ? "no" : "result") + y.Text.Substring(dfm.Len);
                    return n == "DialogResult" ? VLin.Atom(txt) : new VSym(txt);
                }
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
            hint = RegSym(hint, "array", Line(ea), "mảng cục bộ " + ha.Items.Count + " phần tử, truy cập theo chỉ số lúc chạy").Ref;
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
            case VConst kn1 when kn1.Val == null:
                pendingThrow = true; pendingWhat = "truy cập chỉ số trên giá trị null";
                return new VOpaque("null-deref");
            case VCases cs:
                {
                    // nhánh null không đi tiếp được (ngoại lệ): bỏ, giữ các nhánh còn lại; ngoại lệ chắc chắn chỉ khi mọi nhánh đều ngoài mảng
                    var nn = cs.Alts.Where(x => !(x.Val is VConst kz && kz.Val == null)).ToList();
                    if (nn.Count > 0 && nn.Count < cs.Alts.Count) { if (nn[^1].When != null) nn[^1] = (null, nn[^1].Val); cs = new VCases(nn); }
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
                if (on != null) { var so = RegSym(on, "out", Line(inv), "nhận qua đối số out/ref của lời gọi " + CallName(inv) + " tại dòng " + Line(inv)); Assign(st, on, new VSym(so.Ref + Mk.Tag(Mk.Next())), Line(inv)); }
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
                Wild(st, ln, "lời gọi " + name + " không nội tuyến dòng " + ln + " (có thể ghi ô)", "co_the_ghi", new[] { name }); fieldVer = Mk.Next();
                var smt = RegSym(name, "method", ln, "method cùng class không nội tuyến tại dòng " + ln + ": " + crec.Reason);
                crec.EffAfter = Effects.Count; crec.DropAfter = Dropped.Count;
                var cargs = new List<V> { new VSym(smt.Ref) };
                foreach (var a in args) cargs.Add(Eval(a.Expression, st, f));
                if (isStmt) { var eo = NewEffect(st, f, ln); eo.Op = "opaque"; eo.Target = null; eo.Params["call"] = name; eo.Params["reason"] = crec.Reason; AddEffect(eo, st); }
                Opaque(name, inv, isStmt ? "statement" : "value");
                crec.EffAfter = Effects.Count;
                return new VOpaque("call", cargs);
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
                sst.CopyVers(st); sub.Exits = new List<State>();
                foreach (var kv in bind) sst.Env[kv.Key] = new Entry(kv.Value, ImmutableSortedSet.CreateRange(savedUsed).Add(ln));
                // trường lớp đã gán trong hook hiện tại được truyền xuống
                foreach (var kv in st.Env) if (kv.Key.StartsWith("this.", StringComparison.Ordinal)) sst.Env[kv.Key] = kv.Value;
                sub.BasePathLen = st.Path.Count;
                if (m.Body != null) Exec(m.Body, sst, sub);
                else if (m.ExpressionBody != null) { used = new(); sub.Returns.Add((sst.Path, Eval(m.ExpressionBody.Expression, sst, sub))); }
                used = savedUsed;
                crec.EffAfter = Effects.Count; crec.DropAfter = Dropped.Count;
                // phiên bản ô sau lời gọi: hợp các điểm ra (return và cuối thân)
                var exits = new List<State>(sub.Exits); if (!sst.Dead || exits.Count == 0) exits.Add(sst);
                MergeVers(st, exits, ln);
                foreach (var kv in sst.Env) if (kv.Key.StartsWith("this.", StringComparison.Ordinal)) st.Env[kv.Key] = kv.Value;
                // method có kiểu trả về nhưng không ghi nhận được lệnh return nào (ngoại lệ chắc chắn, vòng lặp): không giả định null
                if (sub.Returns.Count == 0)
                {
                    if (m is MethodDeclarationSyntax rmd && rmd.ReturnType.ToString() != "void") { var sr = RegSym(name, "method-result", ln, "giá trị trả về của method cùng class gọi tại dòng " + ln + " không xác định (không điểm kết thúc có giá trị nào tới được)"); return new VOpaque("result", new List<V> { new VSym(sr.Ref) }); }
                    return VConst.Null;
                }
                if (sub.Returns.Count == 1) return sub.Returns[0].Val;
                var alts = sub.Returns.Select(r => ((Cond)C.And(r.Path.Skip(st.Path.Count).Select(p => p.C)), r.Val)).ToList();
                var cv = Ar.Cases(alts.Select((a, i) => (i == alts.Count - 1 ? null : a.Item1, a.Item2)).ToList());
                if (cv is VOpaque) { var sr2 = RegSym(name, "method-result", ln, "giá trị trả về của method cùng class gọi tại dòng " + ln + " có quá " + Ar.MaxAlts + " nhánh"); return new VOpaque("result", new List<V> { new VSym(sr2.Ref) }); }
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
                    if (Lock != null && Lock.Error == null && Lock.Rows.Count > 0) { CheckLockFromConfig(inv, st, f, ln); return VConst.Null; }
                    // GeneralFunctions.checkLockUnLockCells: khóa/mở ô theo bảng cấu hình ở sheet cuối (đã đọc mã nguồn hàm)
                    var e1 = NewEffect(st, f, ln); e1.Op = "lock"; e1.Target = null;
                    e1.Params["targetReason"] = Lock == null ? "class không có mục trong lock_config.json" : Lock.Error != null ? "không đọc được mẫu: " + Lock.Error : "mẫu " + Lock.Template + " không có dòng cấu hình (" + (Lock.Stop ?? "sheet cuối trống") + "): hàm không khóa/mở ô nào";
                    e1.Params["via"] = "checkLockUnLockCells";
                    e1.Params["enabled"] = "ô BA1 của sheet cuối == \"1\"";
                    e1.Params["targets"] = "ô đích theo bảng cấu hình ở sheet cuối, mỗi dòng cấu hình: cột AK xác định sheet và bảng đích, cột AM liệt kê ô đích (nối bằng &)";
                    e1.Params["locked"] = "khóa khi (bảng động) kết quả công thức cột AL áp cho từng dòng dữ liệu == \"0\", hoặc (ô tĩnh) giá trị cột AL == 0, ngược lại mở khóa";
                    e1.Params["clearData"] = "ô cột AN của dòng cấu hình == \"1\"";
                    if (args.Count > 2) e1.Params["rowCheck"] = R.Expr(Arg(inv, 2, st, f), new Flags());
                    if (args.Count > 3) e1.Params["tableCheck"] = R.Expr(Arg(inv, 3, st, f), new Flags());
                    e1.Params["note"] = "ô đích theo cấu hình lúc chạy";
                    AddEffect(e1, st);
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
                    V dv = MapVal(recv, dn, st, f, ln);
                    var oa = args[1];
                    string on = oa.Expression is IdentifierNameSyntax oi ? oi.Identifier.Text : oa.Expression is DeclarationExpressionSyntax de && de.Designation is SingleVariableDesignationSyntax sv ? sv.Identifier.Text : null;
                    if (on != null) RegSym(on, "out", ln, "").Desc = "giá trị tra từ bảng " + dn + " theo khóa " + R.Expr(key, new Flags()) + " tại dòng " + ln + ", không có khóa thì giá trị mặc định";
                    return new VCond(new CAtom("hasKey", dv, key));
                }
            case "ContainsKey" when recv != null && !staticRecv && args.Count == 1:
                {
                    var key = Arg(inv, 0, st, f);
                    string dn = MapName(recv);
                    return new VCond(new CAtom("hasKey", MapVal(recv, dn, st, f, ln), key));
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
                    if (name == "Parse" && rt is "bool" or "Boolean")
                    {
                        // ép kiểu theo nguồn: bool.Parse ra bool(...), không int(...)
                        if (v is VConst kb0) { if (kb0.Val is bool) return kb0; if (kb0.Val is string sb0 && bool.TryParse(sb0.Trim(), out var bb0)) return new VConst(bb0); }
                        return VLin.Atom(((VSym)Fn("bool", v)).Text);
                    }
                    bool isInt = !(name.Contains("Double") || name.Contains("Decimal")) && rt != "double" && rt != "decimal";
                    if (v is VConst k)
                    {
                        if (k.Val is string s) { if (isInt && long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var l)) return new VConst(l); if (double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return isInt ? new VConst((long)d) : new VConst(d); }
                        if (k.IsNum) return k;
                    }
                    if (v is VLin vl && (!isInt || !vl.T.Keys.Any(a => a.Contains("cell(") || a.Contains("num("))) && !vl.T.Keys.Any(strAtoms.Contains)) return v;
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
                    return new VOpaque(name, new List<V> { coll, lam != null ? lam : Arg(inv, 0, st, f) });
                }
            case "GetValueCustom" when args.Count >= 3:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var r = Arg(inv, 1, st, f); var c = Arg(inv, 2, st, f);
                    if (sh != null) return Member(new VCell(sh, "cell", Ar.Add(r, new VConst(-1L)), Ar.Add(c, new VConst(-1L))), "value", st);
                    break;
                }
            case "GetObjectValue" when args.Count == 2:
                {
                    var sh = Arg(inv, 0, st, f) as VSheet; var id = ArrayCellIdOf(Arg(inv, 1, st, f));
                    if (sh != null && id is VArr ar && ar.Items.Count == 2) return Member(new VCell(sh, "cell", ar.Items[1], ar.Items[0]), "value", st);
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
                    AddEffect(e1, st);
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
                            AddEffect(e1, st);
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
                        AddEffect(e1, st);
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
                        AddEffect(e1, st); return VConst.Null;
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
            case "Show" when rt == "MessageBox" || rt.EndsWith(".MessageBox", StringComparison.Ordinal):
                {
                    // kết quả hộp thoại: người dùng chọn lúc chạy, không phải hằng
                    var av = Enumerable.Range(0, args.Count).Select(i => Arg(inv, i, st, f)).ToList();
                    string btn = args.Count >= 3 && av[2] is VConst bk && bk.Val is long bl && bl >= 0 && bl <= 5 ? new[] { "OK", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel" }[bl] : args.Count >= 3 ? "?" : "OK";
                    var si = RegSym("user_choice", "dialog", ln, "nút người dùng bấm trên hộp thoại MessageBox tại dòng " + ln + " (nút " + btn + "), chọn lúc chạy, mã DialogResult: 1 OK, 2 Cancel, 3 Abort, 4 Retry, 5 Ignore, 6 Yes, 7 No");
                    return VLin.Atom(si.Ref + "(MessageBox " + btn + ", dòng " + ln + ")" + Mk.Tag(Mk.Next()));
                }
            case "ShowDialog" when args.Count == 0 && !isStmt:
                {
                    var si = RegSym("user_choice", "dialog", ln, "kết quả hộp thoại ShowDialog tại dòng " + ln + ", người dùng chọn lúc chạy, mã DialogResult: 1 OK, 2 Cancel, 6 Yes, 7 No");
                    if (recv != null) Eval(recv, st, f);
                    return VLin.Atom(si.Ref + "(ShowDialog, dòng " + ln + ")" + Mk.Tag(Mk.Next()));
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
        var argv = args.Select(a => Eval(a.Expression, st, f)).ToList();
        Opaque(qn, inv, isStmt ? "statement" : "value");
        bool classLocal = (recv == null || recv is ThisExpressionSyntax) && f.Cls != null && f.Cls.Methods.ContainsKey(name);
        SymInfo cls0 = null;
        if (classLocal) { cls0 = RegSym(name, "method", ln, "method cùng class không nội tuyến tại dòng " + ln + ": không khớp chữ ký (số đối số " + args.Count + ")"); Wild(st, ln, "lời gọi " + name + " không nội tuyến dòng " + ln + " (có thể ghi ô)", "co_the_ghi", new[] { name }); fieldVer = Mk.Next(); }
        bool localRecv = rv is VArr || rv is VConst || (rv is VSym rsy && rsy.Text.StartsWith("new ", StringComparison.Ordinal));
        bool touches = SheetLike(rv) || argv.Any(SheetLike) || argv.Any(a => a is VSym ts0 && ts0.Text == "this");
        if (isStmt && !localRecv)
        {
            var e1 = NewEffect(st, f, ln); e1.Op = "opaque"; e1.Target = null;
            e1.Params["call"] = qn;
            e1.Params["args"] = argv.Select(a => R.Expr(a, new Flags())).ToList();
            e1.Params["touchesSheet"] = touches;
            AddEffect(e1, st);
        }
        else if (touches && CallWriteKind(qn) is string wk)
            Wild(st, ln, wk == "ghi" ? "lệnh " + qn + " dòng " + ln + " (đổi giá trị hoặc dịch ô)" : "lời gọi chưa giải " + qn + " dòng " + ln + " (có thể ghi ô)", wk, new[] { qn });
        // lời gọi trên đối tượng: đối số đầu của nhãn là đối tượng nhận
        return classLocal ? new VOpaque("call", new List<V> { new VSym(cls0.Ref) }.Concat(argv).ToList()) : new VOpaque(qn, rv != null ? new List<V> { rv }.Concat(argv).ToList() : argv);
    }

    /// checkLockUnLockCells theo cấu hình mẫu: nhánh ô tĩnh (dòng mẫu không thuộc bảng động) và nhánh bảng động (lặp mỗi dòng dữ liệu)
    void CheckLockFromConfig(InvocationExpressionSyntax inv, State st, Frame f, int ln)
    {
        V rowChk = new VConst(-1L), tblChk = new VConst(-1L);
        var a = inv.ArgumentList.Arguments;
        for (int i = 2; i < a.Count; i++)
        {
            string nm = a[i].NameColon?.Name.Identifier.Text ?? (i == 2 ? "iRowCheck" : "tableIndexCheck");
            var v = Eval(a[i].Expression, st, f);
            if (nm == "iRowCheck") rowChk = v; else tblChk = v;
        }
        var cfg = Lock;
        var last = new VSheet(new VConst((long)cfg.LastSheet));
        var ba1 = Member(new VCell(last, "cell", new VConst(0L), new VConst(52L)), "text", st);
        var en = C.Cmp("==", Fn("trim", ba1), new VConst("1"));
        var path0 = st.Path.Add(new PathItem(en, ln));
        var dataRow = RegSym("dataRow", "checkLockUnLockCells", ln, "dòng dữ liệu thứ j của bảng động chứa dòng mẫu (DataLocationTemplateRowStart + j), j chạy qua mọi dòng dữ liệu của bảng");
        Effect Mk1(string op, ImmutableList<PathItem> path, LockRow r)
        {
            var e = NewEffect(st, f, ln); e.Op = op; e.Path = path;
            e.Params["via"] = "checkLockUnLockCells"; e.Params["template"] = cfg.Template;
            e.Params["configRow"] = "S" + cfg.LastSheet + "!AK" + r.Row + ":AN" + r.Row;
            e.Params["control"] = "S" + r.RefSheet + "!" + r.ControlCol + r.ControlRow;
            if (r.Sheet != r.RefSheet) e.Params["sheetNote"] = "DetectIndexOfSheet dùng Contains: theo tên tab trong mẫu thì sheet đầu tiên khớp là S" + r.Sheet + " (tên là chuỗi con của tên sheet được tham chiếu), tên sheet lúc chạy lấy từ thông tin sheet của tờ khai (WinSpreadProcess.cs:1701) nên chưa xác định, ghi theo sheet được tham chiếu S" + r.RefSheet;
            return e;
        }
        foreach (var r in cfg.Rows)
        {
            if (r.Error != null || r.Skip != null) continue;
            // sheet đích: sheet mà AK tham chiếu (DetectIndexOfSheet so Contains trên tên sheet lúc chạy, gán từ sInfo.SheetName tại WinSpreadProcess.cs:1701)
            var dyn = new CAtom("inDynamicTable", new VConst((long)r.RefSheet), new VConst((long)r.ControlRow));
            Cond tblOk = tblChk is VConst tk && tk.Val is long tv && tv < 0 ? CConst.T : C.Or(C.Cmp("<", tblChk, new VConst(0L)), C.Cmp("==", Fn("tableIndexOf", new VConst((long)r.RefSheet), new VConst((long)r.ControlRow)), tblChk));
            var dynC = C.And(dyn, tblOk);
            var sheetV = new VSheet(new VConst((long)r.RefSheet));
            string clear = r.ANformula ? "cell(S" + cfg.LastSheet + "!AN" + r.Row + ").text == \"1\" (ô AN là công thức)" : r.AN == "1" ? "có (AN" + r.Row + " = \"1\" trong mẫu): khi khóa thì xóa giá trị ô (số 0, checkbox false, chữ rỗng)" : null;
            // nhánh bảng động: ghi công thức AL dịch dòng vào cột BB (ColumnLettersToIndex("BA") = 53 dùng làm chỉ số cột gốc 0) rồi đọc lại chữ
            if (r.AL != null)
            {
                var pd = path0.Add(new PathItem(dynC, ln));
                var fcell = new VCell(last, "cell", new VConst((long)(r.Row - 1)), new VConst(53L));
                string tmpl = System.Text.RegularExpressions.Regex.Replace(r.ALn ?? "", "(?<=[A-Z])" + r.ControlRow.ToString(CultureInfo.InvariantCulture), "{" + dataRow.Ref + "}");
                var ef = Mk1("setFormula", pd, r); ef.Target = fcell; ef.Val = new VConst(tmpl);
                ef.Params["note"] = "công thức AL" + r.Row + " (" + r.AL + ") đổi dòng mẫu " + r.ControlRow + " thành dòng dữ liệu, ghi vào cột BB vì ColumnLettersToIndex(\"BA\") = 53 dùng làm chỉ số cột gốc 0";
                AddEffect(ef, st);
                if (!r.AMempty)
                {
                    var txt = Member(fcell, "text", st);
                    var lockedD = C.Cmp("==", txt, new VConst("0"));
                    foreach (var t in r.Targets)
                    {
                        var col = Fmt.LettersToIndex(t.Col) - 1;
                        var tr = t.Row == r.ControlRow ? Ar.Add(VLin.Atom(dataRow.Ref), new VConst(-1L)) : Ar.Add(VLin.Atom(dataRow.Ref), new VConst((long)(t.Row - r.ControlRow - 1)));
                        var el = Mk1("lock", pd, r); el.Target = new VCell(sheetV, "cell", tr, new VConst(col));
                        el.Params["locked"] = R.Cond(lockedD, new Flags()); el.Val = new VCond(lockedD);
                        el.Params["rows"] = "mỗi dòng dữ liệu của bảng, ô mẫu " + t.Col + t.Row + " dịch theo dòng";
                        if (clear != null) el.Params["clearData"] = clear;
                        el.Params["backColor"] = "MA_COLOR_LOCKED khi khóa, MA_COLOR_UNLOCK khi mở";
                        AddEffect(el, st);
                    }
                }
            }
            // nhánh ô tĩnh
            var ps = path0.Add(new PathItem(C.Not(dynC), ln));
            if (r.AMempty) { var es = Mk1("lock", ps, r); es.Target = null; es.Params["targetReason"] = "AM" + r.Row + " không có công thức: vòng lặp cấu hình dừng (break) ở nhánh ô tĩnh"; AddEffect(es, st); break; }
            var alv = Member(new VCell(last, "cell", new VConst((long)(r.Row - 1)), new VConst(37L)), "value", st);
            var lockedS = C.Cmp("==", VLin.Atom(((VSym)Fn("int", alv)).Text), new VConst(0L));
            foreach (var t in r.Targets)
            {
                var rc = C.Or(C.Cmp("<", rowChk, new VConst(0L)), C.Cmp("==", rowChk, new VConst((long)t.Row)));
                var el = Mk1("lock", rc is CConst ? ps : ps.Add(new PathItem(rc, ln)), r);
                el.Target = new VCell(sheetV, "cell", new VConst((long)(t.Row - 1)), new VConst(Fmt.LettersToIndex(t.Col) - 1));
                el.Params["locked"] = R.Cond(lockedS, new Flags()); el.Val = new VCond(lockedS);
                el.Params["alFormula"] = r.ALn ?? "(không có công thức)";
                if (r.ALvalue != null) el.Params["alTemplateValue"] = r.ALvalue;
                if (clear != null) el.Params["clearData"] = clear;
                el.Params["backColor"] = "MA_COLOR_LOCKED khi khóa, MA_COLOR_UNLOCK khi mở";
                AddEffect(el, st);
            }
        }
    }

    void AddLock(State st, Frame f, int ln, VCell cell, V lk, string via, V del = null)
    {
        var e2 = NewEffect(st, f, ln); e2.Target = cell; e2.Params["via"] = via;
        if (lk is VConst k && k.Val is bool b) e2.Op = b ? "lock" : "unlock";
        else { e2.Op = "lock"; e2.Params["locked"] = R.Expr(lk, new Flags()); e2.Val = lk; }
        if (del != null && !(del is VConst dk && dk.Val is bool db && !db)) e2.Params["clearData"] = R.Expr(del, new Flags());
        AddEffect(e2, st);
    }

    static V Fn(string fn, params V[] a) => Ar.Fn(fn, a);
    /// giá trị ký hiệu của bảng tra: biến cục bộ đã có giá trị ký hiệu thì dùng; tên cục bộ chưa có thì khai symbols local-map
    V MapVal(ExpressionSyntax recv, string dn, State st, Frame f, int ln)
    {
        if (recv is IdentifierNameSyntax rin && st.Env.TryGetValue(rin.Identifier.Text, out var de0) && de0.Val is VSym) return de0.Val;
        if (!dn.Contains('.') && (f.Cls == null || !f.Cls.Members.Contains(dn))) return new VSym(RegSym(dn, "local-map", ln, "bảng tra (từ điển) cục bộ, nội dung nạp lúc chạy").Ref);
        return new VSym(dn);
    }
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
    /// từ điển/bảng tra cục bộ tạo mới (nội dung nạp lúc chạy): giữ tên biến làm ký hiệu thay cho obj(Dictionary) vô danh
    V NamedCollection(string n, V v, int line)
    {
        if (v is VSym y && y.Text.StartsWith("obj(", StringComparison.Ordinal) && (y.Text.Contains("Dictionary") || y.Text.Contains("Hashtable") || y.Text.Contains("SortedList")))
        {
            var sl = RegSym(n, "local-map", line, "bảng tra (từ điển) cục bộ tạo mới tại dòng " + line + ", nội dung nạp lúc chạy");
            return new VSym(sl.Ref + Mk.Tag(Mk.Next()));
        }
        return v;
    }
    V UninitSym(string n, int line) => new VSym(RegSym(n, "uninit", line, "khai báo tại dòng " + line + " không có giá trị đầu").Ref);
    V ParamSym(string n, BaseMethodDeclarationSyntax m) => new VSym(RegSym(n, "param", Line(m), "tham số của method " + MethodName(m) + " không được truyền").Ref);

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
        if (s == null) return new VOpaque("Format", new List<V> { fmt }.Concat(vals).ToList());
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
                if (j < 0) return new VOpaque("Format", new List<V> { fmt }.Concat(vals).ToList());
                var full = s.Substring(i + 1, j - i - 1); var spec = full.Split(':', ',')[0];
                if (!int.TryParse(spec, out var ai) || ai >= vals.Count) return new VOpaque("Format", new List<V> { fmt }.Concat(vals).ToList());
                parts.Add(new VConst(sb.ToString())); sb.Clear();
                parts.Add(full.Length > spec.Length ? Ar.Fn("format", vals[ai], new VConst(full.Substring(spec.Length))) : vals[ai]); i = j + 1; continue;
            }
            sb.Append(ch); i++;
        }
        parts.Add(new VConst(sb.ToString()));
        return VCat.Make(parts);
    }

    Dictionary<string, string> outerLam = new(StringComparer.Ordinal);
    /// hàm vô danh dạng trung tính: fn(it: thân), tham số đổi tên it/it1/it2; thân khối nhiều lệnh ghi phạm vi dòng
    string LamText(LambdaExpressionSyntax l, State st, Frame f, int depth)
    {
        var ps = l switch { SimpleLambdaExpressionSyntax s1 => new List<string> { s1.Parameter.Identifier.Text }, ParenthesizedLambdaExpressionSyntax p1 => p1.ParameterList.Parameters.Select(x => x.Identifier.Text).ToList(), _ => new List<string>() };
        var body = l switch { SimpleLambdaExpressionSyntax s2 => s2.Body, ParenthesizedLambdaExpressionSyntax p2 => p2.Body, _ => null };
        var ren = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int i = 0; i < ps.Count; i++) ren[ps[i]] = (depth == 0 ? "it" : "it" + depth + "_") + (ps.Count == 1 ? "" : (i + 1).ToString(CultureInfo.InvariantCulture));
        foreach (var kv in outerLam) if (!ren.ContainsKey(kv.Key)) ren[kv.Key] = kv.Value;
        var saveLam = outerLam; outerLam = ren;
        var set = new HashSet<string>(ren.Keys, StringComparer.Ordinal);
        string bt = body is ExpressionSyntax be ? Neu(be, set, st, f, depth)
            : body is BlockSyntax bk ? (bk.Statements.Count == 1 && bk.Statements[0] is ReturnStatementSyntax rs && rs.Expression != null ? Neu(rs.Expression, set, st, f, depth) : "khối lệnh dòng " + Line(bk) + "-" + (bk.GetLocation().GetLineSpan().EndLinePosition.Line + 1))
            : "?";
        outerLam = saveLam;
        return "fn(" + string.Join(", ", ps.Select(x => ren[x])) + ": " + bt + ")";
    }
    string Neu(ExpressionSyntax e, HashSet<string> ps, State st, Frame f, int depth)
    {
        string N(ExpressionSyntax x) => Neu(x, ps, st, f, depth);
        string A(ArgumentListSyntax al) => string.Join(", ", al.Arguments.Select(a => N(a.Expression)));
        switch (e)
        {
            case ParenthesizedExpressionSyntax p: return "(" + N(p.Expression) + ")";
            case CastExpressionSyntax c: return N(c.Expression);
            case LiteralExpressionSyntax lit: return lit.IsKind(SyntaxKind.StringLiteralExpression) ? R.Expr(new VConst(lit.Token.ValueText), new Flags()) : lit.Token.Text;
            case IdentifierNameSyntax id:
                {
                    var n = id.Identifier.Text;
                    if (ps.Contains(n)) return outerLam[n];
                    // biến/trường bắt được: viết theo giá trị đang có (như mọi chỗ khác); tên kiểu/lớp tĩnh giữ nguyên
                    V cv = st.Env.TryGetValue(n, out var en) ? en.Val : f.Cls != null && f.Cls.Members.Contains(n) ? Lookup(n, st, f, id) : null;
                    if (cv == null) return n;
                    var ct = R.Expr(cv, new Flags());
                    return cv is VConst || cv is VSym || ct.IndexOf(' ') < 0 ? ct : "(" + ct + ")";
                }
            case MemberAccessExpressionSyntax ma:
                {
                    // điều hướng cấu trúc ô/bảng viết như chỗ khác: bỏ CellTemplate/Location, chỉ số dòng/cột dạng hàm
                    var mn = ma.Name.Identifier.Text;
                    if (mn is "CellTemplate" or "Location") return N(ma.Expression);
                    if (mn == "RowIndex") return "rowIndex(" + N(ma.Expression) + ")";
                    if (mn == "ColIndex") return "colIndex(" + N(ma.Expression) + ")";
                    return N(ma.Expression) + "." + mn;
                }
            case InvocationExpressionSyntax iv:
                {
                    var a = A(iv.ArgumentList);
                    if (iv.Expression is MemberAccessExpressionSyntax m)
                    {
                        if (m.Name.Identifier.Text == "Equals" && iv.ArgumentList.Arguments.Count == 1) return N(m.Expression) + " == " + a;
                        return N(m.Expression) + "." + m.Name.Identifier.Text + "(" + a + ")";
                    }
                    return (iv.Expression is IdentifierNameSyntax ii ? ii.Identifier.Text : iv.Expression is GenericNameSyntax gn ? gn.Identifier.Text : "expr:" + iv.Expression.Kind()) + "(" + a + ")";
                }
            case BinaryExpressionSyntax b: return N(b.Left) + " " + b.OperatorToken.Text + " " + N(b.Right);
            case PrefixUnaryExpressionSyntax pu: return pu.OperatorToken.Text + N(pu.Operand);
            case ElementAccessExpressionSyntax ea:
                {
                    var ix = string.Join(", ", ea.ArgumentList.Arguments.Select(a => N(a.Expression)));
                    if (ea.Expression is MemberAccessExpressionSyntax em && em.Name.Identifier.Text is "Cells" or "Rows" or "Columns") return N(em.Expression) + "." + em.Name.Identifier.Text.ToLowerInvariant() + "[" + ix + "]";
                    return N(ea.Expression) + "[" + ix + "]";
                }
            case ConditionalExpressionSyntax ce: return "(" + N(ce.Condition) + " ? " + N(ce.WhenTrue) + " : " + N(ce.WhenFalse) + ")";
            case LambdaExpressionSyntax l2: return depth < 3 ? LamText(l2, st, f, depth + 1) : "fn(...)";
            case ThisExpressionSyntax: return "this";
            case PredefinedTypeSyntax pt: return pt.Keyword.ValueText;
            case GenericNameSyntax g2: return g2.Identifier.Text;
            case InterpolatedStringExpressionSyntax istr: return "concat(" + string.Join(", ", istr.Contents.Select(c => c is InterpolatedStringTextSyntax t ? R.Expr(new VConst(t.TextToken.ValueText), new Flags()) : N(((InterpolationSyntax)c).Expression))) + ")";
            default: return "expr:" + e.Kind();
        }
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
