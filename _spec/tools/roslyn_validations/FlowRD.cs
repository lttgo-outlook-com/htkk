// Phân tích reaching definitions trên ControlFlowGraph (Microsoft.CodeAnalysis.FlowAnalysis) cho biến cục bộ và tham số.
// Chỉ dùng SemanticModel của CSharpCompilation dựng trong bộ nhớ; không Emit, không chạy mã.
using System.Collections;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.FlowAnalysis;
using Microsoft.CodeAnalysis.Operations;

public sealed class DefSite
{
    public int Id;
    public ISymbol Sym;
    public SyntaxNode Node;     // cú pháp của phép gán (VariableDeclarator, AssignmentExpression, Argument, ForEach, ...)
    public string Kind;         // set | init | compound | inc | out | ref | foreach | pattern | catch | param | decon
    public bool Kill;
    public int TargetUsePos = -1; // với compound/inc/ref: vị trí đọc biến tại chính phép gán (để lấy các phép gán trước nó)
}

public sealed class UseInfo
{
    public ISymbol Sym;
    public List<int> Defs = new();
    public bool Reachable;
    public BitArray State;      // trạng thái RD ngay trước điểm dùng
}

public sealed class RDResult
{
    public readonly List<DefSite> Defs = new();
    public readonly Dictionary<int, UseInfo> Uses = new();
    public SemanticModel Model;
    public int Blocks;
    public bool Overflow;

    public IEnumerable<DefSite> ReachingAt(ISymbol sym, BitArray state)
    {
        for (int i = 0; i < Defs.Count; i++)
            if (state[i] && SymbolEqualityComparer.Default.Equals(Defs[i].Sym, sym)) yield return Defs[i];
    }
}

public static class FlowRD
{
    public static RDResult Analyze(SemanticModel model, MethodDeclarationSyntax md)
    {
        IMethodBodyOperation body = model.GetOperation(md) as IMethodBodyOperation;
        if (body == null) return null;
        ControlFlowGraph g = ControlFlowGraph.Create(body);
        if (g == null) return null;
        var an = new Analyzer(model, md);
        an.Run(g);
        an.Res.Model = model;
        return an.Res;
    }

    sealed class Analyzer
    {
        public readonly RDResult Res = new();
        readonly SemanticModel model;
        readonly MethodDeclarationSyntax md;
        readonly Dictionary<(SyntaxNode, ISymbol), int> defIndex = new(new KeyCmp());
        readonly Dictionary<ISymbol, List<int>> defsOfSym = new(SymbolEqualityComparer.Default);
        readonly Dictionary<ISymbol, BitArray> symMask = new(SymbolEqualityComparer.Default);
        bool record;      // lượt cuối: ghi nhận điểm dùng
        bool reachableBlock;

        sealed class KeyCmp : IEqualityComparer<(SyntaxNode, ISymbol)>
        {
            public bool Equals((SyntaxNode, ISymbol) a, (SyntaxNode, ISymbol) b) => a.Item1 == b.Item1 && SymbolEqualityComparer.Default.Equals(a.Item2, b.Item2);
            public int GetHashCode((SyntaxNode, ISymbol) k) => k.Item1.GetHashCode() * 31 + SymbolEqualityComparer.Default.GetHashCode(k.Item2);
        }

        public Analyzer(SemanticModel m, MethodDeclarationSyntax d) { model = m; md = d; }

        int Width => Res.Defs.Count;

        int DefId(ISymbol sym, SyntaxNode node, string kind, bool kill, int targetUse)
        {
            if (defIndex.TryGetValue((node, sym), out int id)) return id;
            id = Res.Defs.Count;
            Res.Defs.Add(new DefSite { Id = id, Sym = sym, Node = node, Kind = kind, Kill = kill, TargetUsePos = targetUse });
            defIndex[(node, sym)] = id;
            if (!defsOfSym.TryGetValue(sym, out var l)) defsOfSym[sym] = l = new List<int>();
            l.Add(id);
            return id;
        }

        // ---- lượt 1: liệt kê mọi phép gán (để cố định độ rộng bit) ----
        BitArray cur; // trạng thái đang đi
        bool enumerating;

        public void Run(ControlFlowGraph g)
        {
            // tham số: phép gán ngầm ở đầu hàm
            foreach (var p in md.ParameterList.Parameters)
            {
                var ps = model.GetDeclaredSymbol(p);
                if (ps != null) DefId(ps, p, "param", true, -1);
            }
            var graphs = new List<ControlFlowGraph>();
            CollectGraphs(g, graphs);
            enumerating = true;
            foreach (var gg in graphs)
                foreach (var b in gg.Blocks)
                {
                    curGraph = gg;
                    cur = null;
                    foreach (var op in b.Operations) Walk(op);
                    if (b.BranchValue != null) Walk(b.BranchValue);
                }
            enumerating = false;
            foreach (var kv in defsOfSym)
            {
                var m = new BitArray(Width);
                foreach (int i in kv.Value) m[i] = true;
                symMask[kv.Key] = m;
            }
            var entry = new BitArray(Width);
            foreach (var d in Res.Defs) if (d.Kind == "param") entry[d.Id] = true;
            var queue = new Queue<(ControlFlowGraph g, BitArray entry)>();
            queue.Enqueue((g, entry));
            while (queue.Count > 0)
            {
                var (gg, en) = queue.Dequeue();
                pendingCapture = new();
                Solve(gg, en);
                // hàm ẩn danh: trạng thái vào = trạng thái tại chỗ tạo hàm
                foreach (var (f, st) in pendingCapture)
                {
                    ControlFlowGraph lg = null;
                    try { lg = gg.GetAnonymousFunctionControlFlowGraph(f); } catch { }
                    if (lg != null) queue.Enqueue((lg, st));
                }
                // hàm cục bộ: trạng thái vào = trạng thái đầu hàm
                foreach (var lf in gg.LocalFunctions)
                {
                    ControlFlowGraph lg = null;
                    try { lg = gg.GetLocalFunctionControlFlowGraph(lf); } catch { }
                    if (lg != null) queue.Enqueue((lg, entry));
                }
            }
        }

        static void CollectGraphs(ControlFlowGraph g, List<ControlFlowGraph> acc)
        {
            acc.Add(g);
            foreach (var b in g.Blocks)
                foreach (var op in b.Operations.Concat(b.BranchValue != null ? new[] { b.BranchValue } : Array.Empty<IOperation>()))
                    foreach (var f in op.DescendantsAndSelf().OfType<IFlowAnonymousFunctionOperation>())
                    {
                        ControlFlowGraph lg = null;
                        try { lg = g.GetAnonymousFunctionControlFlowGraph(f); } catch { }
                        if (lg != null) CollectGraphs(lg, acc);
                    }
            foreach (var lf in g.LocalFunctions)
            {
                ControlFlowGraph lg = null;
                try { lg = g.GetLocalFunctionControlFlowGraph(lf); } catch { }
                if (lg != null) CollectGraphs(lg, acc);
            }
        }

        void Solve(ControlFlowGraph g, BitArray entry)
        {
            int n = g.Blocks.Length;
            curGraph = g;
            Res.Blocks += n;
            var inS = new BitArray[n];
            var outS = new BitArray[n];
            for (int i = 0; i < n; i++) { inS[i] = new BitArray(Width); outS[i] = new BitArray(Width); }
            // khối đầu của catch/filter/finally không có cạnh vào: lấy hợp trạng thái của mọi khối trong vùng try tương ứng
            var excSrc = new Dictionary<int, (int from, int to)>();
            foreach (var b in g.Blocks)
            {
                for (var r = b.EnclosingRegion; r != null; r = r.EnclosingRegion)
                {
                    if (r.FirstBlockOrdinal != b.Ordinal) break;
                    if (r.Kind is ControlFlowRegionKind.Catch or ControlFlowRegionKind.Filter or ControlFlowRegionKind.Finally or ControlFlowRegionKind.FilterAndHandler)
                    {
                        var t = r.EnclosingRegion;
                        while (t != null && t.Kind != ControlFlowRegionKind.TryAndCatch && t.Kind != ControlFlowRegionKind.TryAndFinally) t = t.EnclosingRegion;
                        var tryR = t?.NestedRegions.FirstOrDefault(x => x.Kind == ControlFlowRegionKind.Try);
                        if (tryR != null) excSrc[b.Ordinal] = (tryR.FirstBlockOrdinal, tryR.LastBlockOrdinal);
                        break;
                    }
                }
            }
            bool changed = true;
            int iter = 0;
            while (changed && iter < 500)
            {
                changed = false; iter++;
                foreach (var b in g.Blocks)
                {
                    var nin = new BitArray(Width);
                    if (b.Kind == BasicBlockKind.Entry) nin.Or(entry);
                    foreach (var pr in b.Predecessors) nin.Or(outS[pr.Source.Ordinal]);
                    if (excSrc.TryGetValue(b.Ordinal, out var rg))
                        for (int k = rg.from; k <= rg.to; k++) { nin.Or(inS[k]); nin.Or(outS[k]); }
                    inS[b.Ordinal] = nin;
                    cur = (BitArray)nin.Clone();
                    record = false; reachableBlock = b.IsReachable;
                    foreach (var op in b.Operations) Walk(op);
                    if (b.BranchValue != null) Walk(b.BranchValue);
                    if (!Same(cur, outS[b.Ordinal])) { outS[b.Ordinal] = cur; changed = true; }
                }
            }
            // lượt ghi nhận
            foreach (var b in g.Blocks)
            {
                cur = (BitArray)inS[b.Ordinal].Clone();
                record = true; reachableBlock = b.IsReachable;
                foreach (var op in b.Operations) Walk(op);
                if (b.BranchValue != null) Walk(b.BranchValue);
            }
            record = false;
        }

        static bool Same(BitArray a, BitArray b)
        {
            if (a.Length != b.Length) return false;
            var x = (BitArray)a.Clone(); x.Xor(b);
            foreach (bool v in x) if (v) return false;
            return true;
        }

        // CFG có thể "spill" đích phép gán thành flow capture khi vế phải có rẽ nhánh (?:, ??, &&, ||, ?.)
        ControlFlowGraph curGraph;
        readonly Dictionary<(ControlFlowGraph, CaptureId), ISymbol> caps = new();
        static bool IsAssignTarget(SyntaxNode syn) =>
            syn?.Parent is AssignmentExpressionSyntax ae && ae.Left == syn
            || syn?.Parent is PostfixUnaryExpressionSyntax || syn?.Parent is PrefixUnaryExpressionSyntax pu && (pu.IsKind(SyntaxKind.PreIncrementExpression) || pu.IsKind(SyntaxKind.PreDecrementExpression));
        ISymbol TargetSym(IOperation t) => t switch
        {
            IFlowCaptureReferenceOperation r => caps.TryGetValue((curGraph, r.Id), out var cs) ? cs : null,
            ILocalReferenceOperation l => l.Local,
            IParameterReferenceOperation p => p.Parameter,
            _ => null,
        };

        void Gen(ISymbol sym, SyntaxNode node, string kind, bool kill, int targetUse)
        {
            if (sym == null || node == null) return;
            if (sym is ILocalSymbol ls && ls.IsImplicitlyDeclared) return;
            int id = DefId(sym, node, kind, kill, targetUse);
            if (enumerating || cur == null) return;
            if (id >= cur.Length) { Res.Overflow = true; return; }
            if (kill && symMask.TryGetValue(sym, out var m))
            {
                var nm = (BitArray)m.Clone(); nm.Not(); cur.And(nm);
            }
            cur[id] = true;
        }

        void Use(ISymbol sym, SyntaxNode syn)
        {
            if (enumerating || !record || sym == null || syn == null) return;
            if (sym is ILocalSymbol ls && ls.IsImplicitlyDeclared) return;
            int pos = syn.SpanStart;
            if (!Res.Uses.TryGetValue(pos, out var u)) Res.Uses[pos] = u = new UseInfo { Sym = sym, State = new BitArray(Width) };
            u.Reachable |= reachableBlock;
            u.State.Or(cur);
            if (defsOfSym.TryGetValue(sym, out var ids))
                foreach (int i in ids) if (i < cur.Length && cur[i] && !u.Defs.Contains(i)) u.Defs.Add(i);
            u.Defs.Sort((a, b) => Res.Defs[a].Node.SpanStart.CompareTo(Res.Defs[b].Node.SpanStart));
        }

        static ArgumentSyntax RefArg(SyntaxNode s)
        {
            var p = s?.Parent;
            if (p is DeclarationExpressionSyntax) p = p.Parent;
            if (p is ArgumentSyntax a && (a.RefKindKeyword.IsKind(SyntaxKind.OutKeyword) || a.RefKindKeyword.IsKind(SyntaxKind.RefKeyword))) return a;
            return null;
        }

        void Walk(IOperation op)
        {
            if (op == null) return;
            switch (op)
            {
                case ISimpleAssignmentOperation a:
                    {
                        Walk(a.Value);
                        var s = TargetSym(a.Target);
                        if (s != null) Gen(s, DefSyntax(a), a.Syntax is VariableDeclaratorSyntax ? "init" : a.Syntax is AssignmentExpressionSyntax ? "set" : KindOf(a.Syntax), true, -1);
                        else Walk(a.Target);
                        return;
                    }
                case ICompoundAssignmentOperation c:
                    {
                        var s = TargetSym(c.Target);
                        if (s != null) { Use(s, c.Target.Syntax); Walk(c.Value); Gen(s, c.Syntax, "compound", true, c.Target.Syntax.SpanStart); }
                        else { Walk(c.Target); Walk(c.Value); }
                        return;
                    }
                case ICoalesceAssignmentOperation ca:
                    {
                        var s = TargetSym(ca.Target);
                        if (s != null) { Use(s, ca.Target.Syntax); Walk(ca.Value); Gen(s, ca.Syntax, "compound", false, ca.Target.Syntax.SpanStart); }
                        else { Walk(ca.Target); Walk(ca.Value); }
                        return;
                    }
                case IIncrementOrDecrementOperation inc:
                    {
                        var s = TargetSym(inc.Target);
                        if (s != null) { Use(s, inc.Target.Syntax); Gen(s, inc.Syntax, "inc", true, inc.Target.Syntax.SpanStart); }
                        else Walk(inc.Target);
                        return;
                    }
                case IDeconstructionAssignmentOperation da:
                    {
                        Walk(da.Value);
                        foreach (var t in da.Target.DescendantsAndSelf())
                        {
                            var s = TargetSym(t);
                            if (s != null) Gen(s, da.Syntax, "decon", true, -1);
                        }
                        return;
                    }
                case IDeclarationPatternOperation dp:
                    if (dp.DeclaredSymbol != null) Gen(dp.DeclaredSymbol, dp.Syntax, "pattern", true, -1);
                    return;
                case IRecursivePatternOperation rp:
                    foreach (var ch in rp.ChildOperations) Walk(ch);
                    if (rp.DeclaredSymbol != null) Gen(rp.DeclaredSymbol, rp.Syntax, "pattern", true, -1);
                    return;
                case ILocalReferenceOperation lr:
                    LocalRef(lr.Local, lr.Syntax, lr.IsDeclaration);
                    return;
                case IParameterReferenceOperation pr:
                    LocalRef(pr.Parameter, pr.Syntax, false);
                    return;
                case IFlowCaptureOperation fc when (fc.Value is ILocalReferenceOperation || fc.Value is IParameterReferenceOperation) && IsAssignTarget(fc.Value.Syntax):
                    caps[(curGraph, fc.Id)] = fc.Value is ILocalReferenceOperation cl ? cl.Local : ((IParameterReferenceOperation)fc.Value).Parameter;
                    return;
                case IFlowAnonymousFunctionOperation f:
                    if (!enumerating && record && pendingCapture != null) pendingCapture.Add((f, (BitArray)cur.Clone()));
                    return;
            }
            foreach (var ch in op.ChildOperations) Walk(ch);
        }
        List<(IFlowAnonymousFunctionOperation, BitArray)> pendingCapture;

        void LocalRef(ISymbol sym, SyntaxNode syn, bool isDecl)
        {
            var ra = RefArg(syn);
            if (ra != null)
            {
                if (ra.RefKindKeyword.IsKind(SyntaxKind.RefKeyword)) { Use(sym, syn); Gen(sym, ra, "ref", false, syn.SpanStart); }
                else Gen(sym, ra, "out", true, -1);
                return;
            }
            if (isDecl) return;
            Use(sym, syn);
        }

        static SyntaxNode DefSyntax(ISimpleAssignmentOperation a) => a.Syntax;
        static string KindOf(SyntaxNode s)
        {
            for (var p = s; p != null; p = p.Parent)
            {
                if (p is ForEachStatementSyntax) return "foreach";
                if (p is CatchDeclarationSyntax || p is CatchClauseSyntax) return "catch";
                if (p is StatementSyntax) break;
            }
            return "set";
        }
    }
}
