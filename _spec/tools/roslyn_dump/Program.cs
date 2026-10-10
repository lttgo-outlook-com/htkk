// Chỉ đọc text .cs và gọi CSharpSyntaxTree.ParseText. Không biên dịch, không workspace, không chạy mã decompile.
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

// Dùng: RoslynDump <decompiledDir> <menu.xml> <outDir>
if (args.Length < 3) { Console.Error.WriteLine("Dùng: RoslynDump <decompiledDir> <menu.xml> <outDir>"); return 2; }
string decDir = Path.GetFullPath(args[0]);
string menuPath = args[1];
string outDir = args[2];
Directory.CreateDirectory(outDir);

var hookRx = new Regex(@"^(CellChange|AutoBindData|Refresh)_(\d+)$");
string HookKind(string n)
{
    var m = hookRx.Match(n);
    if (m.Success) return m.Groups[1].Value;
    if (n == "UpdateFormula" || n == "ClearData" || n == "CellEditchange" || n == "ButtonClick") return n;
    if (n.StartsWith("Validate")) return "Validate";
    return null;
}
var sheetRx = new Regex(@"_(\d+)$");

var files = Directory.EnumerateFiles(decDir, "*.cs", SearchOption.AllDirectories)
    .Select(p => Path.GetRelativePath(decDir, p).Replace('\\', '/'))
    .OrderBy(p => p, StringComparer.Ordinal).ToList();

var fileRecs = new List<object>();
var allClasses = new List<(string asm, string rel, string ns, string cls, string baseCls)>();
int filesWithErr = 0, totalErr = 0;
var hookCount = new SortedDictionary<string, int>(StringComparer.Ordinal);
int methodCount = 0;

foreach (var rel in files)
{
    string abs = Path.Combine(decDir, rel);
    byte[] bytes = File.ReadAllBytes(abs);
    string sha = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    string text = File.ReadAllText(abs); // BOM-aware
    var tree = CSharpSyntaxTree.ParseText(text, new CSharpParseOptions(LanguageVersion.Latest));
    var diags = tree.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error).ToList();
    if (diags.Count > 0) { filesWithErr++; totalErr += diags.Count; }
    var root = tree.GetCompilationUnitRoot();
    string asm = rel.Split('/')[0];
    var nsList = new List<string>();
    var classNames = new List<string>();
    var bases = new List<string>();
    var methods = new List<object>();

    foreach (var td in root.DescendantNodes().OfType<TypeDeclarationSyntax>())
    {
        // tên lồng nhau
        var parts = new List<string> { td.Identifier.Text };
        for (var p = td.Parent; p != null; p = p.Parent)
            if (p is TypeDeclarationSyntax pt) parts.Insert(0, pt.Identifier.Text);
        string cname = string.Join(".", parts);
        string ns = "";
        for (var p = td.Parent; p != null; p = p.Parent)
            if (p is BaseNamespaceDeclarationSyntax bn) { ns = bn.Name.ToString(); break; }
        if (!nsList.Contains(ns)) nsList.Add(ns);
        string baseCls = td.BaseList?.Types.FirstOrDefault()?.Type.ToString();
        classNames.Add(cname);
        bases.Add(baseCls);
        if (td.Parent is not TypeDeclarationSyntax) allClasses.Add((asm, rel, ns, cname, baseCls));

        foreach (var md in td.Members.OfType<BaseMethodDeclarationSyntax>())
        {
            string name = md is MethodDeclarationSyntax mm ? mm.Identifier.Text
                : md is ConstructorDeclarationSyntax cc ? ".ctor"
                : md is DestructorDeclarationSyntax ? ".dtor"
                : md is OperatorDeclarationSyntax ? "operator"
                : md is ConversionOperatorDeclarationSyntax ? "conversion" : "?";
            var span = md.GetLocation().GetLineSpan();
            int sl = span.StartLinePosition.Line + 1, el = span.EndLinePosition.Line + 1;
            int bodyLines = 0;
            if (md.Body != null)
            {
                var bs = md.Body.GetLocation().GetLineSpan();
                bodyLines = bs.EndLinePosition.Line - bs.StartLinePosition.Line + 1;
            }
            else if (md.ExpressionBody != null)
            {
                var bs = md.ExpressionBody.GetLocation().GetLineSpan();
                bodyLines = bs.EndLinePosition.Line - bs.StartLinePosition.Line + 1;
            }
            string kind = HookKind(name);
            int? sheet = null;
            if (kind != null || true)
            {
                var sm = sheetRx.Match(name);
                if (sm.Success) sheet = int.Parse(sm.Groups[1].Value);
            }
            string pars = string.Join(", ", md.ParameterList.Parameters.Select(p => (p.Type?.ToString() + " " + p.Identifier.Text).Trim()));
            methods.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
            {
                ["class"] = cname, ["name"] = name, ["params"] = pars,
                ["start_line"] = sl, ["end_line"] = el, ["body_lines"] = bodyLines,
                ["is_hook"] = kind != null, ["hook_kind"] = kind, ["sheet_index"] = sheet,
            });
            methodCount++;
            if (kind != null) hookCount[kind] = hookCount.GetValueOrDefault(kind) + 1;
        }
    }

    fileRecs.Add(new SortedDictionary<string, object>(StringComparer.Ordinal)
    {
        ["path"] = rel, ["sha256"] = sha, ["assembly"] = asm,
        ["namespace"] = nsList.Count == 1 ? nsList[0] : (nsList.Count == 0 ? "" : string.Join(";", nsList)),
        ["class"] = classNames, ["base_class"] = bases.FirstOrDefault(b => b != null),
        ["parse_errors"] = diags.Count,
        ["diagnostics"] = diags.Select(d => { var l = d.Location.GetLineSpan().StartLinePosition; return $"{d.Id} {l.Line + 1}:{l.Character + 1} {d.GetMessage()}"; }).Take(20).ToList(),
        ["methods"] = methods,
    });
}

var jo = new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
var index = new SortedDictionary<string, object>(StringComparer.Ordinal)
{
    ["file_count"] = files.Count, ["files_with_errors"] = filesWithErr, ["total_errors"] = totalErr,
    ["method_count"] = methodCount, ["hook_counts"] = hookCount, ["files"] = fileRecs,
};
File.WriteAllText(Path.Combine(outDir, "index.json"), JsonSerializer.Serialize(index, jo) + "\n", new UTF8Encoding(false));

// classes.json: class nghiệp vụ = lớp cấp cao trong assembly tax.htkk.business, namespace TAX.HTKK.Business
// menu: Class + "Nckt" + VersionXML bỏ dấu chấm = tên lớp
var menu = XDocument.Load(menuPath);
var menuMap = new Dictionary<string, (string id, string maFile, string ver, string cls)>();
foreach (var m in menu.Descendants("Menu"))
{
    string c = (string)m.Attribute("Class"); if (string.IsNullOrEmpty(c)) continue;
    string v = (string)m.Attribute("VersionXML") ?? "";
    string key = c + "Nckt" + v.Replace(".", "");
    if (!menuMap.ContainsKey(key)) menuMap[key] = ((string)m.Attribute("ID"), (string)m.Attribute("MaFile"), v, c);
}
var biz = allClasses.Where(c => c.asm == "tax.htkk.business" && c.ns == "TAX.HTKK.Business")
    .OrderBy(c => c.cls, StringComparer.Ordinal).ToList();
var classRecs = biz.Select(c =>
{
    bool inMenu = menuMap.TryGetValue(c.cls, out var mm);
    return new SortedDictionary<string, object>(StringComparer.Ordinal)
    {
        ["class"] = c.cls, ["path"] = c.rel, ["base_class"] = c.baseCls, ["in_menu"] = inMenu,
        ["MenuID"] = inMenu ? mm.id : null, ["MaFile"] = inMenu ? mm.maFile : null, ["VersionXML"] = inMenu ? mm.ver : null,
        ["MenuClass"] = inMenu ? mm.cls : null,
    };
}).ToList();
var classes = new SortedDictionary<string, object>(StringComparer.Ordinal)
{
    ["class_count"] = classRecs.Count, ["in_menu_count"] = classRecs.Count(r => (bool)r["in_menu"]),
    ["menu_entries_with_class"] = menuMap.Count, ["classes"] = classRecs,
};
File.WriteAllText(Path.Combine(outDir, "classes.json"), JsonSerializer.Serialize(classes, jo) + "\n", new UTF8Encoding(false));

Console.WriteLine($"files_parsed={files.Count} files_with_errors={filesWithErr} total_errors={totalErr}");
Console.WriteLine($"business_classes={classRecs.Count} in_menu={classRecs.Count(r => (bool)r["in_menu"])}");
Console.WriteLine($"methods={methodCount}");
foreach (var kv in hookCount) Console.WriteLine($"hook {kv.Key}={kv.Value}");
return 0;
