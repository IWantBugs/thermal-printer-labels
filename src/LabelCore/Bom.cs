using System.IO.Compression;
using System.Xml.Linq;
namespace LabelCore;
public record Sheet(string Name, List<string[]> Rows);
public record Label(string Line, string Designator, string Part, string Quantity, string Summary, string? Package);
public static class Bom
{
    static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public static List<Sheet> Read(string path)
    {
        using var z = ZipFile.OpenRead(path);
        XDocument Xml(string p) { using var s = z.GetEntry(p)?.Open() ?? throw new InvalidDataException("Нет части Excel: " + p); return XDocument.Load(s); }
        var strings = z.GetEntry("xl/sharedStrings.xml") is null ? [] : Xml("xl/sharedStrings.xml").Descendants(S + "si").Select(x => string.Concat(x.Descendants(S + "t").Select(t => t.Value))).ToArray();
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships", rid = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var links = Xml("xl/_rels/workbook.xml.rels").Descendants(rel + "Relationship").ToDictionary(x => (string)x.Attribute("Id")!, x => (string)x.Attribute("Target")!);
        var result = new List<Sheet>();
        foreach (var sh in Xml("xl/workbook.xml").Descendants(S + "sheet"))
        {
            var target = links[(string)sh.Attribute(rid + "id")!];
            var part = target.StartsWith('/') ? target.TrimStart('/') : new Uri(new Uri("http://local/xl/"), target).AbsolutePath.TrimStart('/');
            var rows = new List<string[]>();
            foreach (var row in Xml(part).Descendants(S + "sheetData").Elements(S + "row"))
            {
                var values = new SortedDictionary<int, string>();
                foreach (var c in row.Elements(S + "c"))
                {
                    int col = 0; foreach (char ch in ((string?)c.Attribute("r") ?? "").TakeWhile(char.IsLetter)) col = col * 26 + char.ToUpperInvariant(ch) - 'A' + 1;
                    if (col == 0) continue;
                    string value = c.Element(S + "v")?.Value ?? "";
                    if ((string?)c.Attribute("t") == "s") value = strings[int.Parse(value)];
                    if ((string?)c.Attribute("t") == "inlineStr") value = string.Concat(c.Descendants(S + "t").Select(t => t.Value));
                    values[col - 1] = value.Trim();
                }
                if (values.Count > 0 && values.Values.Any(v => v.Length > 0)) { var a = new string[values.Keys.Max() + 1]; Array.Fill(a, ""); foreach (var v in values) a[v.Key] = v.Value; rows.Add(a); }
            }
            if (rows.Count > 0) result.Add(new Sheet((string)sh.Attribute("name")!, rows));
        }
        return result;
    }
    public static string Product(string path) { var n = Path.GetFileNameWithoutExtension(path); return n.EndsWith("BOM", StringComparison.OrdinalIgnoreCase) ? n[..^3].TrimEnd('_', '-', ' ') : n; }
    public static string Output(string path) { var n = Path.GetFileNameWithoutExtension(path); return Path.Combine(Path.GetDirectoryName(path)!, (n.EndsWith("BOM", StringComparison.OrdinalIgnoreCase) ? n[..^3] + "ETC" : n + "_ETC") + ".docx"); }
}
