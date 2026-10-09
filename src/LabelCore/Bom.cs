using System.IO.Compression;
using System.Xml.Linq;
namespace LabelCore;
public record Sheet(string Name, List<string[]> Rows);
public record Label(string Line, string Designator, string Part, string Quantity, string Summary, string? Package)
{
    public float DesignatorFont { get; init; } = 8;
    public int DesignatorLines { get; init; } = 1;
    public int SummaryLines { get; init; } = 3;
    public float PartFont { get; init; } = 8;
    public int PartLines { get; init; } = 1;
    public string? PartContinuation { get; init; }
    public string? DesignatorPart { get; init; }
}
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
    public static string[] Columns(Sheet sheet, int headerRow)
    {
        var headers = sheet.Rows[headerRow];
        return Enumerable.Range(0, sheet.Rows.Max(r => r.Length)).Select(c =>
        {
            var name = c < headers.Length && headers[c].Length > 0 ? headers[c] : "(без заголовка)";
            var samples = sheet.Rows.Skip(headerRow + 1).Take(3).Where(r => c < r.Length && r[c].Length > 0).Select(r => r[c]).Take(2);
            return ColumnName(c) + ": " + name + "  —  " + string.Join(" / ", samples);
        }).ToArray();
    }
    public static string ColumnName(int index)
    {
        var name = "";
        for (index++; index > 0; index = (index - 1) / 26) name = (char)('A' + (index - 1) % 26) + name;
        return name;
    }
    public static List<string[]> ProjectRows(Sheet sheet, int headerRow, int[] map) =>
        sheet.Rows.Skip(headerRow + 1).Select(row => map.Select(c => c >= 0 && c < row.Length ? row[c] : "").ToArray()).Where(row => row.Any(v => v.Length > 0)).ToList();
    public static string Product(string path) { var n = Path.GetFileNameWithoutExtension(path); return n.EndsWith("BOM", StringComparison.OrdinalIgnoreCase) ? n[..^3].TrimEnd('_', '-', ' ') : n; }
    public static string Output(string path) { var n = Path.GetFileNameWithoutExtension(path); return Path.Combine(Path.GetDirectoryName(path)!, (n.EndsWith("BOM", StringComparison.OrdinalIgnoreCase) ? n[..^3] + "ETC" : n + "_ETC") + ".docx"); }
}
