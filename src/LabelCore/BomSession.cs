using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using System.Text.RegularExpressions;
namespace LabelCore;
public sealed class BomSession
{
    public string Product { get; set; } = "";
    public int Boards { get; set; }
    public bool Reserve { get; set; }
    public bool Package { get; set; }
    public ReserveRules Rules { get; set; } = new();
    public Dictionary<string, float> Widths { get; set; } = new();
    public static void Write(string path, BomSession state)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        using (var stream = zip.CreateEntry("customXml/item1.xml").Open()) new XDocument(new XElement("SpkBomSession", JsonSerializer.Serialize(state))).Save(stream);
        var entry = zip.GetEntry("xl/_rels/workbook.xml.rels")!;
        XDocument relations; using(var stream = entry.Open()) relations = XDocument.Load(stream);
        XNamespace p = "http://schemas.openxmlformats.org/package/2006/relationships";
        relations.Root!.Add(new XElement(p + "Relationship", new XAttribute("Id", "spkSession"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"), new XAttribute("Target", "../customXml/item1.xml")));
        entry.Delete(); using var output = zip.CreateEntry("xl/_rels/workbook.xml.rels").Open(); relations.Save(output);
    }
    public static BomSession? Read(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        var entry = zip.GetEntry("customXml/item1.xml");
        if (entry == null) return null;
        using var stream = entry.Open(); var xml = XDocument.Load(stream);
        if (xml.Root?.Name.LocalName != "SpkBomSession") return null;
        var result = JsonSerializer.Deserialize<BomSession>(xml.Root.Value) ?? throw new InvalidDataException("Нет настроек редакции BOM.");
        result.Rules.Validate();
        if (result.Boards < 0 || result.Boards > 1000000000) throw new InvalidDataException("Некорректное число плат в редакции BOM.");
        return result;
    }
    public static string?[] Colors(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        XDocument ReadXml(string part) { using var stream = zip.GetEntry(part)!.Open(); return XDocument.Load(stream); }
        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var styles = ReadXml("xl/styles.xml");
        var fills = styles.Root!.Element(s + "fills")!.Elements().Select(f => (string?)f.Descendants(s + "fgColor").FirstOrDefault()?.Attribute("rgb")).ToArray();
        var formats = styles.Root.Element(s + "cellXfs")!.Elements().Select(x => (int?)x.Attribute("fillId") ?? 0).ToArray();
        return ReadXml("xl/worksheets/sheet1.xml").Descendants(s + "row").Skip(1).Select(row =>
        {
            var cell = row.Element(s + "c"); int style = (int?)cell?.Attribute("s") ?? 0;
            int fill = style < formats.Length ? formats[style] : 0;
            return fill < fills.Length ? fills[fill] : null;
        }).ToArray();
    }
    public static BomSession Legacy(string path, string[] headers)
    {
        var state = new BomSession { Product = Bom.Product(Path.GetFileNameWithoutExtension(path).Replace("_EDIT", "", StringComparison.OrdinalIgnoreCase)) };
        var total = headers.FirstOrDefault(h => h.StartsWith("Общее", StringComparison.OrdinalIgnoreCase)) ?? "";
        var count = Regex.Match(total, @"\((\d+) плат\)");
        if (count.Success && int.TryParse(count.Groups[1].Value, out int boards)) state.Boards = boards;
        state.Reserve = total.Contains("техзапас") || total.Contains('+');
        return state;
    }
}
