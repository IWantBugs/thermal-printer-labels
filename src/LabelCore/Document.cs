using System.IO.Compression;
using System.Xml.Linq;
namespace LabelCore;

public static class Document
{
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static XElement E(string name, params object[] content) => new(W + name, content);
    static XAttribute A(string name, object value) => new(W + name, value);

    static XElement Cell(string text, int width, bool heading = false)
    {
        var margins = E("tcMar");
        foreach (var side in new[] { "top", "left", "bottom", "right" })
            margins.Add(E(side, A("w", 20), A("type", "dxa")));
        var properties = E("tcPr", E("tcW", A("w", width), A("type", "dxa")), margins, E("vAlign", A("val", "center")));
        var paragraphProperties = E("pPr", E("spacing", A("before", 0), A("after", 0), A("line", 180), A("lineRule", "exact")));
        var runProperties = E("rPr", E("rFonts", A("ascii", "Arial"), A("hAnsi", "Arial")), E("sz", A("val", 16)));
        if (heading) runProperties.Element(W + "rFonts")!.AddAfterSelf(E("b"));
        var run = E("r", runProperties, E("t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        return E("tc", properties, E("p", paragraphProperties, run));
    }

    public static string Heading(string product, string line) => $"Изделие: {product} [{line}]";

    public static void Save(string path, string product, IReadOnlyList<Label> labels)
    {
        if (labels.Count == 0) throw new ArgumentException("Нет этикеток");
        var body = E("body");
        for (int i = 0; i < labels.Count; i++)
        {
            if (i > 0)
            {
                var spacing = E("spacing", A("before", 0), A("after", 0), A("line", 20), A("lineRule", "exact"));
                body.Add(E("p", E("pPr", spacing), E("r", E("rPr", E("sz", A("val", 2))), E("br", A("type", "page")))));
            }
            var borders = E("tblBorders");
            foreach (var side in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
                borders.Add(E(side, A("val", "single"), A("sz", 4), A("color", "000000")));
            var properties = E("tblPr", E("tblW", A("w", 3175), A("type", "dxa")), borders, E("tblLayout", A("type", "fixed")));
            var table = E("tbl", properties, E("tblGrid", E("gridCol", A("w", 850)), E("gridCol", A("w", 2325))));
            var label = labels[i];
            var rows = new List<(string Key, string Value)>
            {
                ("Изд", Heading(product, label.Line)),
                ("Обозн.", label.Designator), ("Парт", label.Part),
                ("Кол-во", label.Quantity), ("Sum", label.Summary)
            };
            if (label.Package is not null) rows.Add(("Упк, шт", label.Package));
            foreach (var (key, value) in rows)
            {
                int height = key == "Sum" ? 567 : key == "Изд" ? 397 : 227;
                var rowProperties = E("trPr", E("cantSplit"), E("trHeight", A("val", height), A("hRule", key == "Изд" ? "atLeast" : "exact")));
                table.Add(E("tr", rowProperties, Cell(key, 850, key == "Изд"), Cell(value, 2325, key == "Изд")));
            }
            body.Add(table);
        }
        // Word requires a paragraph after the final table. Without explicit
        // formatting it inserts a Normal paragraph that can overflow a label.
        var finalParagraphProperties = E("pPr",
            E("keepNext", A("val", 0)), E("keepLines", A("val", 0)),
            E("pageBreakBefore", A("val", 0)), E("widowControl", A("val", 0)),
            E("spacing", A("before", 0), A("after", 0), A("line", 20), A("lineRule", "exact")),
            E("rPr", E("sz", A("val", 2)), E("szCs", A("val", 2))));
        body.Add(E("p", finalParagraphProperties));
        var page = E("pgSz", A("w", 3288), A("h", 2268));
        var marginsPage = E("pgMar", A("top", 114), A("bottom", 114), A("left", 57), A("right", 57), A("header", 0), A("footer", 0));
        body.Add(E("sectPr", page, marginsPage));
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var zip = ZipFile.Open(temporary, ZipArchiveMode.Create))
            {
                void Put(string name, string value)
                {
                    using var writer = new StreamWriter(zip.CreateEntry(name).Open());
                    writer.Write(value);
                }
                Put("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
                Put("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
                Put("word/document.xml", new XDocument(E("document", body)).ToString());
            }
            File.Move(temporary, path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
