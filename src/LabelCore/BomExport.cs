using System.IO.Compression;
using System.Xml.Linq;
namespace LabelCore;

public record BomExportRow(string[] Values, string? Fill);
public static class BomExport
{
    public static void Save(string path, IReadOnlyList<BomExportRow> rows)
    {
        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace p = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        string[] colors = ["FFC6EFCE", "FFFFEB9C", "FFFFC7CE"];
        XElement Fill(string color) => new(s + "fill", new XElement(s + "patternFill", new XAttribute("patternType", "solid"), new XElement(s + "fgColor", new XAttribute("rgb", color)), new XElement(s + "bgColor", new XAttribute("indexed", 64))));
        XElement Xf(int fill, int font = 0) => new(s + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", font), new XAttribute("fillId", fill), new XAttribute("borderId", 0), new XAttribute("xfId", 0), new XAttribute("applyFill", 1));
        var styles = new XElement(s + "styleSheet",
            new XElement(s + "fonts", new XAttribute("count", 2), new XElement(s + "font", new XElement(s + "sz", new XAttribute("val", 11)), new XElement(s + "name", new XAttribute("val", "Calibri"))), new XElement(s + "font", new XElement(s + "b"), new XElement(s + "sz", new XAttribute("val", 11)), new XElement(s + "name", new XAttribute("val", "Calibri")))),
            new XElement(s + "fills", new XAttribute("count", 5), new XElement(s + "fill", new XElement(s + "patternFill", new XAttribute("patternType", "none"))), new XElement(s + "fill", new XElement(s + "patternFill", new XAttribute("patternType", "gray125"))), colors.Select(Fill)),
            new XElement(s + "borders", new XAttribute("count", 1), new XElement(s + "border", new XElement(s + "left"), new XElement(s + "right"), new XElement(s + "top"), new XElement(s + "bottom"), new XElement(s + "diagonal"))),
            new XElement(s + "cellStyleXfs", new XAttribute("count", 1), Xf(0)),
            new XElement(s + "cellXfs", new XAttribute("count", 5), Xf(0), Xf(2), Xf(3), Xf(4), Xf(0, 1)),
            new XElement(s + "cellStyles", new XAttribute("count", 1), new XElement(s + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0))));
        XElement Row(int number, string[] values, int style) => new(s + "row", new XAttribute("r", number), values.Select((v, i) => new XElement(s + "c", new XAttribute("r", Bom.ColumnName(i) + number), new XAttribute("t", "inlineStr"), new XAttribute("s", style), new XElement(s + "is", new XElement(s + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), v)))));
        var data = new XElement(s + "sheetData", Row(1, ["Line #", "Designator", "PartNumber", "Quantity", "Summary", "Упк, шт"], 4));
        for (int i = 0; i < rows.Count; i++) data.Add(Row(i + 2, rows[i].Values, Array.IndexOf(colors, rows[i].Fill) + 1));
        var sheet = new XElement(s + "worksheet", new XElement(s + "sheetViews", new XElement(s + "sheetView", new XAttribute("workbookViewId", 0), new XElement(s + "pane", new XAttribute("ySplit", 1), new XAttribute("topLeftCell", "A2"), new XAttribute("state", "frozen")))), new XElement(s + "cols", Enumerable.Range(1, 6).Select(i => new XElement(s + "col", new XAttribute("min", i), new XAttribute("max", i), new XAttribute("width", i is 3 or 5 ? 38 : 20), new XAttribute("customWidth", 1)))), data, new XElement(s + "autoFilter", new XAttribute("ref", "A1:F" + (rows.Count + 1))));
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Write(string name, XElement xml) { using var stream = zip.CreateEntry(name).Open(); new XDocument(xml).Save(stream); }
        Write("[Content_Types].xml", new XElement(ct + "Types", new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")), new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")), new[] { ("/xl/workbook.xml", "sheet.main"), ("/xl/worksheets/sheet1.xml", "worksheet"), ("/xl/styles.xml", "styles") }.Select(x => new XElement(ct + "Override", new XAttribute("PartName", x.Item1), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml." + x.Item2 + "+xml")))));
        Write("_rels/.rels", new XElement(p + "Relationships", new XElement(p + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", r.NamespaceName + "/officeDocument"), new XAttribute("Target", "xl/workbook.xml"))));
        Write("xl/workbook.xml", new XElement(s + "workbook", new XAttribute(XNamespace.Xmlns + "r", r), new XElement(s + "sheets", new XElement(s + "sheet", new XAttribute("name", "BOM"), new XAttribute("sheetId", 1), new XAttribute(r + "id", "rId1")))));
        Write("xl/_rels/workbook.xml.rels", new XElement(p + "Relationships", new XElement(p + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", r.NamespaceName + "/worksheet"), new XAttribute("Target", "worksheets/sheet1.xml")), new XElement(p + "Relationship", new XAttribute("Id", "rId2"), new XAttribute("Type", r.NamespaceName + "/styles"), new XAttribute("Target", "styles.xml"))));
        Write("xl/styles.xml", styles); Write("xl/worksheets/sheet1.xml", sheet);
    }
}
