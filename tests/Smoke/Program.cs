using LabelCore;
using System.IO.Compression;
using System.Xml.Linq;
void Check(bool result, string message) { if (!result) throw new Exception(message); Console.WriteLine("PASS: " + message); }
if (args.Length != 2) throw new ArgumentException("Usage: Smoke sample.xlsx output.docx");
var sheets = Bom.Read(args[0]); Check(sheets.Count == 1, "one sample sheet");
var rows = sheets[0].Rows; Check(rows.Count == 6, "five BOM positions plus header");
var header = rows[0]; string Field(string[] row, string name) => row[Array.IndexOf(header, name)];
var labels = rows.Skip(1).Select(r => new Label(Field(r, "Line #"), Field(r, "Designator"), Field(r, "PartNumber"), Field(r, "Quantity"), Field(r, "Name"), "5000")).ToArray();
Check(labels[0].Part == "FYLS-0805URC" && labels[3].Part == "HSSR-S1A05L-2", "part numbers and whitespace");
Check(labels.Count(l => l.Designator == "U1") == 1, "no duplicate alternative U1");
Check(Bom.Product(args[0]) == "RS7_ADP_V1", "product from source filename");
Check(Path.GetFileName(Bom.Output(args[0])) == "RS7_ADP_V1_ETC.docx", "ETC output preserves source product");
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
void Verify(string path, int rowsPerLabel) { using (var word = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false)) { var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word).ToArray(); Check(errors.Length == 0, "Open XML schema validation: " + string.Join("; ", errors.Select(e => e.Description))); } using var z = ZipFile.OpenRead(path); using var stream = z.GetEntry("word/document.xml")!.Open(); var doc = XDocument.Load(stream); var tables = doc.Descendants(w + "tbl").ToArray(); Check(tables.Length == 5, "five label tables");
for (int i = 0; i < tables.Length; i++) {
 var first = tables[i].Elements(w + "tr").First();
 var cells = first.Elements(w + "tc").Select(c => string.Concat(c.Descendants(w + "t").Select(t => t.Value))).ToArray();
 Check(cells[0] == "Изд" && cells[1] == $"Изделие: RS7_ADP_V1 [{labels[i].Line}]", "product heading and BOM Line #: " + labels[i].Line);
 Check((string?)first.Element(w + "trPr")?.Element(w + "trHeight")?.Attribute(w + "hRule") == "atLeast", "heading height can expand without clipping");
}
 Check(tables.All(t => t.Elements(w + "tr").Count() == rowsPerLabel), "optional package row count"); Check(doc.Descendants(w + "br").Count() == 4, "four explicit page breaks"); var tail = doc.Root!.Element(w + "body")!.Elements().Reverse().Skip(1).First();
Check(tail.Name == w + "p" && !tail.Descendants(w + "br").Any(), "explicit trailing paragraph without page break");
Check((string?)tail.Element(w + "pPr")?.Element(w + "spacing")?.Attribute(w + "line") == "20" && (string?)tail.Element(w + "pPr")?.Element(w + "rPr")?.Element(w + "sz")?.Attribute(w + "val") == "2", "trailing paragraph and paragraph mark limited to 1 pt");
var size = doc.Descendants(w + "pgSz").Single(); Check((string?)size.Attribute(w + "w") == "3288" && (string?)size.Attribute(w + "h") == "2268", "58x40 mm page dimensions"); Check(z.GetEntry("[Content_Types].xml") != null && z.GetEntry("_rels/.rels") != null, "DOCX package relationships"); }
Document.Save(args[1], "RS7_ADP_V1", labels); Verify(args[1], 6);
Document.Save(args[1], "RS7_ADP_V1", labels.Select(l => l with { Package = null }).ToArray()); Verify(args[1], 5);
Check(Path.GetFileName(Bom.Output("/tmp/Название-BOM.xlsx")) == "Название-ETC.docx", "source name and separator preserved");
Check(Bom.Product("/tmp/Название-BOM.xlsx") == "Название", "product suffix removal");
var reordered = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
try
{
    using (var input = ZipFile.OpenRead(args[0])) using (var copy = ZipFile.Open(reordered, ZipArchiveMode.Create))
    {
        foreach (var entry in input.Entries)
        {
            using var dest = copy.CreateEntry(entry.FullName).Open(); using var src = entry.Open();
            if (entry.FullName == "xl/worksheets/sheet1.xml")
            {
                XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"; var xml = XDocument.Load(src);
                foreach (var cell in xml.Descendants(s + "c")) { var r = (string)cell.Attribute("r")!; cell.SetAttributeValue("r", ((char)('G' - (r[0] - 'A'))) + r[1..]); }
                foreach (var row in xml.Descendants(s + "row")) { var cells = row.Elements(s + "c").OrderBy(c => (string?)c.Attribute("r")).ToArray(); row.ReplaceNodes(cells); }
                xml.Save(dest);
            }
            else src.CopyTo(dest);
        }
    }
    var reversed = Bom.Read(reordered)[0].Rows; var reversedHeader = reversed[0];
    Check(reversedHeader[0] == "PartNumber", "reordered header read by cell address");
    foreach (var key in new[] { "Line #", "Name", "Designator", "Quantity", "PartNumber" })
        Check(rows.Skip(1).Select(r => r[Array.IndexOf(header, key)]).SequenceEqual(reversed.Skip(1).Select(r => r[Array.IndexOf(reversedHeader, key)])), "reordered column values: " + key);
}
finally { if (File.Exists(reordered)) File.Delete(reordered); }
Console.WriteLine("All smoke checks passed.");
