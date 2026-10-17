using LabelCore;
if (BatchQuantity.Calculate("2",100,0)!=200 || BatchQuantity.Calculate("2",100,10)!=220 || BatchQuantity.Calculate("1",3,10)!=4 || BatchQuantity.Calculate("0,5",3,0)!=2 || BatchQuantity.Calculate("bad",100,10)!=null || BatchQuantity.Calculate("-1",100,10)!=null || BatchQuantity.Calculate("79228162514264337593543950335",100,10)!=null) throw new Exception("Batch calculation failed");
var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
try {
 BomExport.Save(path, new[]{new BomExportRow(new[]{"1","R1","PART","2","220","desc","100"},"FFC6EFCE")},new[]{"Line #","Designator","PartNumber","Quantity","Общее + 10% (100 плат)","Summary","Упк, шт"});
 var data = Bom.Read(path)[0].Rows;
 if(data[0].Length!=7 || data[1][4]!="220" || data[1][3]!="2") throw new Exception("Batch export failed");
} finally { File.Delete(path); }
Console.WriteLine("PASS: batch calculation, reserve, rounding, invalid values, overflow and seven-column Excel export");

var rules = new ReserveRules();
rules.Validate();
foreach(var (designation, expected) in new (string, decimal)[] { ("R1,R2",10), ("C1-C5",10), ("VD1",10), ("VT2",10), ("Q1",10), ("q5",10), ("U1",5), ("DA1",5), ("DD2",5), ("XS1",3), ("XP1",3), ("L1",3), ("r10",10), ("UR1",3) })
 if(rules.PercentFor(designation)!=expected) throw new Exception("Reserve group failed: " + designation);
rules.Groups.Add(new("L",7)); rules.Validate();
if(rules.PercentFor("L1")!=7) throw new Exception("Custom group failed");
Console.WriteLine("PASS: default groups, exact prefixes, lowercase, ranges and custom group");

var old = new ReserveRules { Groups = [new("R, C, VD, VT",10), new("U",5)] };
old.Upgrade(); old.Validate();
if(old.PercentFor("Q1")!=10) throw new Exception("Q upgrade failed");
old.Groups.Add(new("L",7));
old.Groups.RemoveAll(g=>ReserveRules.Prefixes(g.Prefixes).Contains("Q"));
old.Upgrade();
if(old.PercentFor("Q1")!=3) throw new Exception("Repeated upgrade changed user edits");
Console.WriteLine("PASS: Q transistor and one-time settings upgrade");

var sessionFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid()+".xlsx");
try {
 BomExport.Save(sessionFile, new[]{new BomExportRow(new[]{"9","Q1","P","2","",""},"FFFFC7CE")});
 var original = new BomSession { Product="Board",Boards=100,Reserve=true,Package=false,Widths=new(){{"Part",135}} };
 original.Rules.Groups.Add(new("L",7));
 BomSession.Write(sessionFile, original);
 var restored = BomSession.Read(sessionFile) ?? throw new Exception("Session missing");
 if(restored.Product!="Board" || restored.Boards!=100 || !restored.Reserve || restored.Package || restored.Rules.PercentFor("L1")!=7 || restored.Widths["Part"]!=135) throw new Exception("Session state mismatch");
 if(BomSession.Colors(sessionFile)[0]!="FFFFC7CE" || Bom.Read(sessionFile)[0].Rows[1][0]!="9") throw new Exception("Session data or colors mismatch");
} finally { File.Delete(sessionFile); }
Console.WriteLine("PASS: edited BOM session, row color, product, boards, flags, groups and column widths");

if(!PackagePieces.Parse("150\n150").SequenceEqual(new[]{150,150}) || PackagePieces.Parse("100+150+50").Sum()!=300 || PackagePieces.Parse("").Length!=0) throw new Exception("Package pieces failed");
try {PackagePieces.Parse("0");throw new Exception("Invalid piece accepted");} catch(ArgumentException){}
Console.WriteLine("PASS: package pieces, totals, removal and validation");

var splitFile = Path.Combine(Path.GetTempPath(),Guid.NewGuid()+".xlsx");
try {
 BomExport.Save(splitFile,new[]{new BomExportRow(new[]{"1","R1","P","3","","300","150","150"},"FFC6EFCE")},new[]{"Line #","Designator","PartNumber","Quantity","Summary","Упк, шт","Упк 1","Упк 2"});
 var saved = Bom.Read(splitFile)[0].Rows;
 if(saved[1][5]!="300" || saved[1][6]!="150" || saved[1][7]!="150") throw new Exception("Split export failed");
} finally {File.Delete(splitFile);}
var labelFile=Path.Combine(Path.GetTempPath(),Guid.NewGuid()+".docx");
try {
 Document.Save(labelFile,"Test",new[]{new LabelCore.Label("1","R1","P","3","","150"),new LabelCore.Label("1","R1","P","3","","150")});
 using var z=System.IO.Compression.ZipFile.OpenRead(labelFile);
 using var stream=z.GetEntry("word/document.xml")!.Open();
 var xml=System.Xml.Linq.XDocument.Load(stream);System.Xml.Linq.XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
 if(xml.Descendants(w+"tbl").Count()!=2 || xml.Descendants(w+"t").Count(t=>t.Value=="150")!=2) throw new Exception("Piece labels failed");
} finally {File.Delete(labelFile);}
Console.WriteLine("PASS: Excel split columns and two separate 150-piece labels");

if(!StockCheck.IsShort("150",300) || !StockCheck.IsShort("",300) || StockCheck.IsShort("300",300) || StockCheck.IsShort("350",300) || StockCheck.IsShort("bad",300) || StockCheck.IsShort("150",null) || !StockCheck.IsShort("0",300)) throw new Exception("Stock shortage detection failed");
Console.WriteLine("PASS: shortage detection, equality, empty packaging and absent totals");

if(StockCheck.Missing("150",300)!=150 || StockCheck.Missing("300",300)!=0 || StockCheck.Missing("350",300)!=0 || StockCheck.Missing("",300)!=300 || StockCheck.Missing("150",330)!=180) throw new Exception("Deficit difference failed");
var deficitPath=Path.Combine(Path.GetTempPath(),Guid.NewGuid()+".xlsx");
try {
 var stock=new[]{("1","150",300m),("2","300",300m),("3","",300m)};
 var deficits=stock.Where(r=>StockCheck.IsShort(r.Item2,r.Item3)).Select(r=>new BomExportRow(new[]{r.Item1,r.Item2,StockCheck.Missing(r.Item2,r.Item3)!.Value.ToString()},null)).ToArray();
 BomExport.Save(deficitPath,deficits,new[]{"Line #","Упк, шт","Не хватает, шт"});
 var result=Bom.Read(deficitPath)[0].Rows;
 if(result.Count!=3 || result[1][2]!="150" || result[2][2]!="300") throw new Exception("Deficit export failed");
} finally {File.Delete(deficitPath);}
Console.WriteLine("PASS: deficit difference and filtered Excel export");

var notesFile=Path.Combine(Path.GetTempPath(),Guid.NewGuid()+".xlsx");
try {
 BomExport.Save(notesFile,new[]{new BomExportRow(new[]{"1","R1","P","2","","100","Давальческая позиция; Проверить"},"FFDDEBF7")},new[]{"Line #","Designator","PartNumber","Quantity","Summary","Упк, шт","Примечание"});
 if(Bom.Read(notesFile)[0].Rows[1][6]!="Давальческая позиция; Проверить" || BomSession.Colors(notesFile)[0]!="FFDDEBF7") throw new Exception("Notes or customer color lost");
} finally {File.Delete(notesFile);}
Console.WriteLine("PASS: notes and light-blue customer row Excel roundtrip");
