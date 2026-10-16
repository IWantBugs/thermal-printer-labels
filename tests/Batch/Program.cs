using LabelCore;
if (BatchQuantity.Calculate("2",100,0)!=200 || BatchQuantity.Calculate("2",100,10)!=220 || BatchQuantity.Calculate("1",3,10)!=4 || BatchQuantity.Calculate("0,5",3,0)!=2 || BatchQuantity.Calculate("bad",100,10)!=null || BatchQuantity.Calculate("-1",100,10)!=null || BatchQuantity.Calculate("79228162514264337593543950335",100,10)!=null) throw new Exception("Batch calculation failed");
var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
try {
 BomExport.Save(path, new[]{new BomExportRow(new[]{"1","R1","PART","2","220","desc","100"},"FFC6EFCE")},new[]{"Line #","Designator","PartNumber","Quantity","Общее + 10% (100 плат)","Summary","Упк, шт"});
 var data = Bom.Read(path)[0].Rows;
 if(data[0].Length!=7 || data[1][4]!="220" || data[1][3]!="2") throw new Exception("Batch export failed");
} finally { File.Delete(path); }
Console.WriteLine("PASS: batch calculation, reserve, rounding, invalid values, overflow and seven-column Excel export");
