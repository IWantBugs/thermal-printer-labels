using System.Drawing.Printing;
using System.Drawing.Drawing2D;
using LabelCore;
using BomLabel=LabelCore.Label;
namespace LabelApp;
internal sealed class NativePrintForm : Form
{
    readonly PrintDocument document=new();
    readonly PictureBox picture=new() {Dock=DockStyle.Fill,SizeMode=PictureBoxSizeMode.Zoom,BackColor=Color.White};
    readonly ComboBox printers=new() {Width=280,DropDownStyle=ComboBoxStyle.DropDownList};
    readonly NumericUpDown page=new() {Minimum=1,Width=65};
    readonly NumericUpDown copies=new() {Minimum=1,Maximum=999,Value=1,Width=65};
    readonly NumericUpDown x=new() {Value=0.5m,Minimum=-10,Maximum=10,DecimalPlaces=1,Increment=0.1m,Width=65};
    readonly NumericUpDown y=new() {Minimum=-10,Maximum=10,DecimalPlaces=1,Increment=0.1m,Width=65};
    readonly Func<int,Bitmap> render;
    int printIndex,printEnd;
    public NativePrintForm(string product,IReadOnlyList<BomLabel> labels)
        : this(labels.Count,i=>NativeLabelRenderer.Render(product,labels[i]),"SPK · Печать без Word · 58×40 мм") { }
    public NativePrintForm(string product,int boards)
        : this(1,_=>NativeLabelRenderer.RenderOrder(product,boards),"SPK · Этикетка заказа · 58×40 мм") { }
    NativePrintForm(int count,Func<int,Bitmap> render,string title)
    {
        this.render=render;
        Text=title;Size=new Size(1000,700);MinimumSize=new Size(800,500);StartPosition=FormStartPosition.CenterParent;
        page.Maximum=count;
        foreach(string name in PrinterSettings.InstalledPrinters) printers.Items.Add(name);
        printers.SelectedItem=document.PrinterSettings.PrinterName;
        var xp=printers.Items.Cast<string>().FirstOrDefault(n=>n.Contains("XP-365",StringComparison.OrdinalIgnoreCase));
        if(xp!=null) printers.SelectedItem=xp;
        var toolbar=new FlowLayoutPanel {Dock=DockStyle.Top,AutoSize=true,Padding=new Padding(10),WrapContents=true};
        void Field(string title,Control control) {toolbar.Controls.Add(new System.Windows.Forms.Label {Text=title,AutoSize=true,Padding=new Padding(0,5,0,0)});toolbar.Controls.Add(control);}
        Field("Принтер",printers);Field("Этикетка",page);Field($"из {count} · копий",copies);
        Field("Сдвиг X, мм",x);Field("Y, мм",y);
        var settings=new Button {Text="Свойства принтера…",AutoSize=true};
        settings.Click+=(_,_)=>Run(()=> {SelectPrinter();using var setup=new PrintDialog {Document=document,UseEXDialog=true};setup.ShowDialog(this);printers.SelectedItem=document.PrinterSettings.PrinterName;copies.Value=Math.Clamp((int)document.PrinterSettings.Copies,1,999);});
        var print=new Button {Text="Печать…",AutoSize=true};
        print.Click+=(_,_)=>Run(()=>
        {
            SelectPrinter();ConfigurePaper();
            document.PrinterSettings.Copies=(short)copies.Value;
            document.PrinterSettings.MinimumPage=1;document.PrinterSettings.MaximumPage=count;
            document.PrinterSettings.FromPage=1;document.PrinterSettings.ToPage=count;
            using var dialog=new PrintDialog {Document=document,AllowSomePages=true,UseEXDialog=true};
            if(dialog.ShowDialog(this)!=DialogResult.OK)return;
            printers.SelectedItem=document.PrinterSettings.PrinterName;
            copies.Value=Math.Clamp((int)document.PrinterSettings.Copies,1,999);
            ConfigurePaper();
            int start=document.PrinterSettings.PrintRange==PrintRange.SomePages?document.PrinterSettings.FromPage-1:0;
            int end=document.PrinterSettings.PrintRange==PrintRange.SomePages?document.PrinterSettings.ToPage:count;
            if(start<0 || end>count || end<=start) throw new InvalidOperationException("Некорректный диапазон этикеток.");
            for(int i=start;i<end;i++) {using var check=RenderPage(i);}
            document.Print();
        });
        toolbar.Controls.AddRange([settings,print]);
        var hint=new System.Windows.Forms.Label {Dock=DockStyle.Bottom,Height=55,Padding=new Padding(12),Text="Предпросмотр увеличен для просмотра. Печать: 58×40 мм, 100%. В драйвере выберите этикетку 58×40 мм. Начните с одной этикетки; сдвиг X/Y виден в предпросмотре."};
        page.ValueChanged+=(_,_)=>Run(Preview); x.ValueChanged+=(_,_)=>Run(Preview); y.ValueChanged+=(_,_)=>Run(Preview);
        document.BeginPrint+=(_,_)=> {printIndex=document.PrinterSettings.PrintRange==PrintRange.SomePages?document.PrinterSettings.FromPage-1:0;printEnd=document.PrinterSettings.PrintRange==PrintRange.SomePages?document.PrinterSettings.ToPage:count;};
        document.PrintPage+=(_,e)=>
        {
            using var bitmap=RenderPage(printIndex);
            var g=e.Graphics!;g.PageUnit=GraphicsUnit.Display;
            g.TranslateTransform(-e.PageSettings.HardMarginX,-e.PageSettings.HardMarginY);
            g.InterpolationMode=InterpolationMode.NearestNeighbor;g.PixelOffsetMode=PixelOffsetMode.Half;
            g.DrawImage(bitmap,new RectangleF(0,0,58/25.4f*100,40/25.4f*100));
            e.HasMorePages=++printIndex<printEnd;
        };
        Controls.Add(picture);Controls.Add(toolbar);Controls.Add(hint);Appearance.Dialog(this);
        FormClosed+=(_,_)=>{picture.Image?.Dispose();document.Dispose();};
        Preview();
    }
    void SelectPrinter()
    {
        if(printers.SelectedItem is not string name)throw new InvalidOperationException("Выберите установленный принтер Windows.");
        document.PrinterSettings.PrinterName=name;
        if(!document.PrinterSettings.IsValid)throw new InvalidOperationException("Принтер недоступен.");
    }
    void ConfigurePaper()
    {
        var paper=document.PrinterSettings.PaperSizes.Cast<PaperSize>().FirstOrDefault(p=>Math.Abs(p.Width-58/25.4*100)<2 && Math.Abs(p.Height-40/25.4*100)<2);
        document.DefaultPageSettings.PaperSize=paper??new PaperSize("SPK 58x40 mm",(int)Math.Round(58/25.4*100),(int)Math.Round(40/25.4*100));
        document.DefaultPageSettings.Landscape=false;document.DefaultPageSettings.Margins=new Margins(0,0,0,0);document.OriginAtMargins=false;
    }
    Bitmap RenderPage(int index)
    {
        var original=render(index);
        if(x.Value==0 && y.Value==0) return original;
        using(original)
        {
            var shifted=new Bitmap(464,320);shifted.SetResolution(203.2f,203.2f);
            using var g=Graphics.FromImage(shifted);g.Clear(Color.White);
            g.DrawImageUnscaled(original,(int)Math.Round(x.Value*8),(int)Math.Round(y.Value*8));return shifted;
        }
    }
    void Preview(){var next=RenderPage((int)page.Value-1);var old=picture.Image;picture.Image=next;old?.Dispose();}
    void Run(Action action){try{action();}catch(Exception ex){MessageBox.Show(this,ex.Message,"Печать",MessageBoxButtons.OK,MessageBoxIcon.Error);}}
}
