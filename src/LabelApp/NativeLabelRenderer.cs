using LabelCore;
using BomLabel = LabelCore.Label;
namespace LabelApp;
internal static class NativeLabelRenderer
{
    const float Dots = 8;
    public static bool Fits(string text, int lines, bool bold = false, float fontSize = 8)
    {
        using var image = new Bitmap(464,320); image.SetResolution(203.2f,203.2f);
        using var g = Graphics.FromImage(image); g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit; using var font = new Font("Arial",fontSize,bold?FontStyle.Bold:FontStyle.Regular);
        using var format = (StringFormat)StringFormat.GenericTypographic.Clone();
        var size=g.MeasureString(text,font,new SizeF(39*Dots,10000),format);
        float height = bold ? 8 : CriticalFields.Height(fontSize,lines);
        return size.Height <= height*Dots && size.Width<=39*Dots+0.5f;
    }
    public static Bitmap RenderOrder(string product,int boards)
    {
        var image=new Bitmap(464,320); image.SetResolution(203.2f,203.2f);
        try
        {
            using var g=Graphics.FromImage(image); g.Clear(Color.White);
            g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
            using var format=(StringFormat)StringFormat.GenericTypographic.Clone();
            format.Alignment=StringAlignment.Center; format.LineAlignment=StringAlignment.Center;
            void Draw(string text,RectangleF bounds,float maximum)
            {
                for(float size=maximum;size>=6;size-=0.5f)
                {
                    using var font=new Font("Arial",size,FontStyle.Bold);
                    var measured=g.MeasureString(text,font,new SizeF(bounds.Width,10000),format);
                    if(measured.Height<=bounds.Height && measured.Width<=bounds.Width+0.5f)
                    {g.DrawString(text,font,Brushes.Black,bounds,format);return;}
                }
                throw new InvalidDataException("Название заказа слишком длинное. Сократите поле «Изделие».");
            }
            Draw(product,boards>0?new RectangleF(16,16,424,88):new RectangleF(16,16,424,288),24);
            if(boards>0) Draw($"{boards} плат",new RectangleF(16,112,424,96),32);
            return image;
        }
        catch {image.Dispose();throw;}
    }
    public static Bitmap Render(string product, BomLabel label)
    {
        var image = new Bitmap(464,320); image.SetResolution(203.2f,203.2f);
        try
        {
            using var g=Graphics.FromImage(image); g.Clear(Color.White);
            g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
            using var pen=new Pen(Color.Black,1);
            var rows=new List<(string Key,string Value,float Height,float Font,bool Bold)>
            {
                ("Изд",Document.Heading(product,label.Line),8,8,true),
                ("Обозн."+(label.DesignatorPart==null?"":" "+label.DesignatorPart),label.Designator,Designators.Height(label),label.DesignatorFont,false),
                ("Парт"+(label.PartContinuation==null?"":" "+label.PartContinuation),label.Part,CriticalFields.Height(label.PartFont,label.PartLines),label.PartFont,false),
                ("Кол-во",label.Quantity,3.5f,8,false)
            };
            if(label.Summary.Length>0 && label.SummaryLines>0) rows.Add(("Sum",label.Summary,Designators.SummaryHeight(label),8,false));
            if(label.Package!=null) rows.Add(("Упк, шт",label.Package,3.5f,8,false));
            float y=2*Dots;
            foreach(var row in rows)
            {
                float h=row.Height*Dots;
                if(y+h>38*Dots) throw new InvalidDataException("Этикетка не помещается в 58×40 мм.");
                g.DrawRectangle(pen,8,y,448,h); g.DrawLine(pen,128,y,128,y+h);
                Draw(row.Key,new RectangleF(12,y,112,h),8,row.Bold,false);
                Draw(row.Value,new RectangleF(132,y,312,h),row.Font,row.Bold,row.Key!="Sum");
                y+=h;
            }
            void Draw(string value,RectangleF bounds,float size,bool bold,bool required)
            {
                using var font=new Font("Arial",size,bold?FontStyle.Bold:FontStyle.Regular);
                using var format=(StringFormat)StringFormat.GenericTypographic.Clone();
                format.Trimming=StringTrimming.EllipsisCharacter;
                var measured=g.MeasureString(value,font,new SizeF(bounds.Width,10000),format);
                if(required && measured.Height>bounds.Height+0.5f) throw new InvalidDataException($"Line {label.Line}: текст не помещается в прямую печать. Уменьшите название изделия или разбейте позицию.");
                g.DrawString(value,font,Brushes.Black,bounds,format);
            }
            return image;
        }
        catch {image.Dispose();throw;}
    }
}
