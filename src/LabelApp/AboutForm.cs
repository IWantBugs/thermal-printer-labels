using System.Diagnostics;
using System.Reflection;
using Label = System.Windows.Forms.Label;
namespace LabelApp;

internal sealed class AboutForm : Form
{
    public AboutForm(Icon icon)
    {
        Text = "О программе · SPK"; Width = 460; Height = 260; Icon = icon;
        StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;
        var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20), AutoScroll = true };
        layout.Controls.Add(new Label { AutoSize = true, Text = "SPK Thermal Printer Labels\nВерсия 1.0.13\n\nИзготовитель: SPK\nЭтикетки 58×40 мм из BOM Excel." });
        var site = new LinkLabel { AutoSize = true, Text = "spkspb.ru", Margin = new Padding(3, 14, 3, 14) };
        site.LinkClicked += (_, _) =>
        {
            try { Process.Start(new ProcessStartInfo("https://spkspb.ru") { UseShellExecute = true }); }
            catch (Exception e) { MessageBox.Show(this, e.Message, "Не удалось открыть сайт"); }
        };
        layout.Controls.Add(site);
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SPK.SiteQr");
        if (stream != null)
        {
            using var original = Image.FromStream(stream);
            var picture = new PictureBox { Image = new Bitmap(original), Width = 380, Height = 380, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White };
            FormClosed += (_, _) => picture.Image?.Dispose();
            layout.Controls.Add(picture); Height = 650;
        }
        Controls.Add(layout);
        Appearance.Dialog(this);
    }
}
