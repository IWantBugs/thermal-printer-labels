using LabelCore;
namespace LabelApp;
internal sealed class PackagePiecesForm : Form
{
    readonly TextBox input = new() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };
    public int[] Pieces { get; private set; } = [];
    public PackagePiecesForm(int[] pieces)
    {
        Text = "Разбить Упк, шт на куски"; Size = new Size(500, 420); StartPosition = FormStartPosition.CenterParent;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 3 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent,100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new System.Windows.Forms.Label { Text = "Количество в каждом куске — с новой строки.\nНапример: 150 и 150. Сумма = 300.\nУдалите строку для удаления куска; очистите всё для отмены разбиения.", AutoSize = true, Padding = new Padding(0,0,0,12) },0,0);
        input.Text = string.Join(Environment.NewLine,pieces); layout.Controls.Add(input,0,1);
        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top };
        var save = new Button { Text = "Применить", AutoSize = true };
        save.Click += (_, _) => { try { Pieces = PackagePieces.Parse(input.Text); DialogResult = DialogResult.OK; } catch(ArgumentException ex) { MessageBox.Show(this,ex.Message,"Упаковки"); } };
        var cancel = new Button { Text = "Отмена", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange([save,cancel]); layout.Controls.Add(buttons,0,2); Controls.Add(layout); CancelButton=cancel; Appearance.Dialog(this);
    }
}
