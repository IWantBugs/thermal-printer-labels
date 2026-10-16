namespace LabelApp;
internal sealed class PackagePiecesForm : Form
{
    readonly List<NumericUpDown> quantities = [];
    public int[] Pieces { get; private set; } = [];
    public PackagePiecesForm(int count, int[] previous)
    {
        Text = "Количество в каждой упаковке"; Size = new Size(460, 470); MinimumSize = new Size(380, 300); StartPosition = FormStartPosition.CenterParent;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 3 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent,100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new System.Windows.Forms.Label { Text = "Введите количество в каждой упаковке.\nПозже его можно изменить прямо в таблице.", AutoSize = true, Padding = new Padding(0,0,0,12) },0,0);
        var fields = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
        for(int i=0;i<count;i++)
        {
            var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            row.Controls.Add(new System.Windows.Forms.Label { Text = $"Упк {i+1}, шт", Width=130, Padding = new Padding(0,5,0,0) });
            var input = new NumericUpDown { Minimum=1, Maximum=int.MaxValue, Value=i<previous.Length ? previous[i] : 1, Width=150, ThousandsSeparator=true };
            quantities.Add(input); row.Controls.Add(input); fields.Controls.Add(row);
        }
        layout.Controls.Add(fields,0,1);
        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top };
        var save = new Button { Text = "Сохранить", AutoSize = true };
        save.Click += (_, _) =>
        {
            if(quantities.Sum(n=>n.Value)>int.MaxValue) { MessageBox.Show(this,"Сумма упаковок слишком велика."); return; }
            Pieces=quantities.Select(n=>(int)n.Value).ToArray(); DialogResult=DialogResult.OK;
        };
        var cancel = new Button { Text = "Отмена", AutoSize = true, DialogResult = DialogResult.Cancel };
        buttons.Controls.AddRange([save,cancel]); layout.Controls.Add(buttons,0,2); Controls.Add(layout); CancelButton=cancel; Appearance.Dialog(this);
    }
}
