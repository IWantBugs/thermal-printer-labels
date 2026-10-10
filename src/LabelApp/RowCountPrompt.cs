namespace LabelApp;

internal sealed class RowCountPrompt : Form
{
    readonly NumericUpDown count = new() { Minimum = 1, Maximum = 10000, Value = 1, Width = 130 };
    public int Count => (int)count.Value;
    public RowCountPrompt()
    {
        Text = "Добавить строки"; ClientSize = new Size(340, 130);
        StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;
        var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12) };
        layout.Controls.Add(new System.Windows.Forms.Label { Text = "Сколько строк добавить?", AutoSize = true });
        layout.SetFlowBreak(layout.Controls[0], true);
        layout.Controls.Add(count); layout.SetFlowBreak(count, true);
        var ok = new Button { Text = "Добавить", DialogResult = DialogResult.OK, AutoSize = true };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true };
        layout.Controls.AddRange([ok, cancel]); Controls.Add(layout);
        AcceptButton = ok; CancelButton = cancel;
        Appearance.Dialog(this);
    }
}
