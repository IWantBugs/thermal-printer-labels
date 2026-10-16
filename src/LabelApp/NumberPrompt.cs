namespace LabelApp;
internal sealed class NumberPrompt : Form
{
    readonly NumericUpDown number;
    public decimal Value => number.Value;
    public NumberPrompt(string title, string question, decimal value, decimal minimum, decimal maximum, int decimals = 0)
    {
        Text = title; ClientSize = new Size(430, 160); StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
        var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        layout.Controls.Add(new System.Windows.Forms.Label { Text = question, AutoSize = true });
        number = new NumericUpDown { Minimum = minimum, Maximum = maximum, DecimalPlaces = decimals, Value = value, Width = 200, ThousandsSeparator = true, Margin = new Padding(3, 12, 3, 12) };
        layout.Controls.Add(number);
        var buttons = new FlowLayoutPanel { AutoSize = true };
        var ok = new Button { Text = "Применить", DialogResult = DialogResult.OK, AutoSize = true };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true };
        buttons.Controls.AddRange([ok, cancel]); layout.Controls.Add(buttons); Controls.Add(layout);
        AcceptButton = ok; CancelButton = cancel; Appearance.Dialog(this);
    }
}
