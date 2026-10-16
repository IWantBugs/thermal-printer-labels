using LabelCore;
namespace LabelApp;
internal sealed class ReserveSettingsForm : Form
{
    readonly DataGridView groups = new() { Dock = DockStyle.Fill, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true, AllowUserToDeleteRows = true };
    readonly NumericUpDown other = new() { Minimum = 0, Maximum = 1000, DecimalPlaces = 2, Width = 100 };
    public ReserveRules Rules { get; private set; }
    public ReserveSettingsForm(ReserveRules rules)
    {
        Rules = rules; Text = "Настройки · Умный техзапас"; Size = new Size(760, 470); MinimumSize = new Size(600, 350);
        StartPosition = FormStartPosition.CenterParent;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new System.Windows.Forms.Label { Text = "Группы обозначений: префиксы через запятую. Новая группа — в пустой строке; удалить — выделить строку и нажать Delete.", AutoSize = true, MaximumSize = new Size(660, 0), Padding = new Padding(0,0,0,12) }, 0, 0);
        groups.Columns.Add("Prefixes", "Префиксы (R, C, VD, VT)"); groups.Columns.Add("Percent", "Техзапас, %");
        foreach (var group in rules.Groups) groups.Rows.Add(group.Prefixes, group.Percent);
        layout.Controls.Add(groups, 0, 1);
        var fallback = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0,12,0,12) };
        other.Value = rules.OtherPercent;
        fallback.Controls.Add(new System.Windows.Forms.Label { Text = "Разъёмы и прочие обозначения, %", AutoSize = true }); fallback.Controls.Add(other);
        layout.Controls.Add(fallback, 0, 2);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        var save = new Button { Text = "Сохранить", AutoSize = true };
        save.Click += (_, _) =>
        {
            try
            {
                groups.EndEdit(); var result = new ReserveRules { Groups = [], OtherPercent = other.Value, SchemaVersion = 1 };
                foreach (DataGridViewRow row in groups.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (!decimal.TryParse(Convert.ToString(row.Cells[1].Value)?.Replace(',', '.'), System.Globalization.NumberStyles.AllowDecimalPoint, System.Globalization.CultureInfo.InvariantCulture, out var percent)) throw new ArgumentException("Введите числовой процент для каждой группы.");
                    result.Groups.Add(new ReserveGroup(Convert.ToString(row.Cells[0].Value) ?? "", percent));
                }
                result.Validate(); Rules = result; DialogResult = DialogResult.OK;
            }
            catch (ArgumentException ex) { MessageBox.Show(this, ex.Message, "Проверьте настройки"); }
        };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true };
        buttons.Controls.AddRange([save,cancel]); layout.Controls.Add(buttons,0,3); Controls.Add(layout); CancelButton = cancel; Appearance.Dialog(this);
    }
}
