using LabelCore;
using Label = System.Windows.Forms.Label;
namespace LabelApp;

internal sealed class ColumnMapping : Form
{
    readonly Sheet sheet;
    bool refreshing;
    readonly ComboBox header = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
    readonly ComboBox[] fields = Enumerable.Range(0, 5).Select(_ => new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, DropDownWidth = 1000 }).ToArray();
    public int HeaderRow => header.SelectedIndex;
    public int[] Map => fields.Select((c, i) => c.SelectedIndex - (i == 4 ? 1 : 0)).ToArray();

    public ColumnMapping(Sheet sheet, int headerRow)
    {
        this.sheet = sheet;
        Text = "Сопоставление столбцов BOM"; Width = 900; Height = 410;
        StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 2, RowCount = 7 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.Controls.Add(new Label { Text = "Строка заголовков", AutoSize = true }, 0, 0);
        layout.Controls.Add(header, 1, 0);
        var labels = new[] { "Номер (Line #)", "Обозн. (Designator)", "Парт (PartNumber)", "Кол-во (Quantity)", "Sum (необязательно)" };
        for (int i = 0; i < fields.Length; i++)
        {
            layout.Controls.Add(new Label { Text = labels[i], AutoSize = true }, 0, i + 1);
            layout.Controls.Add(fields[i], 1, i + 1);
        }
        var hint = new Label { Text = "Все столбцы доступны для каждого поля. Один столбец можно выбрать несколько раз.", AutoSize = true };
        layout.Controls.Add(hint, 0, 6); layout.SetColumnSpan(hint, 2);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true };
        var accept = new Button { Text = "Загрузить", AutoSize = true };
        accept.Click += (_, _) =>
        {
            if (fields.Any(f => f.SelectedIndex < 0))
            {
                MessageBox.Show(this, "Выберите столбец для каждого поля этикетки.", "Сопоставление", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            DialogResult = DialogResult.OK;
        };
        buttons.Controls.Add(cancel); buttons.Controls.Add(accept);
        Controls.Add(layout); Controls.Add(buttons); AcceptButton = accept; CancelButton = cancel;
        header.Items.AddRange(sheet.Rows.Select((r, i) => $"Строка данных {i + 1}: " + string.Join(" | ", r)).ToArray());
        fields[2].SelectedIndexChanged += (_, _) =>
        {
            // Name may hold the part number, rather than a separate summary.
            if (!refreshing && fields[2].SelectedIndex >= 0 && fields[4].SelectedIndex > 0
                && fields[4].SelectedIndex - 1 == fields[2].SelectedIndex)
                fields[4].SelectedIndex = 0;
        };
        header.SelectedIndexChanged += (_, _) => RefreshColumns();
        header.SelectedIndex = headerRow;
        Appearance.Dialog(this);
    }

    void RefreshColumns()
    {
        refreshing = true;
        var options = Bom.Columns(sheet, HeaderRow);
        var defaults = ColumnDefaults.Resolve(sheet.Rows[HeaderRow]);
        for (int i = 0; i < fields.Length; i++)
        {
            fields[i].Items.Clear();
            if (i == 4) fields[i].Items.Add("Не использовать — заполню вручную при необходимости");
            fields[i].Items.AddRange(options);
            fields[i].SelectedIndex = defaults[i] + (i == 4 ? 1 : 0);
        }
        refreshing = false;
    }
}
