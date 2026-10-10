namespace LabelApp;

internal static class Appearance
{
    public static readonly Color Ink = Color.FromArgb(35, 52, 62);
    public static readonly Color Accent = Color.FromArgb(190, 48, 36);
    public static readonly Color Background = Color.FromArgb(242, 245, 247);
    public static readonly Color Border = Color.FromArgb(214, 223, 228);

    public static void Button(Button button, bool primary = false, bool danger = false)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.UseVisualStyleBackColor = false;
        button.BackColor = primary ? Accent : Color.White;
        button.ForeColor = primary ? Color.White : danger ? Accent : Ink;
        button.FlatAppearance.BorderColor = primary ? Accent : Border;
        button.FlatAppearance.MouseOverBackColor = primary ? Color.FromArgb(165, 38, 28) : Color.FromArgb(229, 237, 241);
        button.FlatAppearance.MouseDownBackColor = primary ? Color.FromArgb(143, 30, 23) : Color.FromArgb(212, 225, 233);
        button.Cursor = Cursors.Hand;
    }

    public static void Dialog(Control control)
    {
        control.ForeColor = Ink;
        if (control is Form) control.BackColor = Background;
        if (control is Button button) Button(button, button.DialogResult == DialogResult.OK || button.Text == "Загрузить");
        if (control is LinkLabel link) { link.LinkColor = Accent; link.ActiveLinkColor = Ink; }
        foreach (Control child in control.Controls) Dialog(child);
    }

    public static void Table(DataGridView grid)
    {
        grid.BackgroundColor = Color.White; grid.BorderStyle = BorderStyle.None;
        grid.GridColor = Border; grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Ink, ForeColor = Color.White, SelectionBackColor = Ink,
            SelectionForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold), Padding = new Padding(8, 0, 8, 0)
        };
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
        grid.ColumnHeadersHeight = 42;
        grid.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.White, ForeColor = Ink, SelectionBackColor = Color.FromArgb(219, 235, 243),
            SelectionForeColor = Ink, Padding = new Padding(8, 3, 8, 3), Font = new Font("Segoe UI", 10)
        };
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 249, 251);
        grid.RowTemplate.Height = 34;
        foreach (DataGridViewColumn column in grid.Columns) column.MinimumWidth = 70;
        grid.Columns["Line"].FillWeight = 45;
        grid.Columns["Designator"].FillWeight = 110;
        grid.Columns["Part"].FillWeight = 135;
        grid.Columns["Quantity"].FillWeight = 75;
        grid.Columns["Summary"].FillWeight = 160;
        grid.Columns["Package"].FillWeight = 65;
    }
    public static void FitColumns(DataGridView grid)
    {
        // Recompute after each import, including any widths changed by dragging headers.
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        float scale = grid.DeviceDpi / 96f;
        foreach (DataGridViewColumn column in grid.Columns)
        {
            if (!column.Visible) continue;
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
            int limit = column.Name == "Line" ? 130 : column.Name is "Quantity" or "Package" ? 180 : 420;
            int preferred = column.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
            column.FillWeight = Math.Clamp(preferred, column.MinimumWidth, Math.Max(column.MinimumWidth, (int)(limit * scale))) / scale;
        }
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
    }

}
