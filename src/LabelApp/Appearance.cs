namespace LabelApp;

internal static class Appearance
{
    public static readonly Color Ink = Color.FromArgb(35, 52, 62);
    public static readonly Color Accent = Color.FromArgb(190, 48, 36);
    public static readonly Color Background = Color.FromArgb(235, 241, 245);
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
        button.FlatAppearance.BorderSize=1;
        button.Font=new Font("Segoe UI",9.5f,primary?FontStyle.Bold:FontStyle.Regular);
        button.MinimumSize=new Size(0,34);
        button.Padding=new Padding(9,4,9,4);
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
        grid.GridColor = Color.FromArgb(177, 191, 201); grid.CellBorderStyle = DataGridViewCellBorderStyle.Single;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Ink, ForeColor = Color.White, SelectionBackColor = Ink,
            SelectionForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold), Padding = new Padding(8, 0, 8, 0)
        };
        grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        grid.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.White, ForeColor = Ink, SelectionBackColor = Color.FromArgb(219, 235, 243),
            SelectionForeColor = Ink, Padding = new Padding(8, 3, 8, 3), Font = new Font("Segoe UI", 10),
            WrapMode = DataGridViewTriState.True, Alignment = DataGridViewContentAlignment.MiddleLeft
        };
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 249, 251);
        grid.RowTemplate.MinimumHeight = 34;
        grid.RowTemplate.Height = 34;
        grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        grid.AllowUserToResizeRows = false;
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
            // Measure natural text width independently of the current wrapped row height.
            int Measure(string text, Font font) => TextRenderer.MeasureText(text.Replace('\r',' ').Replace('\n',' '),font,
                new Size(int.MaxValue,int.MaxValue),TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
            int preferred=Measure(column.HeaderText,grid.ColumnHeadersDefaultCellStyle.Font ?? grid.Font)+24;
            foreach(DataGridViewRow row in grid.Rows)
            {
                preferred=Math.Max(preferred,Measure(Convert.ToString(row.Cells[column.Index].Value)??"",grid.DefaultCellStyle.Font ?? grid.Font)+24);
                if(preferred>=limit*scale) break;
            }
            column.FillWeight = Math.Clamp(preferred, column.MinimumWidth, Math.Max(column.MinimumWidth, (int)(limit * scale))) / scale;
        }
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
    }

}
