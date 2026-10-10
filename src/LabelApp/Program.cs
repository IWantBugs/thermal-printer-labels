using LabelCore;
using Label = System.Windows.Forms.Label;
using BomLabel = LabelCore.Label;
using System.Runtime.InteropServices;
namespace LabelApp;
internal static class Program
{
    [STAThread] static void Main() { ApplicationConfiguration.Initialize(); Application.Run(new MainForm()); }
}
public sealed class MainForm : Form
{
    readonly DataGridView grid = new() { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, MultiSelect = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    readonly TextBox product = new() { Width = 220 };
    readonly CheckBox package = new() { Text = "Добавить «Упк, шт»", AutoSize = true };
    readonly Label status = new() { AutoSize = true, Text = "Выберите BOM (.xlsx). Данные можно исправить перед сохранением." };
    string? source;
    int generatedLabels;
    public MainForm()
    {
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(Environment.ProcessPath!) ?? SystemIcons.Application;
        Text = "SPK · BOM → этикетки 58×40 мм · 1.0.13"; Width = 1280; Height = 780; MinimumSize = new Size(940, 620);
        Font = new Font("Segoe UI", 10); BackColor = Appearance.Background; ForeColor = Appearance.Ink;
        StartPosition = FormStartPosition.CenterScreen; AutoScaleMode = AutoScaleMode.Dpi;
        var load = new Button() { Text = "Открыть BOM…", AutoSize = true }; load.Click += (_, _) => LoadBom();
        var save = new Button() { Text = "Сохранить DOCX", AutoSize = true }; save.Click += (_, _) => Run(() => Generate());
        var preview = new Button() { Text = "Предпросмотр в Word", AutoSize = true }; preview.Click += (_, _) => Run(() => { var p = Generate(); if (p != null) OpenWord(p, false); });
        var print = new Button() { Text = "Печать с предпросмотром", AutoSize = true }; print.Click += (_, _) => Run(() => { var p = Generate(); if (p != null) OpenWord(p, true); });
        var single = new Button() { Text = "Сформировать этикетку из строки", AutoSize = true, Enabled = false };
        single.Click += (_, _) => Run(() =>
        {
            var row = grid.CurrentRow ?? throw new InvalidOperationException("Выберите строку BOM в таблице.");
            var path = Generate(row);
            if (path != null) OpenWord(path, true, generatedLabels);
        });
        grid.SelectionChanged += (_, _) => single.Enabled = source != null && grid.CurrentRow != null;
        var addRows = new Button { Text = "Добавить строки…", AutoSize = true, Enabled = false };
        addRows.Click += (_, _) => Run(AddRows);
        var deleteRows = new Button { Text = "Удалить выбранные строки", AutoSize = true, Enabled = false };
        deleteRows.Click += (_, _) => Run(() =>
        {
            grid.EndEdit();
            var rows = grid.SelectedRows.Cast<DataGridViewRow>().ToArray();
            foreach (var row in rows) grid.Rows.Remove(row);
            status.Text = $"Удалено строк: {rows.Length}. Всего позиций: {grid.Rows.Count}.";
        });
        var sortLine = new Button { Text = "Сортировать по Line ↑", AutoSize = true, Enabled = false };
        sortLine.Click += (_, _) => Run(() => { grid.EndEdit(); grid.Sort(grid.Columns["Line"], System.ComponentModel.ListSortDirection.Ascending); });
        void UpdateRowActions()
        {
            addRows.Enabled = source != null;
            deleteRows.Enabled = source != null && grid.SelectedRows.Count > 0;
            sortLine.Enabled = source != null && grid.Rows.Count > 1;
        }
        grid.SelectionChanged += (_, _) => UpdateRowActions();
        grid.RowsAdded += (_, _) => UpdateRowActions();
        grid.RowsRemoved += (_, _) => UpdateRowActions();
        grid.SortCompare += (_, e) =>
        {
            if (e.Column.Name != "Line") return;
            e.SortResult = LineOrder.Compare(Convert.ToString(e.CellValue1), Convert.ToString(e.CellValue2));
            if (e.SortResult == 0) e.SortResult = e.RowIndex1.CompareTo(e.RowIndex2);
            e.Handled = true;
        };
        var about = new Button() { Text = "О программе", AutoSize = true };
        about.Click += (_, _) => { using var info = new AboutForm(Icon); info.ShowDialog(this); };
        foreach (var (key, title) in new[] { ("Line", "Line #"), ("Designator", "Обозн."), ("Part", "Парт"), ("Quantity", "Кол-во / плата"), ("Summary", "Sum"), ("Package", "Упк, шт") }) grid.Columns.Add(key, title);
        grid.Columns["Package"].Visible = false;
        package.CheckedChanged += (_, _) => { grid.Columns["Package"].Visible = package.Checked; Appearance.FitColumns(grid); if (package.Checked) status.Text = "Введите количество в упаковке для каждой позиции в столбце «Упк, шт»."; };
        BuildWorkspace(load, save, preview, print, single, addRows, deleteRows, sortLine, about);
    }
    void BuildWorkspace(Button load, Button save, Button preview, Button print, Button single,
        Button addRows, Button deleteRows, Button sortLine, Button about)
    {
        SuspendLayout();
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(18) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (int i = 0; i < 6; i++) root.RowStyles.Add(new RowStyle(i == 4 ? SizeType.Percent : SizeType.AutoSize, i == 4 ? 100 : 0));
        var header = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3, BackColor = Appearance.Ink, Padding = new Padding(18, 14, 18, 14), Margin = new Padding(0, 0, 0, 14) };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var title = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        title.Controls.Add(new Label { Text = "SPK  /  Этикетки из BOM", AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 19, FontStyle.Bold) });
        title.Controls.Add(new Label { Text = "58 × 40 мм   ·   Word   ·   Xprinter XP-365B", AutoSize = true, ForeColor = Color.FromArgb(191, 207, 215), Margin = new Padding(3, 6, 3, 0) });
        Appearance.Button(about); about.Padding = new Padding(10, 6, 10, 6); about.Anchor = AnchorStyles.Right;
        using var logoStream = typeof(MainForm).Assembly.GetManifestResourceStream("SPK.Logo")!;
        using var logoIcon = new System.Drawing.Icon(logoStream, new Size(256, 256));
        var logo = new PictureBox { Image = logoIcon.ToBitmap(), SizeMode = PictureBoxSizeMode.Zoom,
            Width = 76, Height = 76, BackColor = Color.White, Margin = new Padding(0, 0, 14, 0), AccessibleName = "Логотип SPK" };
        FormClosed += (_, _) => logo.Image?.Dispose();
        header.Controls.Add(logo, 0, 0); header.Controls.Add(title, 1, 0); header.Controls.Add(about, 2, 0); root.Controls.Add(header, 0, 0);
        FlowLayoutPanel Toolbar(params Control[] controls)
        {
            var panel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, BackColor = Color.White, Padding = new Padding(10), Margin = new Padding(0, 0, 0, 8) };
            foreach (var control in controls)
            {
                control.Margin = new Padding(4, 4, 8, 4);
                if (control is Button button) { Appearance.Button(button, button == load || button == print, button == deleteRows); button.Padding = new Padding(10, 6, 10, 6); }
                panel.Controls.Add(control);
            }
            return panel;
        }
        product.Width = 270;
        package.Padding = new Padding(4, 6, 4, 0);
        root.Controls.Add(Toolbar(load, new Label { Text = "Изделие", AutoSize = true, Padding = new Padding(2, 7, 2, 0) }, product, package), 0, 1);
        root.Controls.Add(Toolbar(save, preview, print, single), 0, 2);
        root.Controls.Add(Toolbar(addRows, deleteRows, sortLine), 0, 3);
        Appearance.Table(grid);
        var table = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = Appearance.Border, Margin = new Padding(0) };
        table.Controls.Add(grid); root.Controls.Add(table, 0, 4);
        status.AutoSize = true; status.Dock = DockStyle.Fill;
        status.Padding = new Padding(4, 12, 4, 4); status.Margin = new Padding(0); status.ForeColor = Color.FromArgb(79, 99, 111);
        root.Controls.Add(status, 0, 5);
        Controls.Add(root); ResumeLayout(true);
    }
    void Run(Action action) { try { action(); } catch (Exception e) { MessageBox.Show(this, e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); } }
    void LoadBom() => Run(() =>
    {
        using var dialog = new OpenFileDialog() { Filter = "Excel BOM (*.xlsx)|*.xlsx", Title = "Выберите BOM" }; if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var sheets = Bom.Read(dialog.FileName); if (sheets.Count == 0) throw new InvalidDataException("В книге нет непустых листов.");
        int si = 0; if (sheets.Count > 1) { using var pick = new Picker("Выберите лист BOM", sheets.Select(s => s.Name).ToArray()); if (pick.ShowDialog(this) != DialogResult.OK) return; si = pick.Index; }
        var sheet = sheets[si];
        int hi = sheet.Rows.FindIndex(r => r.Any(c => c.Equals("Designator", StringComparison.OrdinalIgnoreCase)) && r.Any(c => c.Equals("Quantity", StringComparison.OrdinalIgnoreCase)));
        using var mapping = new ColumnMapping(sheet, Math.Max(hi, 0));
        if (mapping.ShowDialog(this) != DialogResult.OK) return;
        var records = Bom.ProjectRows(sheet, mapping.HeaderRow, mapping.Map);
        if (records.Count == 0) throw new InvalidDataException("После заголовков нет позиций BOM.");
        source = dialog.FileName; product.Text = Bom.Product(source); grid.Rows.Clear();
        foreach (var r in records) grid.Rows.Add(r[0], r[1], r[2], r[3], r[4], "");
        package.Checked = MessageBox.Show(this, "Добавить поле «Упк, шт»? При выборе «Да» введите количество для каждой позиции в таблице.", "Упаковка", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        Appearance.FitColumns(grid);
        status.Text = $"Загружено позиций: {records.Count}. " + (package.Checked ? "Заполните «Упк, шт»." : "Проверьте данные перед генерацией.");
    });
    void AddRows()
    {
        if (source == null) throw new InvalidOperationException("Сначала выберите BOM.");
        grid.EndEdit();
        using var dialog = new RowCountPrompt();
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var next = System.Numerics.BigInteger.Zero;
        foreach (DataGridViewRow row in grid.Rows)
            if (System.Numerics.BigInteger.TryParse(Convert.ToString(row.Cells["Line"].Value), out var number) && number > next) next = number;
        int first = grid.Rows.Count;
        for (int i = 0; i < dialog.Count; i++) grid.Rows.Add((++next).ToString(), "", "", "1", "", "");
        grid.ClearSelection();
        grid.CurrentCell = grid.Rows[first].Cells["Designator"];
        grid.Rows[first].Selected = true;
        status.Text = $"Добавлено строк: {dialog.Count}. Заполните Обозн., Парт" + (package.Checked ? " и Упк, шт." : ".") + " Line и количество можно изменить.";
    }
    // Use a conservative fit check; Word remains the final pagination engine.
    bool Fits(string text, int lines, bool bold = false, float fontSize = 8)
    {
        using var g = CreateGraphics(); using var font = new Font("Arial", fontSize, bold ? FontStyle.Bold : FontStyle.Regular); using var sf = (StringFormat)StringFormat.GenericTypographic.Clone();
        sf.FormatFlags &= ~StringFormatFlags.NoWrap;
        var width = 39f / 25.4f * g.DpiX;
        var size = g.MeasureString(text, font, new SizeF(width, 10000), sf);
        return size.Width <= width + 0.5f && size.Height <= lines * font.GetHeight(g) + 0.5f;
    }
    string Shorten(string value, int lines = 3) { if (Fits(value, lines)) return value; var elements = System.Globalization.StringInfo.ParseCombiningCharacters(value); for (int i = elements.Length - 1; i >= 0; i--) { var s = value[..elements[i]].TrimEnd() + "…"; if (Fits(s, lines)) return s; } return "…"; }
    string? Generate(DataGridViewRow? singleRow = null)
    {
        grid.EndEdit(); Validate();
        if (source is null) throw new InvalidOperationException("Сначала выберите BOM.");
        if (grid.Rows.Count == 0) throw new InvalidDataException("В таблице нет позиций. Добавьте строки или загрузите BOM.");
        var title = product.Text.Trim(); if (title.Length == 0) throw new InvalidDataException("Введите название изделия.");
        var labels = new List<BomLabel>(); int trimmed = 0;
        var selectedRows = singleRow == null ? grid.Rows.Cast<DataGridViewRow>().ToArray() : new[] { singleRow };
        foreach (var row in selectedRows)
        {
            string V(string key) => Convert.ToString(row.Cells[key].Value)?.Trim() ?? "";
            string line = V("Line"), des = V("Designator"), part = V("Part"), qty = V("Quantity"), sum = V("Summary"), pack = V("Package");
            if (new[] { line, des, part, qty }.Any(string.IsNullOrWhiteSpace)) throw new InvalidDataException($"Позиция {row.Index + 1}: заполните Line #, Обозн., Парт и Кол-во.");
            if (!decimal.TryParse(qty.Replace(',', '.'), System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var q) || q <= 0) throw new InvalidDataException($"Позиция {row.Index + 1}: «Кол-во» должно быть положительным числом.");
            if (package.Checked && (!int.TryParse(pack, out var n) || n <= 0)) throw new InvalidDataException($"Позиция {row.Index + 1}: введите целое положительное «Упк, шт».");
            if (!Fits(Document.Heading(title, line), 2, true))
            {
                int limit = title.Length; while (limit > 0 && !Fits(Document.Heading(title[..limit], line), 2, true)) limit--;
                if (limit == 0) throw new InvalidDataException($"Позиция {row.Index + 1}: номер Line # слишком длинный для этикетки.");
                using var prompt = new TextPrompt($"Название изделия не помещается. Сократите его (для текущего текста до {limit} символов).", title, limit);
                if (prompt.ShowDialog(this) != DialogResult.OK) return null;
                product.Text = prompt.Value; return Generate(singleRow);
            }
            foreach (var (name, value) in new[] { ("Кол-во", qty), ("Упк, шт", package.Checked ? pack : "") }) if (!Fits(value, 1)) throw new InvalidDataException($"Позиция {row.Index + 1}: «{name}» не помещается. Сократите значение в таблице.");
            var original = new BomLabel(line, des, part, qty, sum, package.Checked ? pack : null);
            bool FitDesignation(string text, float font, int lines) => Fits(text, lines, false, font);
            var fitted = Designators.Fit(original, FitDesignation);
            List<BomLabel> parts;
            if (fitted != null) parts = [fitted];
            else
            {
                if (MessageBox.Show(this, $"Line # {line}: полные «Обозн.» и «Парт» вместе не помещаются на одной этикетке даже шрифтом 7 пт. Разбить позицию на несколько этикеток? Оба поля будут сохранены полностью, Кол-во останется на одну плату.", "Длинное обозначение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return null;
                parts = Designators.Split(original, FitDesignation);
            }
            foreach (var partLabel in parts)
            {
                var shortSum = partLabel.SummaryLines == 0 ? "" : Shorten(sum, partLabel.SummaryLines); if (shortSum != sum) trimmed++;
                labels.Add(partLabel with { Summary = shortSum });
            }
        }
        string path;
        if (singleRow == null) path = Bom.Output(source);
        else
        {
            var temporaryDirectory = Path.Combine(Path.GetTempPath(), "ThermalPrinterLabels");
            Directory.CreateDirectory(temporaryDirectory);
            path = Path.Combine(temporaryDirectory, Guid.NewGuid().ToString("N") + ".docx");
        }
        if (File.Exists(path) && MessageBox.Show(this, $"Файл уже существует:\n{path}\n\nЗаменить?", "Сохранение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return null;
        var candidate = path + "." + Guid.NewGuid().ToString("N") + ".docx";
        try
        {
            Document.Save(candidate, title, labels);
            WordLayout.Normalize(candidate, labels);
            File.Move(candidate, path, true);
        }
        finally { if (File.Exists(candidate)) File.Delete(candidate); }
        generatedLabels = labels.Count;
        status.Text = singleRow == null
            ? $"Сохранено: {Path.GetFileName(path)}. Этикеток: {labels.Count}; сокращено Sum: {trimmed}."
            : $"Подготовлено этикеток: {labels.Count} для Line # {labels[0].Line}. Сокращено Sum: {trimmed}.";
        return path;
    }
    void OpenWord(string path, bool printing, int? labelCount = null)
    {
        dynamic? word = null, doc = null;
        try
        {
            var type = Type.GetTypeFromProgID("Word.Application") ?? throw new InvalidOperationException("Microsoft Word не установлен.");
            word = Activator.CreateInstance(type)!; word.Visible = true;
            doc = word.Documents.Open(path, ReadOnly: true); doc.Repaginate();
            int pages = doc.ComputeStatistics(2);
            int expectedLabels = labelCount ?? generatedLabels;
            if (pages != expectedLabels) MessageBox.Show(this, $"Word насчитал страниц: {pages}; этикеток: {expectedLabels}. Проверьте переносы и шрифты перед печатью.", "Проверка макета", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            doc.PrintPreview();
            if (printing && pages == expectedLabels && MessageBox.Show(this, "Проверьте макет в открытом Word. Открыть диалог печати?\n\nВыберите Xprinter XP-365B, бумагу 58×40 мм и масштаб 100% (без подгонки). Документ будет напечатан только после подтверждения в Word.", "Печать", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) { doc.ClosePrintPreview(); word.Dialogs[88].Show(); }
        }
        finally { if (doc != null) Marshal.ReleaseComObject(doc); if (word != null) Marshal.ReleaseComObject(word); }
    }
}
public sealed class Picker : Form
{
    readonly ComboBox choices = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    public int Index => choices.SelectedIndex;
    public Picker(string title, string[] items) { Text = title; Width = 720; Height = 150; StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; choices.Items.AddRange(items); choices.SelectedIndex = 0; var ok = new Button() { Text = "Выбрать", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK }; var cancel = new Button() { Text = "Отмена", Dock = DockStyle.Bottom, DialogResult = DialogResult.Cancel }; Controls.Add(choices); Controls.Add(ok); Controls.Add(cancel); AcceptButton = ok; CancelButton = cancel; Appearance.Dialog(this); }
}
public sealed class TextPrompt : Form
{
    readonly TextBox input = new() { Dock = DockStyle.Top }; public string Value => input.Text.Trim();
    public TextPrompt(string message, string value, int maxLength) { Text = "Сократите название изделия"; Width = 700; Height = 170; StartPosition = FormStartPosition.CenterParent; var label = new Label() { Text = message, Dock = DockStyle.Top, Height = 40 }; input.MaxLength = maxLength; input.Text = value[..Math.Min(value.Length, maxLength)]; var ok = new Button() { Text = "Применить", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK }; var cancel = new Button() { Text = "Отмена", Dock = DockStyle.Bottom, DialogResult = DialogResult.Cancel }; Controls.Add(input); Controls.Add(label); Controls.Add(ok); Controls.Add(cancel); AcceptButton = ok; CancelButton = cancel; Appearance.Dialog(this); }
}
