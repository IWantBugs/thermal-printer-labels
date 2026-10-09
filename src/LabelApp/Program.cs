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
    readonly DataGridView grid = new() { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    readonly TextBox product = new() { Width = 220 };
    readonly CheckBox package = new() { Text = "Добавить «Упк, шт»", AutoSize = true };
    readonly Label status = new() { AutoSize = true, Text = "Выберите BOM (.xlsx). Данные можно исправить перед сохранением." };
    string? source;
    public MainForm()
    {
        Text = "BOM → этикетки 58×40 мм · 1.0.6"; Width = 1120; Height = 640; MinimumSize = new Size(850, 450);
        var top = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 80, Padding = new Padding(8), AutoSize = true };
        var load = new Button() { Text = "Открыть BOM…", AutoSize = true }; load.Click += (_, _) => LoadBom();
        var save = new Button() { Text = "Сохранить DOCX", AutoSize = true }; save.Click += (_, _) => Run(() => Generate());
        var preview = new Button() { Text = "Предпросмотр в Word", AutoSize = true }; preview.Click += (_, _) => Run(() => { var p = Generate(); if (p != null) OpenWord(p, false); });
        var print = new Button() { Text = "Печать с предпросмотром", AutoSize = true }; print.Click += (_, _) => Run(() => { var p = Generate(); if (p != null) OpenWord(p, true); });
        var single = new Button() { Text = "Сформировать этикетку из строки", AutoSize = true, Enabled = false };
        single.Click += (_, _) => Run(() =>
        {
            var row = grid.CurrentRow ?? throw new InvalidOperationException("Выберите строку BOM в таблице.");
            var path = Generate(row);
            if (path != null) OpenWord(path, true, 1);
        });
        grid.SelectionChanged += (_, _) => single.Enabled = source != null && grid.CurrentRow != null;
        top.Controls.AddRange([load, new Label() { Text = "Изделие:", AutoSize = true, Padding = new Padding(0, 7, 0, 0) }, product, package, save, preview, print, single]);
        var bottom = new FlowLayoutPanel() { Dock = DockStyle.Bottom, Height = 45, Padding = new Padding(8) }; bottom.Controls.Add(status);
        foreach (var (key, title) in new[] { ("Line", "Line #"), ("Designator", "Обозн."), ("Part", "Парт"), ("Quantity", "Кол-во / плата"), ("Summary", "Sum"), ("Package", "Упк, шт") }) grid.Columns.Add(key, title);
        grid.Columns["Package"].Visible = false;
        package.CheckedChanged += (_, _) => { grid.Columns["Package"].Visible = package.Checked; if (package.Checked) status.Text = "Введите количество в упаковке для каждой позиции в столбце «Упк, шт»."; };
        Controls.Add(grid); Controls.Add(top); Controls.Add(bottom);
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
        status.Text = $"Загружено позиций: {records.Count}. " + (package.Checked ? "Заполните «Упк, шт»." : "Проверьте данные перед генерацией.");
    });
    // Use a conservative fit check; Word remains the final pagination engine.
    bool Fits(string text, int lines, bool bold = false)
    {
        using var g = CreateGraphics(); using var font = new Font("Arial", 8, bold ? FontStyle.Bold : FontStyle.Regular); using var sf = (StringFormat)StringFormat.GenericTypographic.Clone();
        sf.FormatFlags &= ~StringFormatFlags.NoWrap;
        var width = 39f / 25.4f * g.DpiX;
        var size = g.MeasureString(text, font, new SizeF(width, 10000), sf);
        return size.Width <= width + 0.5f && size.Height <= lines * font.GetHeight(g) + 0.5f;
    }
    string Shorten(string value) { if (Fits(value, 3)) return value; var elements = System.Globalization.StringInfo.ParseCombiningCharacters(value); for (int i = elements.Length - 1; i >= 0; i--) { var s = value[..elements[i]].TrimEnd() + "…"; if (Fits(s, 3)) return s; } return "…"; }
    string? Generate(DataGridViewRow? singleRow = null)
    {
        grid.EndEdit(); Validate();
        if (source is null) throw new InvalidOperationException("Сначала выберите BOM.");
        var title = product.Text.Trim(); if (title.Length == 0) throw new InvalidDataException("Введите название изделия.");
        var labels = new List<BomLabel>(); int trimmed = 0;
        var selectedRows = singleRow == null ? grid.Rows.Cast<DataGridViewRow>().ToArray() : new[] { singleRow };
        foreach (var row in selectedRows)
        {
            string V(string key) => Convert.ToString(row.Cells[key].Value)?.Trim() ?? "";
            string line = V("Line"), des = V("Designator"), part = V("Part"), qty = V("Quantity"), sum = V("Summary"), pack = V("Package");
            if (new[] { line, des, part, qty, sum }.Any(string.IsNullOrWhiteSpace)) throw new InvalidDataException($"Позиция {row.Index + 1}: заполните Line #, Обозн., Парт, Кол-во и Sum.");
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
            foreach (var (name, value) in new[] { ("Обозн.", des), ("Парт", part), ("Кол-во", qty), ("Упк, шт", package.Checked ? pack : "") }) if (!Fits(value, 1)) throw new InvalidDataException($"Позиция {row.Index + 1}: «{name}» не помещается. Сократите значение в таблице.");
            var shortSum = Shorten(sum); if (shortSum != sum) trimmed++;
            labels.Add(new BomLabel(line, des, part, qty, shortSum, package.Checked ? pack : null));
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
            WordLayout.Normalize(candidate, labels.Count);
            File.Move(candidate, path, true);
        }
        finally { if (File.Exists(candidate)) File.Delete(candidate); }
        status.Text = singleRow == null
            ? $"Сохранено: {Path.GetFileName(path)}. Этикеток: {labels.Count}; сокращено Sum: {trimmed}."
            : $"Подготовлена этикетка для Line # {labels[0].Line}. Сокращено Sum: {trimmed}.";
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
            int expectedLabels = labelCount ?? grid.Rows.Count;
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
    public Picker(string title, string[] items) { Text = title; Width = 720; Height = 150; StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; choices.Items.AddRange(items); choices.SelectedIndex = 0; var ok = new Button() { Text = "Выбрать", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK }; var cancel = new Button() { Text = "Отмена", Dock = DockStyle.Bottom, DialogResult = DialogResult.Cancel }; Controls.Add(choices); Controls.Add(ok); Controls.Add(cancel); AcceptButton = ok; CancelButton = cancel; }
}
public sealed class TextPrompt : Form
{
    readonly TextBox input = new() { Dock = DockStyle.Top }; public string Value => input.Text.Trim();
    public TextPrompt(string message, string value, int maxLength) { Text = "Сократите название изделия"; Width = 700; Height = 170; StartPosition = FormStartPosition.CenterParent; var label = new Label() { Text = message, Dock = DockStyle.Top, Height = 40 }; input.MaxLength = maxLength; input.Text = value[..Math.Min(value.Length, maxLength)]; var ok = new Button() { Text = "Применить", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK }; var cancel = new Button() { Text = "Отмена", Dock = DockStyle.Bottom, DialogResult = DialogResult.Cancel }; Controls.Add(input); Controls.Add(label); Controls.Add(ok); Controls.Add(cancel); AcceptButton = ok; CancelButton = cancel; }
}
