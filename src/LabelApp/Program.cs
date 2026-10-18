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
    readonly CheckBox reserve = new() { Text = "Прибавить тех запас", AutoSize = true, Enabled = false };
    readonly Button batch = new() { Text = "Ввести общее кол-во", AutoSize = true, Enabled = false };
    int boards;
    ReserveRules reserveRules = new();
    static string SettingsPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ThermalPrinterLabels", "reserve-groups.json");
    readonly Button splitPackages = new() { Text = "Разбить упаковку", AutoSize = true, Enabled = false };
    readonly Button deletePiece = new() { Text = "Удалить отрезок", AutoSize = true, Enabled = false };
    bool updatingPieces;
    bool editingSavedBom;
    string? savedState;
    string? source;
    int generatedLabels;
    public MainForm()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var saved = System.Text.Json.JsonSerializer.Deserialize<ReserveRules>(File.ReadAllText(SettingsPath));
                saved?.Validate();
                if (saved != null)
                {
                    reserveRules = saved;
                    if (saved.Upgrade()) File.WriteAllText(SettingsPath, System.Text.Json.JsonSerializer.Serialize(saved));
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ArgumentException) { }
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(Environment.ProcessPath!) ?? SystemIcons.Application;
        Text = "SPK · BOM → этикетки 58×40 мм · 1.2.7"; Width = 1280; Height = 780; MinimumSize = new Size(940, 620);
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
            var rows = grid.SelectedRows.Cast<DataGridViewRow>().ToArray();
            grid.CancelEdit();
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
        grid.Columns.Add("Total", "Общее кол-во");
        grid.Columns.Add("Notes", "Примечание");
        grid.Columns["Total"].DisplayIndex = grid.Columns["Quantity"].DisplayIndex + 1;
        grid.Columns["Total"].ReadOnly = true; grid.Columns["Total"].Visible = false;
        grid.Columns["Package"].Visible = false;
        grid.CellValueChanged += (_, e) => { if (e.RowIndex >= 0 && (e.ColumnIndex == grid.Columns["Quantity"].Index || e.ColumnIndex == grid.Columns["Designator"].Index)) RecalculateTotals(); };
        grid.RowsAdded += (_, _) => RecalculateTotals();
        reserve.CheckedChanged += (_, _) => RecalculateTotals();
        batch.Click += (_, _) => Run(() =>
        {
            grid.EndEdit();
            using var prompt = new NumberPrompt("Количество плат", "Сколько плат нужно изготовить?", Math.Max(boards, 1), 1, 1000000000);
            if (prompt.ShowDialog(this) != DialogResult.OK) return;
            boards = (int)prompt.Value; reserve.Enabled = true; grid.Columns["Total"].Visible = true;
            batch.Text = $"Плат: {boards} · изменить"; RecalculateTotals(); Appearance.FitColumns(grid);
        });
        package.CheckedChanged += (_, _) => { grid.Columns["Package"].Visible = package.Checked; Appearance.FitColumns(grid); if (package.Checked) status.Text = "Введите количество в упаковке для каждой позиции в столбце «Упк, шт»."; };
        splitPackages.Click += (_, _) => Run(() =>
        {
            if(!grid.EndEdit()) throw new InvalidDataException("Исправьте количество в упаковке.");
            if(grid.SelectedRows.Count!=1) return;
            var row=grid.SelectedRows[0]; var old=row.Tag as int[] ?? [];
            if(old.Length>0)
            {
                using var amount = new NumberPrompt("Добавить отрезок", "Количество компонентов в новом отрезке, шт", 1, 1, int.MaxValue);
                if(amount.ShowDialog(this)!=DialogResult.OK) return;
                if(old.Sum(n=>(long)n)+amount.Value>int.MaxValue) throw new InvalidDataException("Сумма упаковок слишком велика.");
                SetPieces(row,old.Append((int)amount.Value).ToArray());
            }
            else
            {
                using var count = new NumberPrompt("Число упаковок", "На сколько упаковок разбить?", 2, 1, 1000);
                if(count.ShowDialog(this)!=DialogResult.OK) return;
                using var prompt = new PackagePiecesForm((int)count.Value,[]);
                if(prompt.ShowDialog(this)!=DialogResult.OK) return;
                SetPieces(row,prompt.Pieces);
            }
            package.Checked=true; Appearance.FitColumns(grid); UpdateSplitAction();
        });
        deletePiece.Click += (_, _) => Run(() =>
        {
            if(!grid.EndEdit()) throw new InvalidDataException("Исправьте количество в упаковке.");
            if(grid.SelectedRows.Count!=1 || grid.SelectedRows[0].Tag is not int[] { Length: >0 } pieces) return;
            var row=grid.SelectedRows[0];
            using var pick=new Picker("Какой отрезок удалить?",pieces.Select((n,i)=>$"Упк {i+1}: {n} шт").ToArray());
            if(pick.ShowDialog(this)!=DialogResult.OK) return;
            var remaining=pieces.Where((n,i)=>i!=pick.Index).ToArray();
            SetPieces(row,remaining);
            if(remaining.Length==0) row.Cells["Package"].Value="";
            Appearance.FitColumns(grid);
            status.Text=remaining.Length>0 ? $"Отрезок удалён. Осталось: {remaining.Length}; сумма: {remaining.Sum()} шт." : "Последний отрезок удалён. Упк, шт можно заполнить вручную.";
        });
        grid.SelectionChanged += (_, _) => UpdateSplitAction();
        grid.RowsRemoved += (_, _) => UpdateSplitAction();
        grid.CellBeginEdit += (_, e) =>
        {
            if(grid.Columns[e.ColumnIndex].Name.StartsWith("Piece") &&
                (grid.Rows[e.RowIndex].Tag is not int[] pieces || !int.TryParse(grid.Columns[e.ColumnIndex].Name[5..],out int number) || number>pieces.Length)) e.Cancel=true;
        };
        grid.CellValidating += (_, e) =>
        {
            if(!grid.Columns[e.ColumnIndex].Name.StartsWith("Piece") || grid.Rows[e.RowIndex].Tag is not int[] pieces) return;
            int number=int.Parse(grid.Columns[e.ColumnIndex].Name[5..]);
            if(number>pieces.Length) return;
            if(!int.TryParse(Convert.ToString(e.FormattedValue),out int value) || value<=0 || pieces.Sum(n=>(long)n)-pieces[number-1]+value>int.MaxValue)
            { e.Cancel=true; grid.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText="Введите целое положительное количество. Сумма должна помещаться в число."; }
            else grid.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText="";
        };
        grid.CellValueChanged += (_, e) =>
        {
            if(updatingPieces || e.RowIndex<0 || e.ColumnIndex<0 || !grid.Columns[e.ColumnIndex].Name.StartsWith("Piece")) return;
            var row=grid.Rows[e.RowIndex];
            if(row.Tag is not int[] pieces) return;
            int number=int.Parse(grid.Columns[e.ColumnIndex].Name[5..]);
            if(number>pieces.Length || !int.TryParse(Convert.ToString(row.Cells[e.ColumnIndex].Value),out int value) || value<=0) return;
            var updated=pieces.ToArray(); updated[number-1]=value;
            if(updated.Sum(n=>(long)n)<=int.MaxValue) SetPieces(row,updated);
        };
        grid.CellValueChanged += (_, _) => grid.Invalidate();
        grid.RowPostPaint += (_, e) =>
        {
            var row=grid.Rows[e.RowIndex];
            if(boards<=0 || !StockCheck.IsShort(Convert.ToString(row.Cells["Package"].Value),row.Cells["Total"].Value as decimal?)) return;
            using var pen=new Pen(Appearance.Ink,3);
            int right=Math.Min(e.RowBounds.Right,grid.ClientSize.Width-(grid.Controls.OfType<VScrollBar>().Any(s=>s.Visible)?SystemInformation.VerticalScrollBarWidth:0));
            e.Graphics.DrawRectangle(pen,new Rectangle(e.RowBounds.Left+2,e.RowBounds.Top+1,Math.Max(1,right-e.RowBounds.Left-4),e.RowBounds.Height-3));
        };
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
        var settings = new Button { Text = "Настройки", AutoSize = true };
        settings.Click += (_, _) => Run(() =>
        {
            using var prompt = new ReserveSettingsForm(reserveRules);
            if (prompt.ShowDialog(this) != DialogResult.OK) return;
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, System.Text.Json.JsonSerializer.Serialize(prompt.Rules));
            reserveRules = prompt.Rules; RecalculateTotals();
        });
        root.Controls.Add(Toolbar(load, new Label { Text = "Изделие", AutoSize = true, Padding = new Padding(2, 7, 2, 0) }, product, package, batch, reserve, settings), 0, 1);
        var nativePrint=new Button {Text="Печать без Word…",AutoSize=true};
        nativePrint.Click+=(_,_)=>Run(()=>Generate(null,true));
        var nativeRow=new Button {Text="Строка без Word…",AutoSize=true};
        nativeRow.Click+=(_,_)=>Run(()=>
        {
            var row=grid.CurrentRow??throw new InvalidOperationException("Выберите строку.");Generate(row,true);
        });
        var orderLabel=new Button {Text="Этикетка заказа без Word…",AutoSize=true};
        orderLabel.Click+=(_,_)=>Run(()=>
        {
            if(source==null) throw new InvalidOperationException("Сначала загрузите BOM.");
            var name=product.Text.Trim();
            if(name.EndsWith("BOM",StringComparison.OrdinalIgnoreCase)) name=name[..^3].TrimEnd('_','-',' ');
            if(name.Length==0) throw new InvalidDataException("Введите название изделия.");
            using var previewOrder=new NativePrintForm(name,boards);previewOrder.ShowDialog(this);
        });
        root.Controls.Add(Toolbar(nativePrint,nativeRow,orderLabel,save,preview,print,single),0,2);
        var excel = new Button { Text = "Сохранить BOM в Excel…", AutoSize = true };
        excel.Click += (_, _) => Run(() => ExportBom());
        Button ColorButton(string text, Color color)
        {
            var button = new Button { Text = text, AutoSize = true };
            button.Click += (_, _) =>
            {
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    row.DefaultCellStyle.BackColor = color;
                    if(color.ToArgb()==Color.FromArgb(221,235,247).ToArgb())
                    {
                        string note=Convert.ToString(row.Cells["Notes"].Value)??"";
                        if(!note.Contains("Давальческая позиция",StringComparison.OrdinalIgnoreCase))
                            row.Cells["Notes"].Value=note.Length==0 ? "Давальческая позиция" : note+"; Давальческая позиция";
                    }
                    row.DefaultCellStyle.SelectionBackColor = color.IsEmpty ? grid.DefaultCellStyle.SelectionBackColor : ControlPaint.Dark(color, 0.05f);
                }
            };
            return button;
        }
        var subset = new Button {Text="Сохранить часть BOM…",AutoSize=true};
        subset.Click+=(_,_)=>Run(()=>
        {
            if(source==null) throw new InvalidOperationException("Сначала загрузите BOM.");
            using var pick=new Picker("Какие строки сохранить в отдельный Excel?",["Все строки","Жёлтые — в процессе","Зелёные — готово","Красные — дефицит","Выделенные вручную (Ctrl/Shift)","Только светло-синие (давальческие)","Фиолетовые — из наличия"]);
            if(pick.ShowDialog(this)==DialogResult.OK) ExportBom(pick.Index>=5 ? pick.Index+1 : pick.Index,true);
        });
        var shortage=new Button {Text="Сохранить дефицит…",AutoSize=true};
        shortage.Click+=(_,_)=>Run(()=>ExportBom(5,true));
        var rowActions = Toolbar(addRows, deleteRows, sortLine, splitPackages, deletePiece, excel, subset, shortage,
            ColorButton("Зелёный — готово", Color.FromArgb(198,239,206)), ColorButton("Жёлтый — в процессе", Color.FromArgb(255,235,156)),
            ColorButton("Красный — дефицит", Color.FromArgb(255,199,206)), ColorButton("Давальческая позиция", Color.FromArgb(221,235,247)), ColorButton("Фиолетовый — из наличия", Color.FromArgb(228,215,245)), ColorButton("Без заливки", Color.Empty));
        foreach (Button button in rowActions.Controls)
            if (button.Text == "Зелёный — готово") button.BackColor = Color.FromArgb(198,239,206);
            else if (button.Text == "Жёлтый — в процессе") button.BackColor = Color.FromArgb(255,235,156);
            else if (button.Text == "Красный — дефицит") button.BackColor = Color.FromArgb(255,199,206);
            else if(button.Text=="Давальческая позиция") button.BackColor=Color.FromArgb(221,235,247);
            else if(button.Text=="Фиолетовый — из наличия") button.BackColor=Color.FromArgb(228,215,245);
        root.Controls.Add(rowActions, 0, 3);
        Appearance.Table(grid);
        var table = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = Appearance.Border, Margin = new Padding(0) };
        table.Controls.Add(grid); root.Controls.Add(table, 0, 4);
        status.AutoSize = true; status.Dock = DockStyle.Fill;
        status.Padding = new Padding(4, 12, 4, 4); status.Margin = new Padding(0); status.ForeColor = Color.FromArgb(79, 99, 111);
        root.Controls.Add(status, 0, 5);
        Controls.Add(root); ResumeLayout(true);
    }
    void Run(Action action) { try { action(); } catch (Exception e) { MessageBox.Show(this, e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); } }
    string CaptureState()
    {
        return System.Text.Json.JsonSerializer.Serialize(new
        {
            Product = product.Text, Boards = boards, Reserve = reserve.Checked, Package = package.Checked, Rules = reserveRules,
            Columns = grid.Columns.Cast<DataGridViewColumn>().Where(c => !c.Name.StartsWith("Piece") || c.Visible).Select(c => new { c.Name, c.FillWeight }).ToArray(),
            Rows = grid.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => new
            {
                Values = new[] { "Line", "Designator", "Part", "Quantity", "Summary", "Package", "Notes" }.Select(key => Convert.ToString(r.Cells[key].Value) ?? "").ToArray(),
                Pieces = r.Tag as int[] ?? [], Color = r.DefaultCellStyle.BackColor.ToArgb()
            }).ToArray()
        });
    }
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        if(e.Cancel || source == null) return;
        try
        {
            bool committed = grid.EndEdit();
            if(committed && savedState == CaptureState()) return;
            var answer = MessageBox.Show(this,"Есть несохранённые изменения. Сохранить редакцию BOM_EDIT перед закрытием?", "Сохранить BOM_EDIT?",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if(answer == DialogResult.Cancel) e.Cancel=true;
            else if(answer == DialogResult.Yes) e.Cancel=!ExportBom();
        }
        catch(Exception ex)
        {
            e.Cancel=true; MessageBox.Show(this,ex.Message,"Не удалось сохранить BOM_EDIT",MessageBoxButtons.OK,MessageBoxIcon.Error);
        }
    }
    void LoadBom() => Run(() =>
    {
        using var dialog = new OpenFileDialog() { Filter = "Excel BOM (*.xlsx)|*.xlsx", Title = "Выберите BOM" }; if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var sheets = Bom.Read(dialog.FileName); if (sheets.Count == 0) throw new InvalidDataException("В книге нет непустых листов.");
        var session = BomSession.Read(dialog.FileName);
        if (session != null || (Path.GetFileNameWithoutExtension(dialog.FileName).Contains("_EDIT", StringComparison.OrdinalIgnoreCase) && sheets[0].Rows[0].Contains("PartNumber") && sheets[0].Rows[0].Contains("Summary")))
        {
            RestoreBom(dialog.FileName, sheets[0], session); return;
        }
        int si = 0; if (sheets.Count > 1) { using var pick = new Picker("Выберите лист BOM", sheets.Select(s => s.Name).ToArray()); if (pick.ShowDialog(this) != DialogResult.OK) return; si = pick.Index; }
        var sheet = sheets[si];
        int hi = sheet.Rows.FindIndex(r => r.Any(c => c.Equals("Designator", StringComparison.OrdinalIgnoreCase)) && r.Any(c => c.Equals("Quantity", StringComparison.OrdinalIgnoreCase)));
        using var mapping = new ColumnMapping(sheet, Math.Max(hi, 0));
        if (mapping.ShowDialog(this) != DialogResult.OK) return;
        int noteColumn=Array.FindIndex(sheet.Rows[mapping.HeaderRow],h=>h.Trim().Equals("Примечание",StringComparison.OrdinalIgnoreCase) || h.Trim().Equals("Notes",StringComparison.OrdinalIgnoreCase));
        var records = Bom.ProjectRows(sheet, mapping.HeaderRow, mapping.Map.Append(noteColumn).ToArray());
        if (records.Count == 0) throw new InvalidDataException("После заголовков нет позиций BOM.");
        editingSavedBom = false;
        boards = 0; reserve.Checked = false; reserve.Enabled = false; batch.Enabled = true;
        batch.Text = "Ввести общее кол-во"; grid.Columns["Total"].Visible = false;
        source = dialog.FileName; product.Text = Bom.Product(source); grid.Rows.Clear();
        foreach(DataGridViewColumn c in grid.Columns) if(c.Name.StartsWith("Piece")) c.Visible=false;
        foreach (var r in records)
        {
            int index=grid.Rows.Add(r[0],r[1],r[2],r[3],r[4],r[5]); grid.Rows[index].Cells["Notes"].Value=r[6];
        }
        package.Checked = mapping.Map[5] >= 0;
        Appearance.FitColumns(grid);
        status.Text = $"Загружено позиций: {records.Count}. " + (package.Checked ? "Проверьте значения «Упк, шт» из BOM." : "Проверьте данные перед генерацией.");
        savedState = CaptureState();
    });
    void UpdateSplitAction()
    {
        splitPackages.Enabled = source != null && grid.SelectedRows.Count == 1;
        deletePiece.Enabled = splitPackages.Enabled && grid.SelectedRows[0].Tag is int[] { Length: >0 };
        splitPackages.Text = grid.SelectedRows.Count==1 && grid.SelectedRows[0].Tag is int[] { Length: >0 }
            ? "Добавить отрезок" : "Разбить упаковку";
    }
    void SetPieces(DataGridViewRow row, int[] pieces)
    {
        updatingPieces=true;
        try
        {
        row.Tag = pieces;
        row.Cells["Package"].ReadOnly = pieces.Length > 0;
        if(pieces.Length > 0) row.Cells["Package"].Value = pieces.Sum();
        for(int i = 0; i < pieces.Length; i++)
        {
            string key = "Piece" + (i+1);
            if(!grid.Columns.Contains(key)) grid.Columns.Add(new DataGridViewTextBoxColumn { Name=key, HeaderText="Упк " + (i+1), ReadOnly=false, FillWeight=70 });
        }
        foreach(DataGridViewColumn column in grid.Columns)
            if(column.Name.StartsWith("Piece") && int.TryParse(column.Name[5..],out int number))
                row.Cells[column.Index].Value = number <= pieces.Length ? pieces[number-1] : null;
        foreach(DataGridViewColumn column in grid.Columns)
            if(column.Name.StartsWith("Piece")) column.Visible = grid.Rows.Cast<DataGridViewRow>().Any(r=>r.Cells[column.Index].Value!=null);
        }
        finally { updatingPieces=false; grid.Columns["Notes"].DisplayIndex=grid.Columns.Count-1; UpdateSplitAction(); }
    }
    void RestoreBom(string path, Sheet sheet, BomSession? state)
    {
        bool legacy = state == null;
        editingSavedBom = true;
        var headers = sheet.Rows[0]; state ??= BomSession.Legacy(path, headers);
        string[] names = ["Line #", "Designator", "PartNumber", "Quantity", "Summary", "Упк, шт"];
        var map = names.Select(n => Array.IndexOf(headers, n)).ToArray();
        if (map.Take(4).Any(i => i < 0)) throw new InvalidDataException("В редакции BOM отсутствуют обязательные столбцы.");
        var colors = BomSession.Colors(path);
        var records = Bom.ProjectRows(sheet, 0, map);
        boards = 0; grid.Rows.Clear();
        foreach(DataGridViewColumn c in grid.Columns) if(c.Name.StartsWith("Piece")) c.Visible=false;
        source = path; product.Text = state.Product;
        reserveRules = state.Rules;
        for (int i = 0; i < records.Count; i++)
        {
            int index = grid.Rows.Add(records[i].Cast<object>().ToArray());
            int noteColumn=Array.IndexOf(headers,"Примечание");
            if(noteColumn>=0 && noteColumn<sheet.Rows[i+1].Length) grid.Rows[index].Cells["Notes"].Value=sheet.Rows[i+1][noteColumn];
            var pieces = sheet.Rows[i+1].Select((value, c)=>(value,c)).Where(x => x.c < headers.Length && System.Text.RegularExpressions.Regex.IsMatch(headers[x.c], @"^Упк \d+$") && x.value.Length>0).Select(x=>x.value).ToArray();
            if(pieces.Length>0) SetPieces(grid.Rows[index],PackagePieces.Parse(string.Join(";",pieces)));
            if (i < colors.Length && colors[i] is string color && color.Length == 8 && uint.TryParse(color, System.Globalization.NumberStyles.HexNumber, null, out uint argb))
            {
                var fill = Color.FromArgb(unchecked((int)argb));
                grid.Rows[index].DefaultCellStyle.BackColor = fill;
                grid.Rows[index].DefaultCellStyle.SelectionBackColor = ControlPaint.Dark(fill, 0.05f);
            }
        }
        package.Checked = legacy ? records.Any(r => r[5].Length > 0) : state.Package;
        boards = state.Boards; reserve.Checked = state.Reserve; reserve.Enabled = boards > 0; batch.Enabled = true;
        grid.Columns["Total"].Visible = boards > 0;
        batch.Text = boards > 0 ? $"Плат: {boards} · изменить" : "Ввести общее кол-во";
        RecalculateTotals(); Appearance.FitColumns(grid);
        status.Text = $"Открыта редакция BOM: {records.Count} позиций. Заливки восстановлены.";
        savedState = CaptureState();
    }
    void RecalculateTotals()
    {
        if (boards <= 0) return;
        reserve.Text = "Прибавить тех запас (по группам)";
        grid.Columns["Total"].HeaderText = reserve.Checked ? $"Общее с техзапасом ({boards} плат)" : $"Общее кол-во ({boards} плат)";
        foreach (DataGridViewRow row in grid.Rows)
        {
            var percent = reserve.Checked ? reserveRules.PercentFor(Convert.ToString(row.Cells["Designator"].Value) ?? "") : 0;
            row.Cells["Total"].ToolTipText = $"Плат: {boards}. Техзапас: {percent}%. Округление вверх.";
            var total = BatchQuantity.Calculate(Convert.ToString(row.Cells["Quantity"].Value) ?? "", boards, percent);
            row.Cells["Total"].Value = total;
            row.Cells["Total"].ErrorText = total == null ? "Введите положительное количество на одну плату; результат должен помещаться в число." : "";
        }
    }
    bool ExportBom(int filter = 0, bool separate = false)
    {
        if(!grid.EndEdit()) throw new InvalidDataException("Исправьте количество в упаковке перед сохранением.");
        if (source == null) throw new InvalidOperationException("Сначала загрузите BOM.");
        RecalculateTotals();
        if(filter==5 && boards<=0) throw new InvalidOperationException("Сначала введите число плат для расчёта общего количества.");
        var exportRows=grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow && r.DataGridView==grid).Where(row=>filter switch
        {
            1=>row.DefaultCellStyle.BackColor.ToArgb()==Color.FromArgb(255,235,156).ToArgb(),
            2=>row.DefaultCellStyle.BackColor.ToArgb()==Color.FromArgb(198,239,206).ToArgb(),
            3=>row.DefaultCellStyle.BackColor.ToArgb()==Color.FromArgb(255,199,206).ToArgb(),
            4=>row.Selected,
            7=>row.DefaultCellStyle.BackColor.ToArgb()==Color.FromArgb(228,215,245).ToArgb(),
            6=>row.DefaultCellStyle.BackColor.ToArgb()==Color.FromArgb(221,235,247).ToArgb(),
            5=>StockCheck.IsShort(Convert.ToString(row.Cells["Package"].Value),row.Cells["Total"].Value as decimal?),
            _=>true
        }).ToArray();
        if(separate && exportRows.Length==0) throw new InvalidOperationException(filter==5 ? "Дефицита нет: строк с нехваткой компонентов не найдено." : "Нет строк для выбранного варианта сохранения.");
        string suffix=separate ? new[]{"_COPY","_YELLOW","_GREEN","_RED","_SELECTED","_DEFICIT","_CUSTOMER","_PURPLE"}[filter] : "";
        using var dialog = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", DefaultExt = "xlsx", AddExtension = true,
            FileName = Path.GetFileNameWithoutExtension(source).Replace("_EDIT", "", StringComparison.OrdinalIgnoreCase) + suffix + "_EDIT.xlsx", InitialDirectory = Path.GetDirectoryName(source), OverwritePrompt = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return false;
        if (Path.GetFullPath(dialog.FileName).Equals(Path.GetFullPath(source), StringComparison.OrdinalIgnoreCase) && (!editingSavedBom || separate))
            throw new InvalidOperationException("Выберите другое имя файла, чтобы сохранить исходный BOM.");
        RecalculateTotals();
        var columns = grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Name != "SplitPackages" && (c.Name != "Total" || boards > 0) && (!c.Name.StartsWith("Piece") || c.Visible)).OrderBy(c => c.DisplayIndex).ToArray();
        var headers = columns.Select(c => c.Name switch { "Line" => "Line #", "Designator" => "Designator", "Part" => "PartNumber", "Quantity" => "Quantity", "Summary" => "Summary", "Package" => "Упк, шт", _ => c.HeaderText }).ToArray();
        int shortageIndex=Array.FindIndex(columns,c=>c.Name=="Package")+1;
        if(filter==5) headers=headers.Take(shortageIndex).Concat(new[]{"Не хватает, шт"}).Concat(headers.Skip(shortageIndex)).ToArray();
        var rows = exportRows.Select(row =>
        {
            var color = row.DefaultCellStyle.BackColor;
            string? fill = color.IsEmpty ? null : "FF" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            var values=columns.Select(c => Convert.ToString(row.Cells[c.Index].Value) ?? "").ToArray();
            if(filter==5)
            {
                var missing=StockCheck.Missing(Convert.ToString(row.Cells["Package"].Value),row.Cells["Total"].Value as decimal?);
                values=values.Take(shortageIndex).Concat(new[]{missing!.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}).Concat(values.Skip(shortageIndex)).ToArray();
            }
            return new BomExportRow(values, fill);
        }).ToArray();
        var temp = Path.Combine(Path.GetDirectoryName(dialog.FileName)!, Guid.NewGuid() + ".xlsx");
        try
        {
            BomExport.Save(temp, rows, headers);
            BomSession.Write(temp, new BomSession { Product = product.Text, Boards = boards, Reserve = reserve.Checked, Package = package.Checked,
                Rules = reserveRules, Widths = grid.Columns.Cast<DataGridViewColumn>().ToDictionary(c => c.Name, c => c.FillWeight) });
            File.Move(temp, dialog.FileName, true);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
        status.Text = "Редакция BOM сохранена: " + dialog.FileName;
        if(!separate) savedState = CaptureState();
        return true;
    }
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
    string? Generate(DataGridViewRow? singleRow = null, bool direct = false)
    {
        if(!grid.EndEdit()) throw new InvalidDataException("Исправьте количество в упаковке перед печатью."); Validate();
        if (source is null) throw new InvalidOperationException("Сначала выберите BOM.");
        if (grid.Rows.Count == 0) throw new InvalidDataException("В таблице нет позиций. Добавьте строки или загрузите BOM.");
        bool FitText(string text,int lines,bool bold=false,float font=8) => direct ? NativeLabelRenderer.Fits(text,lines,bold,font) : Fits(text,lines,bold,font);
        var title = product.Text.Trim(); if (title.Length == 0) throw new InvalidDataException("Введите название изделия.");
        int? chosenPiece = null;
        if(singleRow?.Tag is int[] selectedPieces && selectedPieces.Length>0 && package.Checked)
        {
            using var pick = new Picker("Какие упаковки печатать?", new[]{"Все упаковки"}.Concat(selectedPieces.Select((n,i)=>$"Упк {i+1}: {n} шт")).ToArray());
            if(pick.ShowDialog(this)!=DialogResult.OK) return null;
            if(pick.Index>0) chosenPiece=pick.Index-1;
        }
        var labels = new List<BomLabel>(); int trimmed = 0;
        var selectedRows = singleRow == null ? grid.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow && r.DataGridView == grid).ToArray() : new[] { singleRow };
        foreach (var row in selectedRows)
        {
            string V(string key) => Convert.ToString(row.Cells[key].Value)?.Trim() ?? "";
            string line = V("Line"), des = V("Designator"), part = V("Part"), qty = V("Quantity"), sum = V("Summary"), pack = V("Package");
            if (new[] { line, des, part, qty }.Any(string.IsNullOrWhiteSpace)) throw new InvalidDataException($"Позиция {row.Index + 1}: заполните Line #, Обозн., Парт и Кол-во.");
            if (!decimal.TryParse(qty.Replace(',', '.'), System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var q) || q <= 0) throw new InvalidDataException($"Позиция {row.Index + 1}: «Кол-во» должно быть положительным числом.");
            if (package.Checked && pack.Length > 0 && (!int.TryParse(pack, out var n) || n <= 0)) throw new InvalidDataException($"Line # {line}: «Упк, шт» должно быть целым положительным числом или пустым.");
            if (!FitText(Document.Heading(title, line), 2, true))
            {
                int limit = title.Length; while (limit > 0 && !FitText(Document.Heading(title[..limit], line), 2, true)) limit--;
                if (limit == 0) throw new InvalidDataException($"Позиция {row.Index + 1}: номер Line # слишком длинный для этикетки.");
                using var prompt = new TextPrompt($"Название изделия не помещается. Сократите его (для текущего текста до {limit} символов).", title, limit);
                if (prompt.ShowDialog(this) != DialogResult.OK) return null;
                product.Text = prompt.Value; return Generate(singleRow,direct);
            }
            foreach (var (name, value) in new[] { ("Кол-во", qty), ("Упк, шт", package.Checked ? pack : "") }) if (!FitText(value, 1)) throw new InvalidDataException($"Позиция {row.Index + 1}: «{name}» не помещается. Сократите значение в таблице.");
            var pieces = package.Checked ? row.Tag as int[] : null;
            string?[] packages = pieces is {Length: >0} ? pieces.Where((n,i)=>chosenPiece==null || i==chosenPiece).Select(n=>(string?)n.ToString()).ToArray() : [package.Checked && pack.Length>0 ? pack : null];
            foreach(var piece in packages) if(piece!=null && !FitText(piece,1)) throw new InvalidDataException($"Line # {line}: количество в упаковке не помещается.");
            var original = new BomLabel(line, des, part, qty, sum, packages[0]);
            bool FitDesignation(string text, float font, int lines) => FitText(text, lines, false, font);
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
                foreach(var piece in packages) labels.Add(partLabel with { Summary = shortSum, Package = piece });
            }
        }
        if(direct)
        {
            using var preview=new NativePrintForm(title,labels);preview.ShowDialog(this);return null;
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
