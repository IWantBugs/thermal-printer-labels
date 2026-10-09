using System.Runtime.InteropServices;
using LabelCore;
namespace LabelApp;

/// <summary>Validate actual Word pagination before replacing the user's ETC file.</summary>
internal static class WordLayout
{
    public static void Normalize(string path, IReadOnlyList<LabelCore.Label> labels)
    {
        int expectedLabels = labels.Count;
        dynamic? word = null, doc = null;
        try
        {
            var type = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Для формирования этикеток нужен установленный Microsoft Word.");
            word = Activator.CreateInstance(type)!;
            word.Visible = false;
            word.DisplayAlerts = 0;
            doc = word.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            if ((int)doc.Tables.Count != expectedLabels)
                throw new InvalidDataException("Word открыл не все таблицы этикеток.");
            dynamic section = doc.Sections[1].PageSetup;
            section.PageWidth = Points(58); section.PageHeight = Points(40);
            section.TopMargin = Points(2); section.BottomMargin = Points(2);
            section.LeftMargin = Points(1); section.RightMargin = Points(1);
            Release(section);
            for (int i = 1; i <= expectedLabels; i++)
            {
                dynamic table = doc.Tables[i];
                try
                {
                    table.AllowAutoFit = false;
                    table.TopPadding = 0f; table.BottomPadding = 0f;
                    table.LeftPadding = 1f; table.RightPadding = 1f;
                    table.Rows.AllowBreakAcrossPages = 0;
                    table.Range.Font.Name = "Arial"; table.Range.Font.Size = 8f;
                    dynamic paragraph = table.Range.ParagraphFormat;
                    paragraph.SpaceBefore = 0f; paragraph.SpaceAfter = 0f;
                    paragraph.LineSpacingRule = 4; // wdLineSpaceExactly
                    paragraph.LineSpacing = 9f;
                    paragraph.KeepTogether = -1; paragraph.KeepWithNext = -1;
                    paragraph.WidowControl = 0;
                    Release(paragraph);
                    int rows = table.Rows.Count;
                    for (int r = 1; r <= rows; r++)
                    {
                        dynamic row = table.Rows[r];
                        try
                        {
                            row.HeightRule = 2; // wdRowHeightExactly
                            row.Height = Points(r == 1 ? 8 : r == 2 ? Designators.Height(labels[i - 1]) : r == 3 ? CriticalFields.Height(labels[i - 1].PartFont, labels[i - 1].PartLines) : r == 5 && labels[i - 1].SummaryLines > 0 && !string.IsNullOrWhiteSpace(labels[i - 1].Summary) ? Designators.SummaryHeight(labels[i - 1]) : 3.5f);
                            if (r == 2 || r == 3)
                            {
                                dynamic cell = table.Cell(r, 2), range = cell.Range;
                                try
                                {
                                    var font = r == 2 ? labels[i - 1].DesignatorFont : labels[i - 1].PartFont;
                                    var allowedLines = r == 2 ? labels[i - 1].DesignatorLines : labels[i - 1].PartLines;
                                    range.Font.Size = font;
                                    range.ParagraphFormat.LineSpacing = font + 1;
                                    dynamic text = range.Duplicate;
                                    int lines;
                                    try
                                    {
                                        // Exclude the paragraph and end-of-cell marks from line statistics.
                                        text.End = Math.Max((int)text.Start, (int)text.End - 2);
                                        lines = text.ComputeStatistics(1);
                                    }
                                    finally { Release(text); }
                                    if (lines > allowedLines)
                                        throw new InvalidDataException($"Word переносит важное поле Line # {labels[i - 1].Line} на {lines} строк. Документ не сохранён, чтобы не обрезать «Обозн.» или «Парт».");
                                }
                                finally { Release(range); Release(cell); }
                            }
                            if (r == rows) row.Range.ParagraphFormat.KeepWithNext = 0;
                        }
                        finally { Release(row); }
                    }
                }
                finally { Release(table); }
            }
            doc.Repaginate();
            int pages = doc.ComputeStatistics(2); // wdStatisticPages
            if (pages != expectedLabels)
                throw new InvalidDataException($"Word сформировал {pages} страниц вместо {expectedLabels}. Документ не сохранён: проверьте настройки бумаги 58×40 мм в Word и драйвере принтера.");
            for (int i = 1; i <= expectedLabels; i++)
            {
                dynamic table = doc.Tables[i], start = table.Range.Duplicate, end = table.Range.Duplicate;
                try
                {
                    start.Collapse(1); // wdCollapseStart
                    end.End = end.End - 1; end.Collapse(0); // wdCollapseEnd
                    if ((int)start.Information[3] != i || (int)end.Information[3] != i)
                        throw new InvalidDataException($"Этикетка {i} перенесена между страницами. Документ не сохранён.");
                }
                finally { Release(start); Release(end); Release(table); }
            }
            doc.Save();
        }
        finally
        {
            try { if (doc != null) doc.Close(0); }
            finally
            {
                Release(doc);
                try { if (word != null) word.Quit(0); }
                finally { Release(word); }
            }
        }
    }
    static float Points(float millimeters) => millimeters * 72f / 25.4f;
    static void Release(object? value)
    {
        if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
