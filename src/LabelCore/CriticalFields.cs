using System.Globalization;
namespace LabelCore;

public static class CriticalFields
{
    public static float Height(float font, int lines) => Math.Max(3.5f, (float)Math.Ceiling(lines * (font + 1) * 25.4 / 72 * 2) / 2);
    public static Label? Fit(Label label, Func<string, float, int, bool> fits)
    {
        var choices = new List<(float Font, int Lines)>[] { new(), new() };
        var texts = new[] { label.Designator, label.Part };
        for (int i = 0; i < texts.Length; i++)
            foreach (float font in new[] { 8f, 7f })
                for (int lines = 1; lines <= 5; lines++)
                    if (fits(texts[i], font, lines)) { choices[i].Add((font, lines)); break; }
        var candidates = from d in choices[0] from p in choices[1]
                         orderby (8 - d.Font) + (8 - p.Font), d.Font descending
                         select (d, p);
        float budget = label.Package is null ? 21.5f : 18f;
        foreach (var (d, p) in candidates)
            foreach (int summary in string.IsNullOrWhiteSpace(label.Summary) ? new[] { 0 } : new[] { 3, 2, 1, 0 })
            {
                var candidate = label with { DesignatorFont = d.Font, DesignatorLines = d.Lines, PartFont = p.Font, PartLines = p.Lines, SummaryLines = summary };
                if (Height(d.Font, d.Lines) + Height(p.Font, p.Lines) + Designators.SummaryHeight(candidate) <= budget)
                    return candidate;
            }
        return null;
    }
    public static List<Label> Split(Label label, Func<string, float, int, bool> fits)
    {
        List<string> Chunks(string text)
        {
            var result = new List<string>();
            while (text.Length > 0)
            {
                var positions = StringInfo.ParseCombiningCharacters(text);
                int low = 1, high = positions.Length, best = 0;
                while (low <= high)
                {
                    int middle = (low + high) / 2;
                    int end = middle == positions.Length ? text.Length : positions[middle];
                    if (fits(text[..end], 7, 2)) { best = end; low = middle + 1; } else high = middle - 1;
                }
                if (best == 0) throw new InvalidDataException("Поле не помещается даже при шрифте 7 пт.");
                if (best < text.Length)
                {
                    int boundary = text[..best].LastIndexOfAny([',', ';', ' ', '\n', '\t']);
                    if (boundary > 0) best = boundary + 1;
                }
                result.Add(text[..best]); text = text[best..];
            }
            return result;
        }
        var designations = Chunks(label.Designator); var parts = Chunks(label.Part);
        var labels = new List<Label>();
        for (int i = 0; i < Math.Max(designations.Count, parts.Count); i++)
        {
            int di = Math.Min(i, designations.Count - 1), pi = Math.Min(i, parts.Count - 1);
            var candidate = label with {
                Designator = designations[di], Part = parts[pi],
                DesignatorPart = designations.Count > 1 ? $"{di + 1}/{designations.Count}" : null,
                PartContinuation = parts.Count > 1 ? $"{pi + 1}/{parts.Count}" : null
            };
            labels.Add(Fit(candidate, fits) ?? throw new InvalidDataException("Часть важных полей не помещается на этикетке."));
        }
        return labels;
    }
}
