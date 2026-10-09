using System.Globalization;
namespace LabelCore;

public static class Designators
{
    public static Label? Fit(Label label, Func<string, float, int, bool> fits)
    {
        foreach (var (font, lines, summary) in new[] { (8f, 1, 3), (8f, 2, 2), (8f, 3, 1), (8f, 4, 0), (7f, 5, 0) })
            if (fits(label.Designator, font, lines))
                return label with { DesignatorFont = font, DesignatorLines = lines, SummaryLines = summary };
        return null;
    }
    public static List<Label> Split(Label label, Func<string, float, int, bool> fits)
    {
        var result = new List<Label>();
        var remaining = label.Designator;
        while (remaining.Length > 0)
        {
            var positions = StringInfo.ParseCombiningCharacters(remaining);
            int low = 1, high = positions.Length, best = 0;
            while (low <= high)
            {
                int middle = (low + high) / 2;
                int end = middle == positions.Length ? remaining.Length : positions[middle];
                if (fits(remaining[..end], 7, 5)) { best = end; low = middle + 1; }
                else high = middle - 1;
            }
            if (best == 0) throw new InvalidDataException("Обозначение не помещается даже при шрифте 7 пт.");
            if (best < remaining.Length)
            {
                int boundary = remaining[..best].LastIndexOfAny([',', ';', ' ', '\n', '\t']);
                if (boundary > 0) best = boundary + 1;
            }
            var part = label with { Designator = remaining[..best] };
            result.Add(Fit(part, fits)!);
            remaining = remaining[best..];
        }
        for (int i = 0; i < result.Count; i++) result[i] = result[i] with { DesignatorPart = $"{i + 1}/{result.Count}" };
        return result;
    }
    public static float Height(Label label) => label.DesignatorLines switch { 1 => 3.5f, 2 => 7f, 3 => 10f, 4 => 13f, _ => 14.5f };
    public static float SummaryHeight(Label label) => label.SummaryLines switch { 3 => 10f, 2 => 6.5f, 1 => 3.5f, _ => 0f };
}
