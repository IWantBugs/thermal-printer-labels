namespace LabelApp;

internal static class ColumnDefaults
{
    public static int[] Resolve(string[] headers)
    {
        static string Normalize(string value) => new(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        var names = headers.Select(Normalize).ToArray();
        int Find(params string[] aliases)
        {
            foreach (var alias in aliases)
            {
                var matches = Enumerable.Range(0, names.Length).Where(i => names[i] == Normalize(alias)).ToArray();
                if (matches.Length > 0) return matches.Length == 1 ? matches[0] : -1;
            }
            return -1;
        }
        int part = Find("PartNumber", "PartNumb", "PartNum", "Name");
        int summary = Find("Summary", "Sum");
        // Older BOMs use Name for Sum; don't copy a fallback part number into Sum.
        bool explicitSummary = names.Any(n => n is "SUMMARY" or "SUM");
        if (summary < 0 && !explicitSummary)
        {
            int name = Find("Name");
            if (name != part) summary = name;
        }
        return [Find("Line #", "Line"), Find("Designator"), part, Find("Quantity"), summary];
    }
}
