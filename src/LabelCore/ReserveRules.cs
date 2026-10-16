using System.Text.RegularExpressions;
namespace LabelCore;
public record ReserveGroup(string Prefixes, decimal Percent);
public sealed class ReserveRules
{
    public List<ReserveGroup> Groups { get; set; } = [new("R, C, VD, VT", 10), new("U, DA, DD", 5)];
    public decimal OtherPercent { get; set; } = 3;
    public static string[] Prefixes(string value) => Regex.Split(value.Trim().ToUpperInvariant(), @"[\s,;]+").Where(x => x.Length > 0).ToArray();
    public void Validate()
    {
        if (OtherPercent < 0 || OtherPercent > 1000 || Groups == null) throw new ArgumentException("Процент должен быть от 0 до 1000.");
        var used = new HashSet<string>();
        foreach (var group in Groups)
        {
            var prefixes = Prefixes(group.Prefixes ?? "");
            if (prefixes.Length == 0 || group.Percent < 0 || group.Percent > 1000) throw new ArgumentException("Заполните префиксы группы и процент от 0 до 1000.");
            foreach (var prefix in prefixes)
                if (!Regex.IsMatch(prefix, "^[A-Z]+$") || !used.Add(prefix)) throw new ArgumentException("Префиксы должны состоять из латинских букв и не повторяться: " + prefix);
        }
    }
    public decimal PercentFor(string designators)
    {
        var matches = Regex.Matches(designators.ToUpperInvariant(), @"(?<![A-Z])([A-Z]+)\s*\d+");
        if (matches.Count == 0) return OtherPercent;
        return matches.Select(m => Groups.FirstOrDefault(g => Prefixes(g.Prefixes).Contains(m.Groups[1].Value))?.Percent ?? OtherPercent).Max();
    }
}
