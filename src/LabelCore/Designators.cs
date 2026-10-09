namespace LabelCore;
public static class Designators
{
    public static Label? Fit(Label label, Func<string, float, int, bool> fits) => CriticalFields.Fit(label, fits);
    public static List<Label> Split(Label label, Func<string, float, int, bool> fits) => CriticalFields.Split(label, fits);
    public static float Height(Label label) => CriticalFields.Height(label.DesignatorFont, label.DesignatorLines);
    public static float SummaryHeight(Label label) => label.SummaryLines switch { 3 => 10f, 2 => 6.5f, 1 => 3.5f, _ => 0f };
}
