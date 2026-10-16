namespace LabelCore;
public static class PackagePieces
{
    public static int[] Parse(string text)
    {
        var values = text.Split(new[] { '\r', '\n', ';', '+' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var pieces = values.Select(v => int.TryParse(v, out int n) && n > 0 ? n : throw new ArgumentException("Количество каждого куска должно быть целым положительным числом.")).ToArray();
        if (pieces.Sum(n => (long)n) > int.MaxValue) throw new ArgumentException("Сумма упаковок слишком велика.");
        return pieces;
    }
}
