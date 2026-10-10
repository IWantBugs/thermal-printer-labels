using System.Numerics;
namespace LabelApp;

internal static class LineOrder
{
    public static int Compare(string? left, string? right)
    {
        left = left?.Trim() ?? ""; right = right?.Trim() ?? "";
        if (left.Length == 0 || right.Length == 0)
            return left.Length == 0 ? (right.Length == 0 ? 0 : 1) : -1;
        bool ln = BigInteger.TryParse(left, out var l), rn = BigInteger.TryParse(right, out var r);
        if (ln && rn) return l.CompareTo(r);
        if (ln != rn) return ln ? -1 : 1;
        return StringComparer.OrdinalIgnoreCase.Compare(left, right);
    }
}
