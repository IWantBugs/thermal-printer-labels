using System.Globalization;
namespace LabelCore;
public static class BatchQuantity
{
    public static decimal? Calculate(string value, int boards, decimal reservePercent)
    {
        if (boards <= 0 || reservePercent < 0 || !decimal.TryParse(value.Trim().Replace(',', '.'), NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var quantity) || quantity <= 0) return null;
        try { return decimal.Ceiling(quantity * boards * (1 + reservePercent / 100m)); }
        catch (OverflowException) { return null; }
    }
}
