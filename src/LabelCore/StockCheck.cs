using System.Globalization;
namespace LabelCore;
public static class StockCheck
{
    public static bool IsShort(string? packaged, decimal? required) =>
        !string.IsNullOrWhiteSpace(packaged) && required > 0 &&
        decimal.TryParse(packaged.Trim().Replace(',', '.'), NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var available) && available >= 0 && available < required;
}
