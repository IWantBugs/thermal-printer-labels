using System.Globalization;
namespace LabelCore;
public static class StockCheck
{
    public static decimal? Missing(string? packaged, decimal? required)
    {
        if(required is null || required <= 0) return null;
        if(string.IsNullOrWhiteSpace(packaged)) return required;
        if(!decimal.TryParse(packaged.Trim().Replace(',', '.'), NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var available) || available < 0) return null;
        return Math.Max(0,required.Value-available);
    }
    public static bool IsShort(string? packaged, decimal? required) => Missing(packaged,required)>0;
}
