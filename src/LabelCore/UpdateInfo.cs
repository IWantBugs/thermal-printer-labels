using System.Text.Json;
namespace LabelCore;
public sealed record UpdateInfo(string Version, string DownloadUrl)
{
    static System.Version Normalized(System.Version value) => new(value.Major,value.Minor,Math.Max(value.Build,0),Math.Max(value.Revision,0));
    public bool IsNewerThan(System.Version current) => Normalized(System.Version.Parse(Version)) > Normalized(current);
    public static UpdateInfo Parse(string json)
    {
        var info=JsonSerializer.Deserialize<UpdateInfo>(json,new JsonSerializerOptions {PropertyNameCaseInsensitive=true}) ?? throw new InvalidDataException("Пустой ответ сервера обновлений.");
        if(!System.Version.TryParse(info.Version,out _)) throw new InvalidDataException("Некорректная версия обновления.");
        if(!Uri.TryCreate(info.DownloadUrl,UriKind.Absolute,out var url) || url.Scheme!="https" || url.UserInfo.Length>0 ||
           !(url.Host.Equals("github.com",StringComparison.OrdinalIgnoreCase) && url.AbsolutePath.StartsWith("/IWantBugs/thermal-printer-labels/raw/",StringComparison.Ordinal) ||
             url.Host.Equals("raw.githubusercontent.com",StringComparison.OrdinalIgnoreCase) && url.AbsolutePath.StartsWith("/IWantBugs/thermal-printer-labels/",StringComparison.Ordinal)) ||
           !url.AbsolutePath.EndsWith(".zip",StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Некорректная ссылка на сборку.");
        return info;
    }
}
