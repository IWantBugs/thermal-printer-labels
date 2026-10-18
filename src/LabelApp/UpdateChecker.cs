using System.Diagnostics;
using System.Net.Http;
using LabelCore;
namespace LabelApp;
internal static class UpdateChecker
{
    const string Manifest="https://raw.githubusercontent.com/IWantBugs/thermal-printer-labels/main/distribution/latest.json";
    static readonly HttpClient Client=new() {Timeout=TimeSpan.FromSeconds(8),MaxResponseContentBufferSize=32768};
    static bool checking;
    public static async Task CheckAsync(Form owner,bool manual)
    {
        if(checking)
        {
            if(manual) MessageBox.Show(owner,"Проверка обновлений уже выполняется.","Обновления");
            return;
        }
        checking=true;
        try
        {
            using var request=new HttpRequestMessage(HttpMethod.Get,Manifest);
            request.Headers.CacheControl=new System.Net.Http.Headers.CacheControlHeaderValue {NoCache=true};
            using var response=await Client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            var info=UpdateInfo.Parse(await response.Content.ReadAsStringAsync());
            if(owner.IsDisposed || owner.Disposing) return;
            var current=typeof(UpdateChecker).Assembly.GetName().Version ?? new Version(0,0,0);
            if(!info.IsNewerThan(current))
            {
                if(manual) MessageBox.Show(owner,$"Установлена актуальная версия {current.ToString(3)}.","Обновления",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            var answer=MessageBox.Show(owner,$"Доступна версия {info.Version}.\nУстановлена версия {current.ToString(3)}.\n\nОткрыть скачивание обновления?\n\nПосле скачивания распакуйте архив, закройте приложение и замените EXE. Ваши BOM-файлы сохраняются отдельно.","Доступно обновление",MessageBoxButtons.YesNo,MessageBoxIcon.Information);
            if(answer==DialogResult.Yes) Process.Start(new ProcessStartInfo(info.DownloadUrl) {UseShellExecute=true});
        }
        catch(Exception ex) when(ex is HttpRequestException or TaskCanceledException or InvalidDataException or System.Text.Json.JsonException or System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            if(manual && !owner.IsDisposed) MessageBox.Show(owner,"Не удалось проверить или открыть обновление. Проверьте подключение к интернету и попробуйте позже.\n\n"+ex.Message,"Обновления",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
        finally {checking=false;}
    }
}
