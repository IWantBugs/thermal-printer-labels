using System.Diagnostics;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using LabelCore;
namespace LabelApp;
internal static class AutoUpdater
{
    public static async Task InstallAsync(Form owner,UpdateInfo info)
    {
        if(info.Sha256 is not {Length:64} || !info.Sha256.All(Uri.IsHexDigit)) throw new InvalidDataException("Для этой сборки нет проверочной суммы. Обновите приложение вручную.");
        string target=Environment.ProcessPath ?? throw new InvalidOperationException("Не найден путь EXE.");
        string folder=Path.Combine(Path.GetTempPath(),"SPK-Update-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        using var cancel=new CancellationTokenSource();
        using var progress=new Form {Text="Обновление SPK",Size=new Size(460,170),StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false};
        var label=new System.Windows.Forms.Label {Dock=DockStyle.Top,Height=55,Padding=new Padding(12),Text="Скачивание обновления…"};
        var bar=new ProgressBar {Dock=DockStyle.Top,Style=ProgressBarStyle.Marquee};
        var stop=new Button {Text="Отмена",Dock=DockStyle.Bottom};stop.Click+=(_,_)=>cancel.Cancel();
        progress.FormClosing+=(_,_)=>cancel.Cancel();progress.Controls.AddRange([bar,label,stop]);Appearance.Dialog(progress);progress.Show(owner);
        bool handedOff=false;
        try
        {
            string zipPath=Path.Combine(folder,"update.zip"),exe=Path.Combine(folder,"ThermalPrinterLabels.exe");
            using var client=new HttpClient {Timeout=TimeSpan.FromMinutes(10)};
            using var response=await client.GetAsync(info.DownloadUrl,HttpCompletionOption.ResponseHeadersRead,cancel.Token);response.EnsureSuccessStatusCode();
            long? size=response.Content.Headers.ContentLength;
            if(size>256L*1024*1024) throw new InvalidDataException("Размер обновления превышает допустимый.");
            using(var input=await response.Content.ReadAsStreamAsync(cancel.Token))
            using(var output=File.Create(zipPath))
            {
                byte[] buffer=new byte[81920];long received=0;int read;
                while((read=await input.ReadAsync(buffer,cancel.Token))>0)
                {
                    received+=read;if(received>256L*1024*1024)throw new InvalidDataException("Размер обновления превышает допустимый.");
                    await output.WriteAsync(buffer.AsMemory(0,read),cancel.Token);
                    label.Text=size>0?$"Скачивание обновления: {received*100/size}%":"Скачивание обновления…";
                }
            }
            label.Text="Проверка сборки…";
            using(var input=File.OpenRead(zipPath))
                if(!Convert.ToHexString(await SHA256.HashDataAsync(input,cancel.Token)).Equals(info.Sha256,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Проверочная сумма не совпала. Обновление не установлено; попробуйте позже.");
            using(var zip=ZipFile.OpenRead(zipPath))
            {
                var entry=zip.GetEntry("ThermalPrinterLabels.exe")??throw new InvalidDataException("В архиве нет EXE.");
                if(entry.Length>512L*1024*1024)throw new InvalidDataException("Некорректный размер EXE.");
                entry.ExtractToFile(exe);
            }
            if(!Version.TryParse(FileVersionInfo.GetVersionInfo(exe).FileVersion,out var version) || !info.Matches(version))throw new InvalidDataException("Версия EXE не совпадает с версией обновления.");
            string probe=Path.Combine(Path.GetDirectoryName(target)!,Guid.NewGuid()+".tmp");
            try {using(var file=File.Create(probe)){} } finally {if(File.Exists(probe))File.Delete(probe);}
            cancel.Token.ThrowIfCancellationRequested();
            var main=Application.OpenForms.OfType<MainForm>().FirstOrDefault()??throw new InvalidOperationException("Основное окно закрыто.");
            string ready=Path.Combine(folder,"ready"),abort=Path.Combine(folder,"cancel");
            string script=Path.Combine(folder,"install.ps1"),config=Path.Combine(folder,"config.json");
            File.WriteAllText(script,Installer,new UTF8Encoding(true));
            File.WriteAllText(config,JsonSerializer.Serialize(new {Pid=Environment.ProcessId,Target=target,Source=exe,Ready=ready,Cancel=abort,Folder=folder,Started=Path.Combine(folder,"started")}));
            var start=new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"WindowsPowerShell","v1.0","powershell.exe")) {UseShellExecute=false,CreateNoWindow=true};
            foreach(var arg in new[]{"-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass","-File",script,"-Config",config})start.ArgumentList.Add(arg);
            using var helper=Process.Start(start)??throw new InvalidOperationException("Не удалось запустить установщик.");
            try
            {
                for(int i=0;i<50 && !File.Exists(Path.Combine(folder,"started"));i++)
                {
                    if(helper.HasExited) throw new InvalidOperationException("Установщик завершился до начала обновления.");
                    await Task.Delay(100,cancel.Token);
                }
                if(!File.Exists(Path.Combine(folder,"started"))) throw new InvalidOperationException("Установщик не отвечает.");
            }
            catch {File.WriteAllText(abort,"");throw;}
            handedOff=true;
            // Only a completed FormClosed authorizes replacing the running executable.
            FormClosedEventHandler closed=(_,_)=>File.WriteAllText(ready,"");
            main.FormClosed+=closed;
            progress.Hide();
            if(owner!=main)owner.Close();
            main.Close();
            if(!main.IsDisposed)
            {
                main.FormClosed-=closed;File.WriteAllText(abort,"");
            }
        }
        catch(OperationCanceledException) { }
        finally
        {
            if(!progress.IsDisposed)progress.Close();
            if(!handedOff)try {Directory.Delete(folder,true);}catch(IOException){}catch(UnauthorizedAccessException){}
        }
    }
    const string Installer="""
param([string]$Config)
$ErrorActionPreference='Stop'
$c=Get-Content -LiteralPath $Config -Encoding UTF8 -Raw | ConvertFrom-Json
[IO.File]::WriteAllText($c.Started,'')
$stage=$c.Target+'.update-'+[Guid]::NewGuid().ToString('N')
$backup=$c.Target+'.previous'
$replaced=$false
try {
    $limit=(Get-Date).AddMinutes(15)
    while($true) {
        if(Test-Path -LiteralPath $c.Cancel) {exit}
        $running=Get-Process -Id $c.Pid -ErrorAction SilentlyContinue
        if(!$running) {
            if(!(Test-Path -LiteralPath $c.Ready)) {exit}
            break
        }
        if((Get-Date) -gt $limit) {throw 'Истекло время ожидания закрытия программы.'}
        Start-Sleep -Milliseconds 300
    }
    Copy-Item -LiteralPath $c.Source -Destination $stage
    for($attempt=0;$attempt -lt 30;$attempt++) {
        try {[IO.File]::Replace($stage,$c.Target,$backup);$replaced=$true;break}
        catch {if($attempt -eq 29){throw};Start-Sleep -Milliseconds 500}
    }
    Start-Process -FilePath $c.Target -WorkingDirectory ([IO.Path]::GetDirectoryName($c.Target))
} catch {
    $failure=$_.Exception.Message
    if($replaced -and (Test-Path -LiteralPath $backup)) {
        try {[IO.File]::Replace($backup,$c.Target,$null);Start-Process -FilePath $c.Target}catch{}
    }
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.MessageBox]::Show('Не удалось установить обновление. '+$failure+' Резервный EXE: '+$backup,'Обновление SPK') | Out-Null
} finally {
    if(Test-Path -LiteralPath $stage){Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue}
    Remove-Item -LiteralPath $c.Folder -Recurse -Force -ErrorAction SilentlyContinue
}
""";
}
