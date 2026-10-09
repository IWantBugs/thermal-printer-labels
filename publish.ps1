$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    dotnet publish src/LabelApp/LabelApp.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -o artifacts/windows-x64
    if ($LASTEXITCODE -ne 0) { throw "Ошибка сборки: $LASTEXITCODE" }
} finally { Pop-Location }
