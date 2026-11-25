$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$DllPath = Join-Path $Root "dist\SauloTools\SauloTools.dll"
$TestSrc = Join-Path $Root "tests\LogicTests.cs"
$TestOutDir = Join-Path $Root "tests\bin"
$TestExe = Join-Path $TestOutDir "LogicTests.exe"

if (-not (Test-Path $DllPath)) {
    throw "Execute build.ps1 antes de rodar os testes (DLL não encontrada: $DllPath)."
}

New-Item -ItemType Directory -Path $TestOutDir -Force | Out-Null
Copy-Item $DllPath (Join-Path $TestOutDir "SauloTools.dll") -Force

$Csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$Wpf = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF"
$VgCore = "C:\Program Files\Corel\CorelDRAW Graphics Suite 2020\Programs64\Assemblies\Corel.Interop.VGCore.dll"

& $Csc /nologo /platform:x64 /target:exe `
    "/out:$TestExe" `
    "/r:$DllPath" `
    "/l:$VgCore" `
    "/r:$Wpf\PresentationFramework.dll" `
    "/r:$Wpf\PresentationCore.dll" `
    "/r:$Wpf\WindowsBase.dll" `
    "/r:C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Xaml.dll" `
    $TestSrc

if ($LASTEXITCODE -ne 0) {
    throw "Falha ao compilar LogicTests.cs."
}

& $TestExe
if ($LASTEXITCODE -ne 0) {
    throw "Alguns testes falharam!"
}
