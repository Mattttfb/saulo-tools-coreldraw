param(
    [string]$Configuration = "Release",
    [string]$CorelInstallDir = "C:\Program Files\Corel\CorelDRAW Graphics Suite 2020"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Csproj = Join-Path $Root "src\SauloTools\SauloTools.csproj"
$DistDir = Join-Path $Root "dist\SauloTools"

# Locate MSBuild
$MsBuildCandidates = @(
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\amd64\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe",
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
)
$MsBuild = $MsBuildCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $MsBuild) {
    throw "MSBuild.exe não encontrado neste PC."
}

Write-Host "Usando MSBuild: $MsBuild" -ForegroundColor Cyan
& $MsBuild $Csproj /t:Rebuild "/p:Configuration=$Configuration" /p:Platform=x64 "/p:CorelInstallDir=$CorelInstallDir" /nologo /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Falha na compilação do SauloTools (exit code $LASTEXITCODE)."
}

# Stage portable Addon package in dist\SauloTools
if (Test-Path $DistDir) { Remove-Item $DistDir -Recurse -Force }
New-Item -ItemType Directory -Path $DistDir -Force | Out-Null

$BuiltDll = Join-Path $Root "src\SauloTools\bin\x64\$Configuration\SauloTools.dll"
if (-not (Test-Path $BuiltDll)) {
    throw "DLL compilada não encontrada em: $BuiltDll"
}

Copy-Item $BuiltDll (Join-Path $DistDir "SauloTools.dll") -Force
Copy-Item (Join-Path $Root "addon\CorelDRW.addon") (Join-Path $DistDir "CorelDRW.addon") -Force
Copy-Item (Join-Path $Root "addon\AppUI.xslt") (Join-Path $DistDir "AppUI.xslt") -Force
Copy-Item (Join-Path $Root "addon\UserUI.xslt") (Join-Path $DistDir "UserUI.xslt") -Force

# Also copy installer/uninstaller into dist\ for portability across print-shop PCs
Copy-Item (Join-Path $Root "scripts\install.ps1") (Join-Path $Root "dist\install.ps1") -Force
Copy-Item (Join-Path $Root "scripts\uninstall.ps1") (Join-Path $Root "dist\uninstall.ps1") -Force

Write-Host ""
Write-Host "Build concluído com sucesso! Pacote portátil em:" -ForegroundColor Green
Write-Host "  $DistDir" -ForegroundColor Yellow
Get-ChildItem $DistDir | Select-Object Name, Length | Format-Table -AutoSize
