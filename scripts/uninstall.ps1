$ErrorActionPreference = "Stop"

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Solicitando permissão de Administrador para remover de Program Files..." -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Wait
    exit 0
}

$base = "C:\Program Files\Corel"
if (Test-Path $base) {
    Get-ChildItem $base -Directory | ForEach-Object {
        $dest = Join-Path $_.FullName "Programs64\Addons\SauloTools"
        if (Test-Path $dest) {
            Remove-Item $dest -Recurse -Force
            Write-Host "Removido: $dest" -ForegroundColor Yellow
        }
    }
}
Write-Host "Desinstalação concluída. Reinicie o CorelDRAW." -ForegroundColor Green
