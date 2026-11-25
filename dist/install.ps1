param(
    [string]$CorelDir = ""
)

$ErrorActionPreference = "Stop"

# Resolve package folder (works both from repo\scripts\ and from portable dist\)
$ScriptDir = $PSScriptRoot
$PkgDir = Join-Path $ScriptDir "SauloTools"
if (-not (Test-Path $PkgDir)) {
    $PkgDir = Join-Path (Split-Path -Parent $ScriptDir) "dist\SauloTools"
}
if (-not (Test-Path (Join-Path $PkgDir "SauloTools.dll"))) {
    throw "Pacote SauloTools não encontrado. Execute scripts\build.ps1 primeiro."
}

# Detect CorelDRAW 64-bit installations
$CorelTargets = @()
if ($CorelDir -and (Test-Path $CorelDir)) {
    $CorelTargets += $CorelDir
} else {
    $base = "C:\Program Files\Corel"
    if (Test-Path $base) {
        $CorelTargets = Get-ChildItem $base -Directory |
            Where-Object { Test-Path (Join-Path $_.FullName "Programs64\Addons") } |
            Select-Object -ExpandProperty FullName
    }
}

if ($CorelTargets.Count -eq 0) {
    throw "Nenhuma instalação 64-bit do CorelDRAW encontrada em C:\Program Files\Corel."
}

# Check if CorelDRAW is currently running
$running = Get-Process -Name "CorelDRW" -ErrorAction SilentlyContinue
if ($running) {
    Write-Warning "O CorelDRAW está aberto no momento. Feche o CorelDRAW antes de instalar ou reinicie-o após a instalação."
}

# Elevate if not running as Administrator (needed for C:\Program Files)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Solicitando permissão de Administrador para instalar em Program Files..." -ForegroundColor Yellow
    $argList = "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    if ($CorelDir) { $argList += " -CorelDir `"$CorelDir`"" }
    Start-Process powershell -Verb RunAs -ArgumentList $argList -Wait
    exit 0
}

# Backup user Workspaces safely (non-destructive)
$appDataCorel = Join-Path $env:APPDATA "Corel"
if (Test-Path $appDataCorel) {
    Get-ChildItem $appDataCorel -Directory | ForEach-Object {
        $ws = Join-Path $_.FullName "Draw\Workspace"
        if (Test-Path $ws) {
            $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $bak = Join-Path $_.FullName "Draw\Workspace_Backup_SauloTools_$stamp"
            Copy-Item $ws $bak -Recurse -Force
            Write-Host "Backup do Workspace salvo em: $bak" -ForegroundColor Cyan
        }
    }
}

foreach ($target in $CorelTargets) {
    $dest = Join-Path $target "Programs64\Addons\SauloTools"
    if (-not (Test-Path $dest)) {
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
    }
    Copy-Item (Join-Path $PkgDir "*") $dest -Force
    Write-Host "Instalado com sucesso em: $dest" -ForegroundColor Green
}

Write-Host ""
Write-Host "INSTALAÇÃO CONCLUÍDA!" -ForegroundColor Green
Write-Host "1. Reabra o CorelDRAW."
Write-Host "2. Vá em: Janela -> Janelas de encaixe -> Saulo Tools (ou clique no '+' da barra lateral direita)."
Write-Host "3. Seus arquivos .gms antigos NÃO foram alterados."
