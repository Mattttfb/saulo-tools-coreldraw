param(
    [string]$DrawUiPath = "C:\Program Files\Corel\CorelDRAW Graphics Suite 2020\Draw\UIConfig\DrawUI.xml"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$AppXslt = Join-Path $Root "addon\AppUI.xslt"
$UserXslt = Join-Path $Root "addon\UserUI.xslt"

if (-not (Test-Path $DrawUiPath)) {
    throw "DrawUI.xml não encontrado em: $DrawUiPath"
}

Write-Host "1. Verificando ausência de colisão de GUIDs no DrawUI.xml original..." -ForegroundColor Cyan
$origText = [System.IO.File]::ReadAllText($DrawUiPath)
foreach ($g in @("7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d01", "7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d02", "7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03")) {
    if ($origText.Contains($g)) {
        throw "Colisão de GUID detectada: $g"
    }
}
Write-Host "   [PASS] Nenhum conflito de GUID." -ForegroundColor Green

Write-Host "2. Aplicando AppUI.xslt em memória sobre DrawUI.xml..." -ForegroundColor Cyan
$xslt1 = New-Object System.Xml.Xsl.XslCompiledTransform
$xslt1.Load($AppXslt)
$sw1 = New-Object System.IO.StringWriter
$xslt1.Transform($DrawUiPath, $null, $sw1)
$xmlAfterApp = $sw1.ToString()

$doc1 = New-Object System.Xml.XmlDocument
$doc1.LoadXml($xmlAfterApp)

$wpfNode = $doc1.SelectSingleNode("//uiConfig/items/itemData[@guid='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d02']")
$btnNode = $doc1.SelectSingleNode("//uiConfig/items/itemData[@guid='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03']")
$dockerNode = $doc1.SelectSingleNode("//uiConfig/dockers/dockerData[@guid='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d01']")
$menuNode = $doc1.SelectSingleNode("//commandBarData[@guid='3eaa9bbe-28fd-4672-9128-02974ee96332']/menu/item[@guidRef='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03']")

if (-not $wpfNode) { throw "itemData wpfhost não injetado!" }
if (-not $btnNode) { throw "itemData checkButton não injetado!" }
if (-not $dockerNode) { throw "dockerData Saulo Tools não injetado!" }
if (-not $menuNode) { throw "item no menu Janela > Janelas de encaixe não injetado!" }

Write-Host "   [PASS] wpfhost:    $($wpfNode.GetAttribute('hostedType'))" -ForegroundColor Green
Write-Host "   [PASS] checkBtn:   $($btnNode.GetAttribute('check'))" -ForegroundColor Green
Write-Host "   [PASS] dockerData: $($dockerNode.GetAttribute('userCaption'))" -ForegroundColor Green
Write-Host "   [PASS] Dockers menu item injetado com sucesso." -ForegroundColor Green

Write-Host "3. Aplicando UserUI.xslt (teste de idempotência)..." -ForegroundColor Cyan
$xslt2 = New-Object System.Xml.Xsl.XslCompiledTransform
$xslt2.Load($UserXslt)
$sr = New-Object System.IO.StringReader($xmlAfterApp)
$xr = [System.Xml.XmlReader]::Create($sr)
$sw2 = New-Object System.IO.StringWriter
$xslt2.Transform($xr, $null, $sw2)
$doc2 = New-Object System.Xml.XmlDocument
$doc2.LoadXml($sw2.ToString())

$menuCount = $doc2.SelectNodes("//commandBarData[@guid='3eaa9bbe-28fd-4672-9128-02974ee96332']/menu/item[@guidRef='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03']").Count
if ($menuCount -ne 1) {
    throw "Idempotência falhou: esperado 1 item no menu Dockers, encontrado $menuCount."
}
Write-Host "   [PASS] Idempotência verificada (exatamente 1 entrada no menu Dockers)." -ForegroundColor Green
Write-Host ""
Write-Host "VALIDAÇÃO XSLT CONCLUÍDA 100% COM SUCESSO!" -ForegroundColor Green
