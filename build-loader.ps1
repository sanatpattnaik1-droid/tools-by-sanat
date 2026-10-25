# Builds build\ToolsBySanatLoader.xlam from src\loader\.
# The loader is installed once and checks GitHub for new versions of ToolsBySanat.xlam.
# The version.json address comes from updater.config.json ("repo": "owner/name").
# With no repo set, the loader still opens the main add-in but never checks for updates.
# Needs Excel option "Trust access to the VBA project object model" (build PC only).

$ErrorActionPreference = 'Stop'

$root     = $PSScriptRoot
$buildDir = Join-Path $root 'build'
$out      = Join-Path $buildDir 'ToolsBySanatLoader.xlam'

$cfg = Get-Content (Join-Path $root 'updater.config.json') -Raw | ConvertFrom-Json
$versionUrl = if ($cfg.repo) { "https://raw.githubusercontent.com/$($cfg.repo)/main/version.json" } else { '__VERSION_URL__' }
$prefix = if ($cfg.repo) { "https://github.com/$($cfg.repo)/releases/download/" } else { '__DOWNLOAD_PREFIX__' }
$keyFile = Join-Path $root 'keys\update-public.xml'
$pubKey = if (Test-Path $keyFile) { (Get-Content $keyFile -Raw).Trim() } else { '__PUBLIC_KEY__' }
if ($pubKey -match '"') { throw 'The public key file must not contain quotes.' }

New-Item -ItemType Directory -Force $buildDir | Out-Null
if (Test-Path $out) {
    try { Remove-Item $out -Force } catch { throw "Cannot replace $out. Close Excel and retry." }
}

$stage = Join-Path $env:TEMP ('tbs_loader_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $stage | Out-Null
$mod = Join-Path $stage 'modLoader.bas'
(Get-Content (Join-Path $root 'src\loader\modLoader.bas') -Raw).Replace('__VERSION_URL__', $versionUrl).Replace('__DOWNLOAD_PREFIX__', $prefix).Replace('__PUBLIC_KEY__', $pubKey) |
    Set-Content $mod -Encoding Default
$thisWbCode = Get-Content (Join-Path $root 'src\loader\ThisWorkbook.bas') -Raw

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    [void]$wb.VBProject.VBComponents.Import($mod)
    $wb.VBProject.VBComponents.Item($wb.CodeName).CodeModule.AddFromString($thisWbCode)
    $wb.Title = 'Tools By Sanat loader'
    $wb.SaveAs($out, 55)   # xlOpenXMLAddIn
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    Remove-Item $stage -Recurse -ErrorAction SilentlyContinue
}

"Built $out"
"Version URL   : $versionUrl"
"Download prefix: $prefix"
"Public key    : $(if ($pubKey -like '__*') { 'NOT SET - the updater stays OFF until keys\update-public.xml exists (run new-signing-key.ps1)' } else { 'embedded (' + $pubKey.Length + ' characters)' })"
