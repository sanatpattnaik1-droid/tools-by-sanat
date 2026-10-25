# Installs the "Tools By Sanat" add-in for the current Windows user.
#   1. copies ToolsBySanat.xlam (main add-in) and ToolsBySanatLoader.xlam to %APPDATA%\Microsoft\AddIns
#   2. removes the "downloaded from the internet" block
#   3. registers ONLY the loader in Excel; the loader opens the main add-in and installs updates
# Excel must be closed. Run via Install.bat (double-click) or: powershell -File install.ps1

$ErrorActionPreference = 'Stop'

$main   = 'ToolsBySanat.xlam'
$loader = 'ToolsBySanatLoader.xlam'

function Find-File($name) {
    @((Join-Path $PSScriptRoot $name), (Join-Path $PSScriptRoot "build\$name")) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
$srcMain = Find-File $main
$srcLoader = Find-File $loader
if (-not $srcMain)   { throw "$main not found next to install.ps1." }
if (-not $srcLoader) { throw "$loader not found next to install.ps1." }

if (Get-Process excel -ErrorAction SilentlyContinue) {
    Write-Host 'Excel is running. Please close all Excel windows and run the installer again.' -ForegroundColor Yellow
    exit 1
}

$destDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
New-Item -ItemType Directory -Force $destDir | Out-Null
Copy-Item $srcMain   (Join-Path $destDir $main)   -Force
Copy-Item $srcLoader (Join-Path $destDir $loader) -Force
Unblock-File (Join-Path $destDir $main)
Unblock-File (Join-Path $destDir $loader)
Write-Host "Copied to $destDir"

# a download that was waiting for the next start is older than what was just installed
$pending = Join-Path $destDir 'ToolsBySanat_data\pending'
if (Test-Path $pending) { Remove-Item $pending -Recurse -Force -ErrorAction SilentlyContinue }

$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    [void]$xl.Workbooks.Add()   # AddIns needs a workbook context

    # earlier versions registered the main add-in directly; now the loader opens it
    foreach ($a in @($xl.AddIns | Where-Object { $_.Name -eq $main })) { $a.Installed = $false }

    $l = $xl.AddIns | Where-Object { $_.Name -eq $loader } | Select-Object -First 1
    if (-not $l) { $l = $xl.AddIns.Add((Join-Path $destDir $loader), $false) }
    $l.Installed = $true
    Write-Host 'Registered in Excel.'
}
finally {
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}

Write-Host 'Done. Open Excel and look for the "Tools By Sanat" tab.' -ForegroundColor Green
