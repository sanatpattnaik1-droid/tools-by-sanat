# Tests dist\ToolsBySanat-Setup-<version>.exe on THIS PC: silent install, uninstall, reinstall.
# Close all Excel windows first. Your settings files are checked (hash) and must survive the uninstall.
# The PC ends up installed, as before. Run:  powershell -ExecutionPolicy Bypass -File test-setup.ps1

param([string]$Version)

$ErrorActionPreference = 'Continue'
if (Get-Process excel -ErrorAction SilentlyContinue) { Write-Host 'Excel is running - close all Excel windows first.' -ForegroundColor Yellow; exit 1 }

$r = $PSScriptRoot
if (-not $Version) { $Version = (Get-Content (Join-Path $r 'version.json') -Raw | ConvertFrom-Json).version }
$exe = "$r\dist\ToolsBySanat-Setup-$Version.exe"
if (-not (Test-Path $exe)) { Write-Host "Not found: $exe (run build-setup.ps1)" -ForegroundColor Yellow; exit 1 }

$add = "$env:APPDATA\Microsoft\AddIns"
$app = "$env:LOCALAPPDATA\Programs\ToolsBySanat"
$unKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{B1D0C7F2-5E3A-4C8B-9A64-2F7E1D3C8A51}_is1'
$opt = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$res = New-Object System.Collections.Generic.List[string]
function Check($n, $ok, $x = '') { $script:res.Add(("{0,-62} {1} {2}" -f $n, $(if ($ok) { 'PASS' } else { 'FAIL' }), $x)) }
function Hash($p) { if (Test-Path $p) { (Get-FileHash $p -Algorithm SHA256).Hash } else { '' } }
function OpenEntries { (Get-ItemProperty $opt -ErrorAction SilentlyContinue).PSObject.Properties | Where-Object { $_.Name -match '^OPEN' } | ForEach-Object { $_.Value } }
function SettingsHash { (Hash "$add\ToolsBySanat_data\settings.csv") + (Hash "$add\ToolsBySanat_data\party_master.csv") }

$settingsBefore = SettingsHash

# ---- 1. install ----
$p = Start-Process $exe -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', "/LOG=$env:TEMP\tbs_setup.log" -Wait -PassThru
Check 'install exit code 0' ($p.ExitCode -eq 0) "code $($p.ExitCode)"
Check 'main add-in copied (hash = build)' ((Hash "$add\ToolsBySanat.xlam") -eq (Hash "$r\build\ToolsBySanat.xlam"))
Check 'loader copied (hash = build)' ((Hash "$add\ToolsBySanatLoader.xlam") -eq (Hash "$r\build\ToolsBySanatLoader.xlam"))
Check 'uninstaller + register.ps1 in app folder' ((Test-Path "$app\unins000.exe") -and (Test-Path "$app\register.ps1"))
Check 'listed in Windows installed apps' (Test-Path $unKey) ((Get-ItemProperty $unKey -ErrorAction SilentlyContinue).DisplayName)
Check 'Excel starts the loader (registry OPEN entry)' ((OpenEntries) -match 'ToolsBySanatLoader\.xlam')
Check 'main not registered directly any more' (-not ((OpenEntries) -match '"?ToolsBySanat\.xlam'))
Check 'settings files untouched by install' ((SettingsHash) -eq $settingsBefore)

# ---- 2. uninstall ----
$p = Start-Process "$app\unins000.exe" -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait -PassThru
Start-Sleep 3
Check 'uninstall exit code 0' ($p.ExitCode -eq 0) "code $($p.ExitCode)"
Check 'add-in files removed' (-not (Test-Path "$add\ToolsBySanat.xlam") -and -not (Test-Path "$add\ToolsBySanatLoader.xlam"))
Check 'Excel no longer loads the loader' (-not ((OpenEntries) -match 'ToolsBySanatLoader'))
Check 'removed from Windows installed apps' (-not (Test-Path $unKey))
Check 'app folder removed' (-not (Test-Path $app))
Check 'YOUR SETTINGS SURVIVED the uninstall (kept by default)' ((SettingsHash) -eq $settingsBefore)

# ---- 3. install again (leave this PC installed) ----
$p = Start-Process $exe -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait -PassThru
Check 'reinstall exit code 0' ($p.ExitCode -eq 0) "code $($p.ExitCode)"
Check 'reinstall: loader registered again' ((OpenEntries) -match 'ToolsBySanatLoader\.xlam')

# ---- 4. real chain: loader -> main -> settings ----
Start-Sleep 3
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
try {
    [void]$xl.Workbooks.Open("$add\ToolsBySanatLoader.xlam")
    Check "loader opens main (version $Version)" ($xl.Run("'ToolsBySanat.xlam'!modRibbon.AddinVersion") -eq $Version)
    Check 'settings still readable by the macros' ([bool]$xl.Run("'ToolsBySanat.xlam'!modSettings.SettingValue", 'vendor_email.KEYA'))
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

$res
$fails = @($res | Where-Object { $_ -match ' FAIL' }).Count
if ($fails) { Write-Host "`n$fails check(s) FAILED. Setup log: $env:TEMP\tbs_setup.log" -ForegroundColor Red; exit 1 }
Write-Host "`nAll checks passed." -ForegroundColor Green
