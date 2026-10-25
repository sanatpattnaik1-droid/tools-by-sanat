param(
    [Parameter(Mandatory)][ValidateSet('Install', 'Uninstall')][string]$Action
)

# Registers / unregisters the "Tools By Sanat" loader add-in in Excel for the current user.
# Called by Setup.exe (install and uninstall). Exit codes: 0 ok, 2 Excel not available, 1 other error.

$ErrorActionPreference = 'Stop'
$main   = 'ToolsBySanat.xlam'
$loader = 'ToolsBySanatLoader.xlam'
$dir    = Join-Path $env:APPDATA 'Microsoft\AddIns'

try { $xl = New-Object -ComObject Excel.Application } catch { exit 2 }
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    [void]$xl.Workbooks.Add()   # AddIns needs a workbook context

    if ($Action -eq 'Install') {
        # earlier versions registered the main add-in directly; the loader opens it now
        foreach ($a in @($xl.AddIns | Where-Object { $_.Name -eq $main })) { $a.Installed = $false }
        $l = $xl.AddIns | Where-Object { $_.Name -eq $loader } | Select-Object -First 1
        if (-not $l) { $l = $xl.AddIns.Add((Join-Path $dir $loader), $false) }
        $l.Installed = $true
    }
    else {
        foreach ($n in $loader, $main) {
            foreach ($a in @($xl.AddIns | Where-Object { $_.Name -eq $n })) { $a.Installed = $false }
        }
        # drop the entries from Excel's add-in list too, so no "cannot find" entry stays behind
        $key = "HKCU:\Software\Microsoft\Office\$($xl.Version)\Excel\Add-in Manager"
        if (Test-Path $key) {
            foreach ($n in $loader, $main) { Remove-ItemProperty -Path $key -Name (Join-Path $dir $n) -ErrorAction SilentlyContinue }
        }
    }
    exit 0
}
catch { exit 1 }
finally {
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
