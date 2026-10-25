param(
    [string]$Source = (Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\PERSONAL.XLSB'),
    [string]$OutDir = (Join-Path $PSScriptRoot 'src\modules')
)

# Exports every VBA component from a workbook into $OutDir.
# Works on a temp copy so the original file is never opened or locked.
# Needs Excel option "Trust access to the VBA project object model".

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Source)) { throw "Source not found: $Source" }
New-Item -ItemType Directory -Force $OutDir | Out-Null

$tmp = Join-Path $env:TEMP ('export_' + [guid]::NewGuid().ToString('N') + '.xlsb')
Copy-Item $Source $tmp

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: do not run auto macros
    $wb = $xl.Workbooks.Open($tmp)

    $ext = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }
    $kind = @{ 1 = 'Module'; 2 = 'Class'; 3 = 'UserForm'; 100 = 'Document' }

    $rows = foreach ($c in $wb.VBProject.VBComponents) {
        $lines = $c.CodeModule.CountOfLines
        # Document modules (ThisWorkbook / sheets) are exported only if they hold code.
        if ($c.Type -eq 100 -and $lines -eq 0) { continue }
        $file = Join-Path $OutDir ($c.Name + $ext[[int]$c.Type])
        $c.Export($file)
        [pscustomobject]@{ Name = $c.Name; Type = $kind[[int]$c.Type]; Lines = $lines }
    }
    $rows | Sort-Object Type, Name | Format-Table -AutoSize
    "Exported $(@($rows).Count) components to $OutDir"
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    Remove-Item $tmp -ErrorAction SilentlyContinue
}
