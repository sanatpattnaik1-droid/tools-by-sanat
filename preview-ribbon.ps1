param(
    [Parameter(Mandatory)][string]$Xml,
    [Parameter(Mandatory)][string]$Png
)

# Previews a ribbon XML without rebuilding the add-in: copies build\ToolsBySanat.xlam, swaps in the XML, opens it in a visible
# Excel and saves a picture of ONLY the Excel window's top 200 px (PrintWindow: nothing else on the screen is captured).

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.IO.Compression.FileSystem
Add-Type @'
using System; using System.Runtime.InteropServices;
public class PW {
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
 public struct RECT { public int L, T, R, B; } }
'@
$tmp = Join-Path $env:TEMP 'ToolsBySanat_preview.xlam'
Copy-Item (Join-Path $PSScriptRoot 'build\ToolsBySanat.xlam') $tmp -Force
$zip = [IO.Compression.ZipFile]::Open($tmp, 'Update')
try {
    $zip.GetEntry('customUI/customUI14.xml').Delete()
    $e = $zip.CreateEntry('customUI/customUI14.xml'); $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false)))
    $w.Write((Get-Content $Xml -Raw)); $w.Close()
} finally { $zip.Dispose() }

for ($i = 0; $i -lt 15 -and (Get-Process excel -ErrorAction SilentlyContinue); $i++) { Start-Sleep 1 }
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $true; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
    [void]$xl.Workbooks.Add(); [void]$xl.Workbooks.Open($tmp)
    $xl.WindowState = -4137; Start-Sleep 3
    $h = [IntPtr]$xl.Hwnd
    $win = [System.Windows.Automation.AutomationElement]::FromHandle($h)
    $tabs = $win.FindAll('Descendants', (New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)), (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, 'Tools by Sanat v2')))))
    try { $tabs[0].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch {}
    Start-Sleep 2
    $r = New-Object PW+RECT; [void][PW]::GetWindowRect($h, [ref]$r)
    $wd = $r.R - $r.L; $ht = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap($wd, $ht); $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][PW]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose()
    $crop = $bmp.Clone((New-Object System.Drawing.Rectangle(0, 0, $wd, [Math]::Min(200, $ht))), $bmp.PixelFormat); $bmp.Dispose()
    $crop.Save($Png); $crop.Dispose()
    "saved $Png ($wd px wide window)"
}
finally { try { foreach ($b in @($xl.Workbooks)) { try { $b.Close($false) } catch {} }; $xl.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); Remove-Item $tmp -ErrorAction SilentlyContinue }
