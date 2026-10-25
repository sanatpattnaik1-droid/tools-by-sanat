param(
    [string]$Version
)

# Builds build\ToolsBySanat.xlam from src\.
#   1. creates a new workbook in a hidden Excel and imports every module + UserForm1
#   2. pastes src\addin\ThisWorkbook.bas into the ThisWorkbook module
#   3. saves it as .xlam
#   4. injects src\ribbon\customUI14.xml into the package
#   5. writes the SHA-256 into version.json
# Needs Excel option "Trust access to the VBA project object model" (build PC only).

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root      = $PSScriptRoot
$buildDir  = Join-Path $root 'build'
$out       = Join-Path $buildDir 'ToolsBySanat.xlam'
$verFile   = Join-Path $root 'version.json'
$ribbonXml = Join-Path $root 'src\ribbon\customUI14.xml'

$ver = if (Test-Path $verFile) { Get-Content $verFile -Raw | ConvertFrom-Json } else { [pscustomobject]@{ version = '0.1.0'; sha256 = ''; url = '' } }
if ($Version) { $ver.version = $Version }

New-Item -ItemType Directory -Force $buildDir | Out-Null
if (Test-Path $out) {
    try { Remove-Item $out -Force } catch { throw "Cannot replace $out. Close Excel (the add-in may be loaded) and retry." }
}

# Stage files whose text changes at build time (version stamp) outside src\.
$stage = Join-Path $env:TEMP ('tbs_build_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $stage | Out-Null
$ribbonMod = Join-Path $stage 'modRibbon.bas'
(Get-Content (Join-Path $root 'src\addin\modRibbon.bas') -Raw) -replace '(?m)(ADDIN_VERSION As String = ")[^"]*(")', "`${1}$($ver.version)`${2}" |
    Set-Content $ribbonMod -Encoding Default

$thisWbCode = (Get-Content (Join-Path $root 'src\addin\ThisWorkbook.bas') -Raw) -replace '(?m)^Option Explicit\r?\n', ''

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $vbp = $wb.VBProject

    $imports = @(Get-ChildItem (Join-Path $root 'src\modules') -Include *.bas, *.frm -Recurse | Sort-Object Name) +
               @(Get-Item $ribbonMod)
    foreach ($f in $imports) { [void]$vbp.VBComponents.Import($f.FullName) }

    $doc = $vbp.VBComponents.Item($wb.CodeName).CodeModule
    $doc.AddFromString($thisWbCode)

    $wb.Title = 'Tools By Sanat'
    $wb.SaveAs($out, 55)   # xlOpenXMLAddIn
    "Imported $($imports.Count) components."
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    Remove-Item $stage -Recurse -ErrorAction SilentlyContinue
}

# Excel can take a moment to let go of the saved file
for ($try = 0; $try -lt 40; $try++) {
    try { [IO.File]::Open($out, 'Open', 'ReadWrite', 'None').Close(); break } catch { Start-Sleep -Milliseconds 500 }
}

# --- inject the ribbon part ---
$utf8 = New-Object System.Text.UTF8Encoding($false)
$zip = [IO.Compression.ZipFile]::Open($out, 'Update')
try {
    $e = $zip.CreateEntry('customUI/customUI14.xml')
    $w = New-Object IO.StreamWriter($e.Open(), $utf8)
    $w.Write((Get-Content $ribbonXml -Raw)); $w.Close()

    # icons: customUI/images/<buttonId>.png, wired up through customUI/_rels/customUI14.xml.rels
    $iconRels = ''
    foreach ($png in Get-ChildItem (Join-Path $root 'src\icons') -Filter *.png) {
        $id = $png.BaseName
        $ie = $zip.CreateEntry("customUI/images/$($png.Name)")
        $s = $ie.Open(); $bytes = [IO.File]::ReadAllBytes($png.FullName); $s.Write($bytes, 0, $bytes.Length); $s.Close()
        $iconRels += "<Relationship Id=`"$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"images/$($png.Name)`"/>"
    }
    $re = $zip.CreateEntry('customUI/_rels/customUI14.xml.rels')
    $w = New-Object IO.StreamWriter($re.Open(), $utf8)
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $iconRels + '</Relationships>')
    $w.Close()

    # make sure .png parts have a content type
    $ctEntry = $zip.GetEntry('[Content_Types].xml')
    $r = New-Object IO.StreamReader($ctEntry.Open()); $ct = $r.ReadToEnd(); $r.Close()
    if ($ct -notmatch 'Extension="png"') {
        $ctEntry.Delete()
        $ct = $ct -replace '(<Types[^>]*>)', '$1<Default Extension="png" ContentType="image/png"/>'
        $ce = $zip.CreateEntry('[Content_Types].xml')
        $w = New-Object IO.StreamWriter($ce.Open(), $utf8); $w.Write($ct); $w.Close()
    }

    $relEntry = $zip.GetEntry('_rels/.rels')
    $r = New-Object IO.StreamReader($relEntry.Open()); $rels = $r.ReadToEnd(); $r.Close()
    $relEntry.Delete()
    $new = '<Relationship Id="rIdToolsBySanatUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
    $rels = $rels -replace '</Relationships>', ($new + '</Relationships>')
    $e2 = $zip.CreateEntry('_rels/.rels')
    $w = New-Object IO.StreamWriter($e2.Open(), $utf8)
    $w.Write($rels); $w.Close()
}
finally { $zip.Dispose() }

# --- version.json ---
$ver.sha256 = (Get-FileHash $out -Algorithm SHA256).Hash.ToLower()
$cfg = Get-Content (Join-Path $root 'updater.config.json') -Raw | ConvertFrom-Json
$ver.url = if ($cfg.repo) { "https://github.com/$($cfg.repo)/releases/download/v$($ver.version)/ToolsBySanat.xlam" } else { '' }
$ver | ConvertTo-Json | Set-Content $verFile -Encoding UTF8

"Built $out"
"Version $($ver.version)  SHA-256 $($ver.sha256)"
