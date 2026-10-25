param(
    [string]$OutDir = (Join-Path $PSScriptRoot 'src\icons'),
    [string]$Sheet,
    [switch]$Dark
)

# Generates one 64x64 transparent PNG per ribbon button, flat modern style:
#   - plain buttons: a bold coloured pictogram (Segoe Fluent Icons) on a soft tinted disc
#   - customer buttons: the customer's logo (src\icons\source\<name>.png) with a small action badge
# The file name is the button id from src\ribbon\customUI14.xml.
# -Sheet <png> also writes a labelled contact sheet (-Dark = dark background) for review.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$blue = '#3B82F6'; $indigo = '#6366F1'; $violet = '#8B5CF6'; $teal = '#14B8A6'; $green = '#22C55E'
$amber = '#F59E0B'; $orange = '#F97316'; $red = '#EF4444'; $pink = '#EC4899'

# id, Segoe Fluent Icons code point, colour
$icons = @(
    @('btnCompWiseSaleRegister',   0xE8A5, $blue),   @('btnDaburEInvoice',         0xE7C3, $teal),
    @('btnIrnUpload',              0xE898, $indigo), @('btnCollectionOutstanding', 0xE8EF, $green),
    @('btnVanWiseOutstanding',     0xE81C, $amber),  @('btnKeepOnlyHighlighted',   0xE71C, $violet),
    @('btnFillDownBlanks',         0xE74B, $blue),   @('btnPartyReceivingSummary', 0xE716, $teal),
    @('btnEInvoiceValidation',     0xE73E, $green),  @('btnDaburSaleFormatHA',     0xE70F, $indigo),
    @('btnPasswordBreaker',        0xE785, $red),    @('btnWisdomStkreport',       0xE719, $orange),
    @('btnConvertSelectedCells',   0xE8AB, $violet), @('btnCombineHeaderClean',    0xE71D, $blue),
    @('btnMaricoHideColumns',      0xED1A, $pink),
    @('btnFlipkartPO',             0xE70F, $blue),   @('btnBlinkitPO',             0xE70F, $blue),
    @('btnDmsCustomerSale',        0xE77B, $orange), @('btnDmsHsnSummary',         0xE8FD, $amber),
    @('btnDmsStock',               0xE719, $orange), @('btnEInvoiceDmsMax',        0xE8A5, $amber),
    @('btnSrmPoExcel',             0xE8C8, $blue),   @('btnPrintSetup',            0xE749, $blue),
    @('btnRilPoToExcel',           0xE7C3, $blue),   @('btnProcessAmendment',      0xE72C, $blue),
    @('btnInsertBrandAsn',         0xE8EC, $blue),   @('btnRilAppointment',        0xE787, $blue),
    @('btnCheckUpdates',           0xE895, $teal),   @('btnSettings',              0xE713, $violet)
)

$brand = @{
    btnFlipkartPO = 'flipkart'; btnBlinkitPO = 'blinkit'
    btnSrmPoExcel = 'reliance'; btnPrintSetup = 'reliance'; btnRilPoToExcel = 'reliance'
    btnProcessAmendment = 'reliance'; btnInsertBrandAsn = 'reliance'; btnRilAppointment = 'reliance'
}
# x, y, w, h of the part of the source image to use; $true = make the background transparent
$crop = @{
    flipkart = @(612, 175, 226, 226, $true)
    reliance = @(94, 0, 82, 66, $true)
    blinkit  = @(0, 0, 1280, 1280, $false)
}
$srcDir = Join-Path $PSScriptRoot 'src\icons\source'
$iconFont = if ((New-Object System.Drawing.Text.InstalledFontCollection).Families.Name -contains 'Segoe Fluent Icons') { 'Segoe Fluent Icons' } else { 'Segoe MDL2 Assets' }

function Get-Logo([string]$name) {
    $c = $crop[$name]
    $src = New-Object System.Drawing.Bitmap (Join-Path $srcDir "$name.png")
    $bm = $src.Clone((New-Object System.Drawing.Rectangle $c[0], $c[1], $c[2], $c[3]), [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $src.Dispose()
    if ($c[4]) {
        # flood-fill the background from the corners so colours inside the logo are kept
        $bg = $bm.GetPixel(1, 1); $w = $bm.Width; $h = $bm.Height
        $seen = New-Object 'bool[]' ($w * $h)
        $q = New-Object System.Collections.Generic.Queue[int]
        foreach ($p in @(0, ($w - 1), (($h - 1) * $w), (($h - 1) * $w + $w - 1))) { $q.Enqueue($p); $seen[$p] = $true }
        while ($q.Count) {
            $p = $q.Dequeue(); $x = $p % $w; $y = [math]::Floor($p / $w)
            $px = $bm.GetPixel($x, $y)
            if ([math]::Abs($px.R - $bg.R) + [math]::Abs($px.G - $bg.G) + [math]::Abs($px.B - $bg.B) -gt 70) { continue }
            $bm.SetPixel($x, $y, [System.Drawing.Color]::Transparent)
            foreach ($n in @(@(($x - 1), $y), @(($x + 1), $y), @($x, ($y - 1)), @($x, ($y + 1)))) {
                if ($n[0] -ge 0 -and $n[0] -lt $w -and $n[1] -ge 0 -and $n[1] -lt $h) {
                    $k = $n[1] * $w + $n[0]
                    if (-not $seen[$k]) { $seen[$k] = $true; $q.Enqueue($k) }
                }
            }
        }
    }
    $bm
}

function New-RoundedPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $p.AddArc($x, $y, $d, $d, 180, 90); $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90); $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure(); $p
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$S = 64
$glyphFont = New-Object System.Drawing.Font($iconFont, 36, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
$badgeFont = New-Object System.Drawing.Font($iconFont, 14, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
$fmt = New-Object System.Drawing.StringFormat
$fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'

foreach ($i in $icons) {
    $id, $cp, $hex = $i
    $col = [System.Drawing.ColorTranslator]::FromHtml($hex)
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.InterpolationMode = 'HighQualityBicubic'; $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([System.Drawing.Color]::Transparent)

    if ($brand.ContainsKey($id)) {
        $logo = Get-Logo $brand[$id]
        if ($brand[$id] -eq 'blinkit') {
            $full = New-RoundedPath 3 3 ($S - 6) ($S - 6) 14
            $g.SetClip($full); $g.DrawImage($logo, 3, 3, ($S - 6), ($S - 6)); $g.ResetClip()
        } else {
            $scale = [math]::Min(54 / $logo.Width, 52 / $logo.Height)
            $lw = $logo.Width * $scale; $lh = $logo.Height * $scale
            $g.DrawImage($logo, (3 + (54 - $lw) / 2), (3 + (52 - $lh) / 2), $lw, $lh)
        }
        $logo.Dispose()
        # action badge, bottom right
        $badge = New-Object System.Drawing.SolidBrush $col
        $g.FillEllipse($badge, 38, 38, 25, 25)
        $g.DrawString([string][char]$cp, $badgeFont, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 38, 39, 25, 25), $fmt)
        $badge.Dispose()
    } else {
        $tint = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(48, $col))
        $g.FillEllipse($tint, 3, 3, 58, 58)
        $solid = New-Object System.Drawing.SolidBrush $col
        # the Fluent glyphs are thin; draw them with small offsets to make them bolder
        foreach ($o in @(@(0, 0), @(0.7, 0), @(-0.7, 0), @(0, 0.7), @(0, -0.7))) {
            $g.DrawString([string][char]$cp, $glyphFont, $solid, (New-Object System.Drawing.RectangleF $o[0], (2 + $o[1]), $S, $S), $fmt)
        }
        $tint.Dispose(); $solid.Dispose()
    }
    $bmp.Save((Join-Path $OutDir "$id.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
}
"Wrote $($icons.Count) icons to $OutDir (font: $iconFont)"

if ($Sheet) {
    $cols = 9; $cw = 90; $ch = 96; $rows = [math]::Ceiling($icons.Count / $cols)
    $sb = New-Object System.Drawing.Bitmap ($cols * $cw), ($rows * $ch)
    $g = [System.Drawing.Graphics]::FromImage($sb)
    $bgc = if ($Dark) { '#202020' } else { '#F7F7F7' }; $fg = if ($Dark) { [System.Drawing.Brushes]::White } else { [System.Drawing.Brushes]::Black }
    $g.Clear([System.Drawing.ColorTranslator]::FromHtml($bgc))
    $lf = New-Object System.Drawing.Font('Segoe UI', 7)
    for ($n = 0; $n -lt $icons.Count; $n++) {
        $x = ($n % $cols) * $cw; $y = [math]::Floor($n / $cols) * $ch
        $img = [System.Drawing.Image]::FromFile((Join-Path $OutDir "$($icons[$n][0]).png"))
        $g.DrawImage($img, $x + 13, $y + 4, 64, 64); $img.Dispose()
        $g.DrawString(($icons[$n][0] -replace '^btn', ''), $lf, $fg, (New-Object System.Drawing.RectangleF $x, ($y + 70), $cw, 24), $fmt)
    }
    $sb.Save($Sheet, [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $sb.Dispose()
    "Contact sheet: $Sheet"
}
