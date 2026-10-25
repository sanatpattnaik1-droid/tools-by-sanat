param(
    [ValidateSet('A', 'B', 'C')][string]$Style = 'B',
    [string]$OutDir = (Join-Path $PSScriptRoot 'src\icons'),
    [string[]]$Only
)

# Generates one 64x64 transparent PNG per ribbon button. The file name is the button id from
# src\ribbon\customUI14.xml. Pictograms come from "Segoe Fluent Icons" (ships with Windows 11);
# customer logos come from src\icons\source\<name>.png (not in git).
#   Style A: bold gradient pictogram, no tile
#   Style B: vivid gradient squircle with a white pictogram (default)
#   Style C: white card with a coloured pictogram and a soft accent blob

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function C($hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
$blue = '#2F80ED'; $indigo = '#5B5BD6'; $violet = '#8E5CF7'; $teal = '#12B5A6'; $green = '#22B15B'
$amber = '#F5A623'; $orange = '#F2711C'; $red = '#E5484D'; $pink = '#E5469A'

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
$fontName = if ((New-Object System.Drawing.Text.InstalledFontCollection).Families.Name -contains 'Segoe Fluent Icons') { 'Segoe Fluent Icons' } else { 'Segoe MDL2 Assets' }
$family = New-Object System.Drawing.FontFamily($fontName)

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

# The pictogram as a vector path, scaled to fit a box of $size px and centred on ($cx, $cy).
function Get-GlyphPath([int]$cp, [float]$size, [float]$cx, [float]$cy) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $fmt = New-Object System.Drawing.StringFormat
    $path.AddString([string][char]$cp, $family, 0, 100, (New-Object System.Drawing.PointF 0, 0), $fmt)
    $b = $path.GetBounds()
    $s = $size / [math]::Max($b.Width, $b.Height)
    $m = New-Object System.Drawing.Drawing2D.Matrix
    $m.Translate(-($b.X + $b.Width / 2), -($b.Y + $b.Height / 2))
    $m.Scale($s, $s, [System.Drawing.Drawing2D.MatrixOrder]::Append)
    $m.Translate($cx, $cy, [System.Drawing.Drawing2D.MatrixOrder]::Append)
    $path.Transform($m)
    $path
}

function Blend($c, $to, [double]$t) {
    [System.Drawing.Color]::FromArgb(255, [int]($c.R + ($to.R - $c.R) * $t), [int]($c.G + ($to.G - $c.G) * $t), [int]($c.B + ($to.B - $c.B) * $t))
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$S = 64
$white = [System.Drawing.Color]::White; $black = [System.Drawing.Color]::Black

$list = if ($Only) { $icons | Where-Object { $_[0] -in $Only } } else { $icons }
foreach ($i in $list) {
    $id, $cp, $hex = $i
    $col = C $hex
    $light = Blend $col $white 0.30; $dark = Blend $col $black 0.18
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.InterpolationMode = 'HighQualityBicubic'; $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([System.Drawing.Color]::Transparent)
    $tile = New-RoundedPath 2 2 60 60 17

    if ($brand.ContainsKey($id)) {
        $logo = Get-Logo $brand[$id]
        if ($Style -ne 'A' -and $brand[$id] -ne 'blinkit') {
            # logo sits on a card
            $fill = if ($Style -eq 'B') { $white } else { Blend $col $white 0.88 }
            $g.FillPath((New-Object System.Drawing.SolidBrush $fill), $tile)
            $pen = New-Object System.Drawing.Pen ((C '#D5DCE6')), 1; $g.DrawPath($pen, $tile); $pen.Dispose()
        }
        if ($brand[$id] -eq 'blinkit') {
            $full = New-RoundedPath 3 3 58 58 15
            $g.SetClip($full); $g.DrawImage($logo, 3, 3, 58, 58); $g.ResetClip()
        } else {
            $box = if ($Style -eq 'A') { 56 } else { 40 }
            $scale = [math]::Min($box / $logo.Width, $box / $logo.Height)
            $lw = $logo.Width * $scale; $lh = $logo.Height * $scale
            $ox = if ($Style -eq 'A') { 2 + (56 - $lw) / 2 } else { 6 + (40 - $lw) / 2 }
            $oy = if ($Style -eq 'A') { 2 + (46 - $lh) / 2 } else { 6 + (40 - $lh) / 2 }
            $g.DrawImage($logo, $ox, $oy, $lw, $lh)
        }
        $logo.Dispose()
        # action badge, bottom right (bigger for the stacked 16 px buttons; Blinkit keeps it small so its wordmark stays readable)
        $bs = if ($brand[$id] -eq 'blinkit') { 25 } else { 32 }
        $bx = 63 - $bs
        $bb = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.RectangleF $bx, $bx, $bs, $bs)), $light, $dark, 90
        $g.FillEllipse($bb, $bx, $bx, $bs, $bs)
        $wp = New-Object System.Drawing.Pen $white, 2; $g.DrawEllipse($wp, ($bx + 1), ($bx + 1), ($bs - 2), ($bs - 2)); $wp.Dispose()
        $gp = Get-GlyphPath $cp ($bs * 0.56) ($bx + $bs / 2) ($bx + $bs / 2)
        $g.FillPath((New-Object System.Drawing.SolidBrush $white), $gp); $gp.Dispose(); $bb.Dispose()
    }
    elseif ($Style -eq 'A') {
        $gp = Get-GlyphPath $cp 50 32 32
        $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.RectangleF 0, 6, 64, 52)), $light, $dark, 90
        $pen = New-Object System.Drawing.Pen $br, 2.4; $pen.LineJoin = 'Round'
        $g.FillPath($br, $gp); $g.DrawPath($pen, $gp)
        $gp.Dispose(); $br.Dispose(); $pen.Dispose()
    }
    elseif ($Style -eq 'B') {
        $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.RectangleF 0, 2, 64, 60)), $light, $dark, 60
        $g.FillPath($br, $tile)
        # soft gloss on the top half
        $g.SetClip($tile)
        $gl = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.RectangleF 0, 2, 64, 30)), ([System.Drawing.Color]::FromArgb(70, 255, 255, 255)), ([System.Drawing.Color]::FromArgb(0, 255, 255, 255)), 90
        $g.FillRectangle($gl, 0, 2, 64, 30); $g.ResetClip()
        $gp = Get-GlyphPath $cp 34 32 32.5
        $g.FillPath((New-Object System.Drawing.SolidBrush $white), $gp)
        $pen = New-Object System.Drawing.Pen $white, 1.3; $pen.LineJoin = 'Round'; $g.DrawPath($pen, $gp)
        $gp.Dispose(); $br.Dispose(); $gl.Dispose(); $pen.Dispose()
    }
    else {  # C
        $sh = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(28, 0, 0, 0))
        $g.FillPath($sh, (New-RoundedPath 2 4 60 59 17))
        $g.FillPath((New-Object System.Drawing.SolidBrush $white), $tile)
        $g.SetClip($tile)
        $g.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(70, $col.R, $col.G, $col.B))), 26, -8, 44, 44)
        $g.ResetClip()
        $gp = Get-GlyphPath $cp 32 29 35
        $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.RectangleF 0, 16, 64, 38)), $light, $dark, 90
        $pen = New-Object System.Drawing.Pen $br, 1.4; $pen.LineJoin = 'Round'
        $g.FillPath($br, $gp); $g.DrawPath($pen, $gp)
        $gp.Dispose(); $br.Dispose(); $pen.Dispose(); $sh.Dispose()
    }
    $bmp.Save((Join-Path $OutDir "$id.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
}
"Wrote $(@($list).Count) icons (style $Style) to $OutDir (font: $fontName)"
