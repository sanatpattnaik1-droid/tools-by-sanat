param(
    [string]$Out = (Join-Path $PSScriptRoot 'ribbon-layout.xlsx')
)

# Writes an Excel sheet describing the CURRENT ribbon (src\ribbon\customUI14.xml): one row per button.
# Edit the sheet, then run apply-layout.ps1 to turn it back into the ribbon XML.

$ErrorActionPreference = 'Stop'
$x = [xml](Get-Content (Join-Path $PSScriptRoot 'src\ribbon\customUI14.xml') -Raw)
$tab = $x.GetElementsByTagName('tab')[0]
$ns = New-Object Xml.XmlNamespaceManager $x.NameTable; $ns.AddNamespace('u', $x.DocumentElement.NamespaceURI)

$rows = New-Object System.Collections.Generic.List[object]
$groups = New-Object System.Collections.Generic.List[string]; $gidList = New-Object System.Collections.Generic.List[string]
foreach ($g in $tab.SelectNodes('u:group', $ns)) {
    $groups.Add($g.label); $gidList.Add($g.id); $col = 0
    foreach ($c in $g.ChildNodes) {
        if ($c.NodeType -ne 'Element') { continue }
        if ($c.LocalName -eq 'box') {
            $col++; $pos = 0
            foreach ($b in $c.ChildNodes) {
                if ($b.NodeType -ne 'Element') { continue }
                $pos++
                $rows.Add([pscustomobject]@{ Group = $g.label; Col = $col; Pos = $pos; Label = $(if ($b.LocalName -eq 'labelControl') { '(version number)' } else { $b.label }); Full = $b.screentip; Size = $(if ($b.LocalName -eq 'labelControl') { 'Label' } else { 'Normal' }); Id = $b.id })
            }
        }
        else { $col++; $rows.Add([pscustomobject]@{ Group = $g.label; Col = $col; Pos = 1; Label = $c.label; Full = $c.screentip; Size = 'Large'; Id = $c.id }) }
    }
}

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Layout'
    $wg = $wb.Worksheets.Item(2); $wg.Name = 'Groups'
    $wh = $wb.Worksheets.Item(3); $wh.Name = 'How to use'

    # ---- Groups sheet ----
    $wg.Range('A1').Value2 = 'Group name'; $wg.Range('B1').Value2 = 'Order on the ribbon (left to right)'
    $wg.Range('C1').Value2 = 'Internal id (do not change)'
    for ($i = 0; $i -lt $groups.Count; $i++) { $wg.Cells($i + 2, 1).Value2 = $groups[$i]; $wg.Cells($i + 2, 2).Value2 = [double]($i + 1); $wg.Cells($i + 2, 3).Value2 = $gidList[$i] }
    $wg.Range('A1:C1').Font.Bold = $true; $wg.Range('A1:C1').Interior.Color = 0x6B3A1F; $wg.Range('A1:C1').Font.Color = 0xFFFFFF
    $wg.Columns('A').ColumnWidth = 28; $wg.Columns('B').ColumnWidth = 36; $wg.Columns('C').ColumnWidth = 28
    $wg.Range('C2:C20').Font.Color = 0x808080
    $wg.Range('E1').Value2 = 'You can rename a group in column A (also change that name in the Layout sheet). Keep column C as it is. To add a group, type a new name in the next free row and give it an order number.'
    $wg.Range('E1').Font.Italic = $true

    # ---- Layout sheet ----
    $hdr = 'Group', 'Column', 'Position', 'Button text on the ribbon', 'Full name (shown as tooltip)', 'Size', 'Button id (do not change)'
    for ($c = 0; $c -lt $hdr.Count; $c++) { $ws.Cells(1, $c + 1).Value2 = $hdr[$c] }
    $n = $rows.Count
    for ($i = 0; $i -lt $n; $i++) {
        $r = $rows[$i]; $row = $i + 2
        $ws.Cells($row, 1).Value2 = [string]$r.Group
        $ws.Cells($row, 2).Value2 = [double]$r.Col
        $ws.Cells($row, 3).Value2 = [double]$r.Pos
        $ws.Cells($row, 4).Value2 = [string]$r.Label
        $ws.Cells($row, 5).Value2 = [string]$r.Full
        $ws.Cells($row, 6).Value2 = [string]$r.Size
        $ws.Cells($row, 7).Value2 = [string]$r.Id
    }
    $h = $ws.Range('A1:G1'); $h.Font.Bold = $true; $h.Interior.Color = 0x6B3A1F; $h.Font.Color = 0xFFFFFF; $h.WrapText = $true
    $w = 22, 9, 10, 34, 40, 10, 28
    for ($c = 0; $c -lt 7; $c++) { $ws.Cells(1, $c + 1).EntireColumn.ColumnWidth = $w[$c] }
    $ws.Range("G2:G$($n + 1)").Font.Color = 0x808080
    $ws.Range("E2:E$($n + 1)").Font.Color = 0x595959
    $ws.Range("B2:C$($n + 1)").HorizontalAlignment = -4108
    $ws.Activate(); $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.FreezePanes = $true
    [void]$ws.Range("A1:G$($n + 1)").AutoFilter()

    # dropdowns
    $v = $ws.Range("A2:A$($n + 30)").Validation; $v.Delete(); $v.Add(3, 1, 1, '=Groups!$A$2:$A$14')
    $v = $ws.Range("B2:B$($n + 30)").Validation; $v.Delete(); $v.Add(3, 1, 1, '1,2,3,4,5,6,7,8')
    $v = $ws.Range("C2:C$($n + 30)").Validation; $v.Delete(); $v.Add(3, 1, 1, '1,2,3')
    $v = $ws.Range("F2:F$($n + 30)").Validation; $v.Delete(); $v.Add(3, 1, 1, 'Normal,Large,Menu,Label')

    # a different light colour per group (follows the group name in column A, also after edits)
    $cols = 0xF3E4D6, 0xDDEFDC, 0xDFF0FB, 0xEBDDF5, 0xDDE7F7, 0xDDF1EF, 0xE8E8E8, 0xF5DDE8
    $rng = $ws.Range("A2:G$($n + 30)")
    for ($g = 0; $g -lt [Math]::Min(8, $groups.Count); $g++) {
        $fc = $rng.FormatConditions.Add(2, 0, "=`$A2=Groups!`$A`$$($g + 2)")
        $fc.Interior.Color = $cols[$g]
    }

    # ---- How to use ----
    $lines = @(
        'HOW TO ARRANGE THE RIBBON',
        '',
        '1. Go to the "Layout" sheet. Each row is one button.',
        '2. To MOVE a button to another group, change its Group (pick from the dropdown).',
        '3. COLUMN is the stack from the left inside the group (1, 2, 3 ...). POSITION is top to bottom inside that stack (1, 2, 3).',
        '   A stack holds at most 3 normal buttons.',
        '4. SIZE: Normal = small button with the text beside it (stacked). Large = big icon with the text under it.',
        '   Large buttons stand alone: give each one its own Column, with Position 1.',
        '   Menu = the whole group becomes ONE small drop-down button that lists its tools (set every row of that group to Menu).',
        '5. You can change "Button text" to rename a button. Keep it short (about 22 letters). "Full name" is the tooltip.',
        '6. Do NOT change the Button id. It links the button to its macro and icon.',
        '7. The "Groups" sheet sets the group names and their left-to-right order. Add a new group by typing a name in the next free row.',
        '8. To hide a button from the ribbon, change its Group to  (hidden)  and add that name on the Groups sheet with order 99.',
        '9. Save the file and send it back (or tell me where you saved it). I turn it into the ribbon, check it in Excel, and ask before publishing.',
        '',
        'The first row "(version number)" in the About group shows the version. Leave it where it is.'
    )
    for ($i = 0; $i -lt $lines.Count; $i++) { $wh.Cells($i + 1, 1).Value2 = $lines[$i] }
    $wh.Range('A1').Font.Bold = $true; $wh.Range('A1').Font.Size = 14; $wh.Columns('A').ColumnWidth = 130
    $ws.Activate()
    if (Test-Path $Out) { Remove-Item $Out -Force }
    $wb.SaveAs($Out, 51); $wb.Close($false)
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
"Wrote $Out ($($rows.Count) rows, $($groups.Count) groups)"
