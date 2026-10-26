param(
    [string]$Sheet = (Join-Path $PSScriptRoot 'ribbon-layout.xlsx'),
    [string]$OutXml = (Join-Path $PSScriptRoot 'src\ribbon\customUI14.xml'),
    [switch]$AutoFix            # repair column/position conflicts: every Large button gets its own column, ties keep sheet order, stacks hold 3
)

# Turns the edited ribbon-layout.xlsx (see make-layout-sheet.ps1) into the ribbon XML.
# Button ids, macro links and icons come from the current XML, so they cannot be broken by the sheet.
# Stops with a clear message if the sheet breaks a rule (more than 3 buttons in a stack, unknown button, ...).

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'src\ribbon\customUI14.xml'
$cur = [xml](Get-Content $src -Raw)
$ns = New-Object Xml.XmlNamespaceManager $cur.NameTable; $ns.AddNamespace('u', $cur.DocumentElement.NamespaceURI)

# master data from the current ribbon: id -> attributes the sheet may not change
$master = @{}
foreach ($b in $cur.SelectNodes('//u:button', $ns)) { $master[$b.id] = @{ tag = $b.tag; onAction = $b.onAction; kind = 'button' } }
$master['lblVersion'] = @{ kind = 'label' }
$groupId = @{}; foreach ($g in $cur.SelectNodes('//u:group', $ns)) { $groupId[$g.label] = $g.id }

function Esc($s) { [Security.SecurityElement]::Escape([string]$s) }
function Cell($ws, $r, $c) { $v = $ws.Cells($r, $c).Value2; if ($null -eq $v) { '' } else { "$v".Trim() } }

$xl = New-Object -ComObject Excel.Application
$rows = New-Object System.Collections.Generic.List[object]; $groups = @{}; $gids = @{}
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($Sheet, 0, $true)
    $wg = $wb.Worksheets.Item('Groups'); $ws = $wb.Worksheets.Item('Layout')
    for ($r = 2; Cell $wg $r 1; $r++) {
        $o = Cell $wg $r 2; if ($o -notmatch '^\d+$') { throw "Groups sheet row ${r}: the order of '$(Cell $wg $r 1)' must be a number." }
        $groups[(Cell $wg $r 1)] = [int]$o
        $gids[(Cell $wg $r 1)] = Cell $wg $r 3          # internal id: stays the same when the group is renamed
    }
    for ($r = 2; Cell $ws $r 7; $r++) {
        $rows.Add([pscustomobject]@{ Row = $r; Group = Cell $ws $r 1; Col = Cell $ws $r 2; Pos = Cell $ws $r 3; Label = Cell $ws $r 4; Full = Cell $ws $r 5; Size = Cell $ws $r 6; Id = Cell $ws $r 7 })
    }
    $wb.Close($false)
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

# ---- optional repair of Column / Position conflicts ----
if ($AutoFix) {
    foreach ($gname in ($rows | ForEach-Object Group | Sort-Object -Unique)) {
        $inG = @($rows | Where-Object { $_.Group -eq $gname -and $_.Col -match '^\d+$' -and $_.Pos -match '^\d+$' } | Sort-Object { [int]$_.Col }, { [int]$_.Pos }, { $_.Row })
        if (@($inG | Where-Object { $_.Size -notin 'Normal', 'Label' }).Count -eq 0) {
            # only small buttons in this group: fill the stacks of 3 in order, whatever Column they had (fewest columns = least width)
            $k = 0; foreach ($r in $inG) { $r.Col = "$([math]::Floor($k / 3) + 1)"; $r.Pos = "$(($k % 3) + 1)"; $k++ }
            continue
        }
        $newCol = 0; $count = 0; $lastCol = $null; $lastLarge = $true
        foreach ($r in $inG) {
            if ($r.Size -eq 'Large') { $newCol++; $r.Col = "$newCol"; $r.Pos = '1'; $count = 0; $lastLarge = $true }
            else {
                if ($lastLarge -or $r.Col -ne $lastCol -or $count -ge 3) { $newCol++; $count = 0 }
                $count++; $lastCol = $r.Col; $lastLarge = $false
                $r.Col = "$newCol"; $r.Pos = "$count"
            }
            if ($r.Size -eq 'Large') { $lastCol = $null }
        }
    }
}

# ---- validate ----
$errs = New-Object System.Collections.Generic.List[string]
$seen = @{}
foreach ($r in $rows) {
    if (-not $master.ContainsKey($r.Id)) { $errs.Add("Row $($r.Row): unknown button id '$($r.Id)' (do not edit that column).") ; continue }
    if ($seen.ContainsKey($r.Id)) { $errs.Add("Row $($r.Row): button '$($r.Id)' appears twice.") }; $seen[$r.Id] = $true
    if (-not $groups.ContainsKey($r.Group)) { $errs.Add("Row $($r.Row) ('$($r.Label)'): group '$($r.Group)' is not on the Groups sheet.") }
    if ($r.Col -notmatch '^\d+$' -or $r.Pos -notmatch '^\d+$') { $errs.Add("Row $($r.Row) ('$($r.Label)'): Column and Position must be whole numbers (1, 2, 3 ...).") }
    if ($r.Size -notin 'Normal', 'Large', 'Label', 'Menu') { $errs.Add("Row $($r.Row) ('$($r.Label)'): Size must be Normal, Large, Menu or Label.") }
    if (-not $r.Label -and $r.Size -ne 'Label') { $errs.Add("Row $($r.Row): the button text is empty.") }
}
foreach ($id in $master.Keys) { if (-not $seen.ContainsKey($id)) { $errs.Add("Button '$id' is missing from the sheet. Keep every row (move it to the hidden group instead of deleting it).") } }
if ($errs.Count) { $errs | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }; throw "The sheet has $($errs.Count) problem(s). Fix them and run again." }

$hidden = { param($name) $name -match '^\(?hidden\)?$' -or $groups[$name] -ge 99 }
$visible = $rows | Where-Object { -not (& $hidden $_.Group) }

# ---- build the XML ----
$x = New-Object System.Collections.Generic.List[string]
$x.Add('<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">'); $x.Add('  <ribbon>'); $x.Add('    <tabs>')
$x.Add('      <tab id="tabToolsBySanat" label="Tools by Sanat v2" insertAfterMso="TabDeveloper">')
$used = @{}
foreach ($gname in ($visible | ForEach-Object Group | Sort-Object -Unique { $groups[$_] })) {
    $gid = if ($gids[$gname]) { $gids[$gname] } elseif ($groupId.ContainsKey($gname)) { $groupId[$gname] } else { 'grp_' + ($gname -replace '[^A-Za-z0-9]', '') }
    if ($used.ContainsKey($gid)) { throw "Two groups would get the id '$gid'. Rename one of them." }; $used[$gid] = $true
    $x.Add(''); $x.Add("        <group id=`"$gid`" label=`"$(Esc $gname)`">")
    $inGroup = $visible | Where-Object Group -eq $gname
    if (@($inGroup | Where-Object Size -eq 'Menu').Count) {
        # a "Menu" group is ONE compact drop-down button that lists its buttons (good for rarely used tools)
        if (@($inGroup | Where-Object Size -ne 'Menu').Count) { throw "Group '$gname': if one button is a Menu, all of them must be Menu." }
        $mid = 'mnu_' + ($gid -replace '^grp_?', '')
        $x.Add("          <menu id=`"$mid`" label=`"$(Esc $gname)`" size=`"large`" image=`"$mid`" itemSize=`"normal`">")
        foreach ($i in ($inGroup | Sort-Object { [int]$_.Col }, { [int]$_.Pos }, { $_.Row })) {
            $m = $master[$i.Id]; $tip = if ($i.Full) { " screentip=`"$(Esc $i.Full)`"" } else { '' }
            $act = if ($m.tag) { " tag=`"$(Esc $m.tag)`" onAction=`"RunMacro`"" } elseif ($m.onAction) { " onAction=`"$($m.onAction)`"" } else { '' }
            $x.Add("            <button id=`"$($i.Id)`" label=`"$(Esc $i.Label)`"$tip image=`"$($i.Id)`"$act/>")
        }
        $x.Add('          </menu>'); $x.Add('        </group>')
        continue
    }
    $stacks = $inGroup | Group-Object { [int]$_.Col } | Sort-Object { [int]$_.Name }
    $boxN = 0
    foreach ($s in $stacks) {
        $items = @($s.Group | Sort-Object { [int]$_.Pos })
        $large = @($items | Where-Object Size -eq 'Large')
        if ($large.Count -and $large.Count -ne $items.Count) { throw "Group '$gname', column $($s.Name): Large buttons cannot share a column with Normal ones. Give each Large button its own column." }
        if ($large.Count) {
            foreach ($i in $items) { $x.Add($(
                $m = $master[$i.Id]; $tip = if ($i.Full) { " screentip=`"$(Esc $i.Full)`"" } else { '' }
                $act = if ($m.tag) { " tag=`"$(Esc $m.tag)`" onAction=`"RunMacro`"" } elseif ($m.onAction) { " onAction=`"$($m.onAction)`"" } else { '' }
                "          <button id=`"$($i.Id)`" label=`"$(Esc $i.Label)`"$tip size=`"large`" image=`"$($i.Id)`"$act/>")) }
        }
        else {
            if ($items.Count -gt 3) { throw "Group '$gname', column $($s.Name) has $($items.Count) buttons. The most a stack holds is 3. Use another Column." }
            if (@($items.Pos | Sort-Object -Unique).Count -ne $items.Count) { throw "Group '$gname', column $($s.Name): two buttons have the same Position." }
            $boxN++; $x.Add("          <box id=`"box_${gid}_$boxN`" boxStyle=`"vertical`">")
            foreach ($i in $items) {
                $m = $master[$i.Id]
                if ($m.kind -eq 'label') { $x.Add("            <labelControl id=`"lblVersion`" getLabel=`"GetVersionLabel`"/>"); continue }
                $tip = if ($i.Full) { " screentip=`"$(Esc $i.Full)`"" } else { '' }
                $act = if ($m.tag) { " tag=`"$(Esc $m.tag)`" onAction=`"RunMacro`"" } elseif ($m.onAction) { " onAction=`"$($m.onAction)`"" } else { '' }
                $x.Add("            <button id=`"$($i.Id)`" label=`"$(Esc $i.Label)`"$tip size=`"normal`" image=`"$($i.Id)`"$act/>")
            }
            $x.Add('          </box>')
        }
    }
    $x.Add('        </group>')
}
$x.Add(''); $x.Add('      </tab>'); $x.Add('    </tabs>'); $x.Add('  </ribbon>'); $x.Add('</customUI>')
$text = ($x -join "`r`n") + "`r`n"
$new = [xml]$text     # must be well-formed
[IO.File]::WriteAllText($OutXml, $text, (New-Object Text.UTF8Encoding($false)))
$btn = @($new.GetElementsByTagName('button')).Count
"Wrote $OutXml : $(@($new.GetElementsByTagName('group')).Count) groups, $btn buttons visible, $(@($rows | Where-Object { & $hidden $_.Group }).Count) hidden."
