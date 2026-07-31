$ErrorActionPreference = 'Stop'

# Probe (not a fixture recipe): does PowerPoint support turning text wrap OFF in a
# table cell, and what does `a:tcPr/@horzOverflow` actually do?
#
# It authors a base deck with desktop PowerPoint, hand-patches two cells that COM
# cannot reach, reopens the result in PowerPoint, reads the object model back, renders
# both slides to PNG, and re-saves so you can see which attributes survive.
#
# Findings (2026-07-31, PowerPoint desktop, this machine) — recorded in
# `src/gen/slide/objects/table.ts` at the cell `a:tcPr` build:
#   1. `TextFrame.WordWrap` / `TextFrame2.WordWrap` are READ-ONLY on a table cell:
#      every assignment raises "The specified value is out of range" while the same
#      assignment on a text box succeeds. The getter always reports msoTrue (-1).
#   2. `<a:bodyPr wrap="none"/>` inside a cell's `a:txBody` is INERT: the cell wraps
#      exactly like its unpatched neighbour, the OM still reports WordWrap=msoTrue,
#      and PowerPoint strips the attribute on the next save (`<a:bodyPr/>` comes back).
#   3. `<a:tcPr horzOverflow="overflow"/>` is LIVE and preserved across a save. It is
#      not a wrap switch (ECMA-376 §20.1.10.68): with a glyph wider than the column,
#      the default `clip` cuts the glyph at the cell edge and `overflow` draws it whole.
#
# Outputs land in <repo>/.tmp/ (gitignored); nothing here is committed as a fixture.

$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$TMP  = Join-Path $REPO '.tmp'
if (-not (Test-Path $TMP)) { New-Item -ItemType Directory -Path $TMP | Out-Null }

$base    = Join-Path $TMP 'table-cell-wrap-base.pptx'
$patched = Join-Path $TMP 'table-cell-wrap-patched.pptx'
$resaved = Join-Path $TMP 'table-cell-wrap-resaved.pptx'
$pngBase = Join-Path $TMP 'table-cell-wrap-base.png'
$pngPtch = Join-Path $TMP 'table-cell-wrap-patched.png'
foreach ($f in @($base, $patched, $resaved, $pngBase, $pngPtch)) { if (Test-Path $f) { Remove-Item $f -Force } }

# "No Style, No Grid" — keeps the table style from contributing anything of its own.
$noStyleNoGrid = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
$long = 'Wrapping probe with a deliberately long single line of text'

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1

  # ---------------------------------------------------------------- STEP 1: author
  $pres = $pp.Presentations.Add(1)
  $slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank

  # WrapTable: wide columns, long text. A1 gets patched to wrap="none"; A2 is the control.
  $wrapShape = $slide.Shapes.AddTable(1, 2, 60, 60, 480, 120)
  $wrapShape.Name = 'WrapTable'
  $wrapShape.Table.ApplyStyle($noStyleNoGrid, $true)
  foreach ($c in 1..2) { $wrapShape.Table.Cell(1, $c).Shape.TextFrame.TextRange.Text = $long }

  # GlyphTable: columns narrower than one 60pt glyph. A1 gets horzOverflow="overflow";
  # A2 keeps the schema default (clip).
  $glyphShape = $slide.Shapes.AddTable(1, 2, 60, 260, 120, 100)
  $glyphShape.Name = 'GlyphTable'
  $glyphShape.Table.ApplyStyle($noStyleNoGrid, $true)
  $glyphShape.Table.Columns(1).Width = 40
  $glyphShape.Table.Columns(2).Width = 40
  foreach ($c in 1..2) {
    $tr = $glyphShape.Table.Cell(1, $c).Shape.TextFrame.TextRange
    $tr.Text = 'M'
    $tr.Font.Size = 60
  }

  # Confirm the read-only-ness of WordWrap on a cell, against a text box as the control.
  $cell = $wrapShape.Table.Cell(1, 1).Shape
  foreach ($v in @(0, -1)) {
    try { $cell.TextFrame.WordWrap = $v; Write-Output ("cell TextFrame.WordWrap={0}: ACCEPTED" -f $v) }
    catch { Write-Output ("cell TextFrame.WordWrap={0}: REJECTED - {1}" -f $v, $_.Exception.Message) }
    try { $cell.TextFrame2.WordWrap = $v; Write-Output ("cell TextFrame2.WordWrap={0}: ACCEPTED" -f $v) }
    catch { Write-Output ("cell TextFrame2.WordWrap={0}: REJECTED - {1}" -f $v, $_.Exception.Message) }
  }
  $tb = $slide.Shapes.AddTextbox(1, 600, 60, 300, 60)
  $tb.TextFrame.TextRange.Text = 'text box control'
  $tb.TextFrame.WordWrap = 0
  Write-Output ("text box TextFrame.WordWrap=0: ACCEPTED (now {0})" -f $tb.TextFrame.WordWrap)

  $pres.SaveAs($base)
  $pres.Saved = $true
  $pres.Slides(1).Export($pngBase, 'PNG', 960, 540)
  $pres.Close()
  $pres = $null

  # ------------------------------------------------- STEP 2: patch what COM cannot
  Copy-Item $base $patched
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::Open($patched, 'Update')
  try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()

    # First cell of each table, in document order: WrapTable's A1 then GlyphTable's A1.
    $i1 = $xml.IndexOf('<a:tc>')
    $i2 = $xml.IndexOf('<a:tc>', $xml.IndexOf('</a:tbl>', $i1))
    if ($i1 -lt 0 -or $i2 -lt 0) { throw 'could not locate both first cells' }
    $head = $xml.Substring(0, $i2)
    $tail = $xml.Substring($i2)
    # WrapTable A1 (in $head): wrap="none" on the cell's bodyPr.
    $head = ([regex]'<a:bodyPr/>').Replace($head, '<a:bodyPr wrap="none"/>', 1, $i1)
    # GlyphTable A1 (start of $tail): horzOverflow on the cell's tcPr.
    $tail = ([regex]'<a:tcPr/>').Replace($tail, '<a:tcPr horzOverflow="overflow"/>', 1)
    $xml = $head + $tail
    if ($xml -notmatch 'wrap="none"' -or $xml -notmatch 'horzOverflow') { throw 'patch matched nothing' }

    $entry.Delete()
    $new = $zip.CreateEntry('ppt/slides/slide1.xml')
    $writer = New-Object System.IO.StreamWriter($new.Open())
    $writer.Write($xml)
    $writer.Close()
  }
  finally { $zip.Dispose() }
  Write-Output "patched: $patched"

  # ------------------------------------------ STEP 3: what does PowerPoint make of it
  $pres = $pp.Presentations.Open($patched, 0, 0, 0)   # throws on a repair prompt
  Write-Output ("reopened OK, shapes={0}" -f $pres.Slides(1).Shapes.Count)
  foreach ($c in 1..2) {
    $tf = $pres.Slides(1).Shapes('WrapTable').Table.Cell(1, $c).Shape.TextFrame
    Write-Output ("WrapTable cell 1,{0} after patch: TextFrame.WordWrap={1}" -f $c, $tf.WordWrap)
  }
  $pres.Slides(1).Export($pngPtch, 'PNG', 960, 540)
  $pres.SaveAs($resaved)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null
  $pp.Quit()

  # ------------------------------------------- STEP 4: which attributes survived?
  $zip = [System.IO.Compression.ZipFile]::OpenRead($resaved)
  try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $out = $reader.ReadToEnd()
    $reader.Close()
    # Scope each check to the cell that was patched. A whole-document match would be a
    # false positive: the text box control legitimately keeps its own wrap="none".
    $cells = [regex]::Matches($out, '(?s)<a:tc>.*?</a:tc>')
    Write-Output ("after PowerPoint re-save: cell wrap=none kept?    {0}" -f ($cells[0].Value -match 'wrap="none"'))
    Write-Output ("after PowerPoint re-save: cell horzOverflow kept? {0}" -f ($cells[2].Value -match 'horzOverflow'))
    Write-Output ("after PowerPoint re-save: text box wrap=none kept? {0} (control)" -f ($out -match '<a:bodyPr[^>]*wrap="none"'))
  }
  finally { $zip.Dispose() }

  Write-Output "compare renders: $pngBase vs $pngPtch"
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
