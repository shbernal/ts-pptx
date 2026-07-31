$ErrorActionPreference = 'Stop'

# Produces ../table-cell-horzoverflow.pptx — the read-side oracle for
# `a:tcPr/@horzOverflow` (TableCell.horzOverflow).
#
# Unlike every other recipe here, this one hand-writes the attribute before letting
# PowerPoint own the result, because PowerPoint has **no COM surface for horzOverflow**
# and no UI control for it either. The sequence is:
#   author a plain table with PowerPoint -> inject the attribute into slide1.xml ->
#   REOPEN in PowerPoint and SaveAs the fixture path.
# The committed bytes are therefore PowerPoint's own serialization of the construct, and
# the round-trip is itself the evidence that PowerPoint accepts and preserves it (contrast
# `probe-table-cell-wrap.ps1`, where the same treatment of `<a:bodyPr wrap="none"/>` in a
# cell is silently dropped on save). The probe also renders the visual difference: with a
# glyph wider than the column, `clip` cuts it at the cell edge, `overflow` draws it whole.
#
# Only `overflow` is injected. An earlier revision also injected an explicit
# `horzOverflow="clip"` into a second cell and PowerPoint **dropped it** on the re-save,
# leaving `<a:tcPr/>`: `clip` is the schema default, so PowerPoint normalizes it away.
# There is consequently no such thing as a PowerPoint-authored explicit `clip` to pin.

$FIX = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$out = Join-Path $FIX 'table-cell-horzoverflow.pptx'
$tmp = Join-Path $FIX 'table-cell-horzoverflow.base.pptx'
foreach ($f in @($out, $tmp)) { if (Test-Path $f) { Remove-Item $f -Force } }

# "No Style, No Grid" — the table style contributes nothing of its own.
$noStyleNoGrid = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1

  # --- author: one 1x2 table, columns narrower than a single 60pt glyph ---------
  $pres = $pp.Presentations.Add(1)
  $slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank
  $shape = $slide.Shapes.AddTable(1, 2, 60, 200, 120, 100)
  $shape.Name = 'HorzOverflowTable'
  $shape.Table.ApplyStyle($noStyleNoGrid, $true)
  foreach ($c in 1..2) {
    $shape.Table.Columns($c).Width = 40
    $tr = $shape.Table.Cell(1, $c).Shape.TextFrame.TextRange
    $tr.Text = 'M'
    $tr.Font.Size = 60
  }
  $pres.SaveAs($tmp)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # --- inject: cell 1 overflow, cell 2 left bare as the negative control --------
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::Open($tmp, 'Update')
  try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()

    if (([regex]::Matches($xml, '<a:tcPr/>')).Count -ne 2) { throw 'expected exactly two bare <a:tcPr/> cells' }
    $xml = ([regex]'<a:tcPr/>').Replace($xml, '<a:tcPr horzOverflow="overflow"/>', 1)

    $entry.Delete()
    $new = $zip.CreateEntry('ppt/slides/slide1.xml')
    $writer = New-Object System.IO.StreamWriter($new.Open())
    $writer.Write($xml)
    $writer.Close()
  }
  finally { $zip.Dispose() }

  # --- hand it back to PowerPoint: it must open clean and re-serialize the attrs -
  $pres = $pp.Presentations.Open($tmp, 0, 0, 0)   # throws on a repair prompt
  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null
  $pp.Quit()
  Remove-Item $tmp -Force

  # --- verify what PowerPoint wrote --------------------------------------------
  $zip = [System.IO.Compression.ZipFile]::OpenRead($out)
  try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $saved = $reader.ReadToEnd()
    $reader.Close()
    foreach ($m in [regex]::Matches($saved, '<a:tcPr[^>]*>')) { Write-Output ("  " + $m.Value) }
  }
  finally { $zip.Dispose() }

  Write-Output ("SAVED: {0}" -f $out)
  Write-Output ("SHA256: {0}" -f (Get-FileHash $out -Algorithm SHA256).Hash.ToLower())
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
