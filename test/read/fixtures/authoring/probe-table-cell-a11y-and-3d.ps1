$ErrorActionPreference = 'Stop'

# Probe (not a fixture recipe): does PowerPoint PRESERVE the three `a:tc` / `a:tcPr`
# constructs that have no COM surface and no UI control?
#
#   1. `a:tc/@id`                        — a cell's unique identifier
#   2. `a:tcPr/a:headers/a:header/@val`  — which header cells govern a data cell
#   3. `a:tcPr/a:cell3D`                 — a cell's 3-D bevel
#
# All three are ECMA-376 constructs the write path can emit, but emitting one is only
# worth doing if PowerPoint keeps it: `@id` in particular is a transitional-profile
# addition, and a construct PowerPoint strips on the first save is an inert feature
# rather than a shipped one. The question is answered the only way it can be — hand-patch
# the part, hand it back to PowerPoint, and read out what it re-serializes.
#
# The sequence mirrors `author-table-cell-horzoverflow.ps1`:
#   author a plain 3x3 table with PowerPoint -> inject the constructs into slide1.xml ->
#   REOPEN in PowerPoint (which throws here if it wants to repair the file) -> SaveAs ->
#   print every `a:tc` opening tag and every `a:tcPr` block PowerPoint wrote back.
#
# The table follows the ECMA-376 §21.1.3.4 worked example: header cells A/B across the
# top and C/D down the side, four data cells, each data cell naming its row and column
# header. `cell3D` goes on one data cell so its survival is answered independently.
#
# Outputs land in <repo>/.tmp/ (gitignored); nothing here is committed as a fixture.
#
# FINDINGS (2026-07-31, PowerPoint desktop, this machine):
#
#   1. `a:tc/@id` is STRIPPED. All nine cells came back as bare `<a:tc>`.
#   2. `a:tcPr/a:headers` is STRIPPED. Not one `<a:header>` survived.
#   3. `a:tcPr/a:cell3D` SURVIVES VERBATIM — element, `@prstMaterial`, `a:bevel`'s
#      `@w`/`@h`/`@prst` and the whole `a:lightRig`, byte for byte.
#
# (1) and (2) are what settles the header-association feature: PowerPoint opens the file
# without complaint and then quietly discards both halves, so an emitter for them would
# ship a feature that is inert the moment anyone saves the deck. It is NOT implemented on
# the write side; the read accessors stay, because a deck from another producer may carry
# what PowerPoint will not write. `hasHeader` (`a:tblPr/@firstRow`) remains the only
# header marker PowerPoint keeps, and it is what its own accessibility checker reads.
#
# The comparison is controlled rather than circumstantial: `a:cell3D` and `a:headers` were
# injected into the SAME `a:tcPr` on the same cell, and PowerPoint kept one and dropped the
# other. So this is not "the patch did not land" — it is a deliberate normalization.

$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$TMP = Join-Path $REPO '.tmp'
if (-not (Test-Path $TMP)) { New-Item -ItemType Directory -Path $TMP | Out-Null }

$base = Join-Path $TMP 'table-cell-a11y-base.pptx'
$resaved = Join-Path $TMP 'table-cell-a11y-resaved.pptx'
foreach ($f in @($base, $resaved)) { if (Test-Path $f) { Remove-Item $f -Force } }

# "No Style, No Grid" — the table style contributes nothing of its own.
$noStyleNoGrid = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1

  # --- author: the 3x3 table from the spec's own worked example ------------------
  $pres = $pp.Presentations.Add(1)
  $slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank
  $shape = $slide.Shapes.AddTable(3, 3, 60, 120, 540, 240)
  $shape.Name = 'HeaderAssociationTable'
  $shape.Table.ApplyStyle($noStyleNoGrid, $true)
  $text = @(
    @('', 'A', 'B'),
    @('C', 'x1', 'x2'),
    @('D', 'y1', 'y2')
  )
  foreach ($r in 1..3) {
    foreach ($c in 1..3) {
      $shape.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text = $text[$r - 1][$c - 1]
    }
  }
  $pres.SaveAs($base)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # --- inject ---------------------------------------------------------------------
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::Open($base, 'Update')
  try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()

    $cellCount = ([regex]::Matches($xml, '<a:tc>')).Count
    if ($cellCount -ne 9) { throw ("expected 9 bare <a:tc> cells, found {0}" -f $cellCount) }
    if (([regex]::Matches($xml, '<a:tcPr/>')).Count -ne 9) { throw 'expected 9 bare <a:tcPr/> elements' }

    # `@id` onto the four header cells, in document order: (r1c2)=A (r1c3)=B
    # (r2c1)=C (r3c1)=D. Cell indices below are 0-based over the nine <a:tc> in order.
    $ids = @{ 1 = 'HeaderA'; 2 = 'HeaderB'; 3 = 'HeaderC'; 6 = 'HeaderD' }
    # `a:headers` onto the four data cells, each naming its column then its row header.
    $headers = @{
      4 = @('HeaderA', 'HeaderC')   # x1
      5 = @('HeaderB', 'HeaderC')   # x2
      7 = @('HeaderA', 'HeaderD')   # y1
      8 = @('HeaderB', 'HeaderD')   # y2
    }
    # A 3-D bevel on x1 only, so its survival is answered separately from the a11y pair.
    $cell3D = @{ 4 = '<a:cell3D prstMaterial="metal"><a:bevel w="88900" h="88900" prst="artDeco"/><a:lightRig rig="threePt" dir="t"/></a:cell3D>' }

    $index = -1
    $xml = [regex]::Replace($xml, '<a:tc>', {
        param($m)
        $script:index++
        if ($ids.ContainsKey($script:index)) { return ('<a:tc id="{0}">' -f $ids[$script:index]) }
        return $m.Value
      })

    $index = -1
    $xml = [regex]::Replace($xml, '<a:tcPr/>', {
        param($m)
        $script:index++
        $inner = ''
        if ($cell3D.ContainsKey($script:index)) { $inner += $cell3D[$script:index] }
        if ($headers.ContainsKey($script:index)) {
          $inner += '<a:headers>'
          foreach ($h in $headers[$script:index]) { $inner += ('<a:header val="{0}"/>' -f $h) }
          $inner += '</a:headers>'
        }
        if ($inner -eq '') { return $m.Value }
        # Schema order: cell3D, then the fill group, then headers. No fill here.
        return ('<a:tcPr>{0}</a:tcPr>' -f $inner)
      })

    $entry.Delete()
    $new = $zip.CreateEntry('ppt/slides/slide1.xml')
    $writer = New-Object System.IO.StreamWriter($new.Open())
    $writer.Write($xml)
    $writer.Close()
  }
  finally { $zip.Dispose() }

  # --- hand it back to PowerPoint --------------------------------------------------
  # Open() throws if PowerPoint wants to repair the package, so reaching the next line
  # is itself the "the injected markup is acceptable" result.
  $pres = $pp.Presentations.Open($base, 0, 0, 0)
  $pres.SaveAs($resaved)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null
  $pp.Quit()

  # --- read out what PowerPoint re-serialized ---------------------------------------
  $zip = [System.IO.Compression.ZipFile]::OpenRead($resaved)
  try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $saved = $reader.ReadToEnd()
    $reader.Close()
  }
  finally { $zip.Dispose() }

  Write-Output 'RESULT ------------------------------------------------------------'
  Write-Output ('  opened without a repair prompt: yes')
  Write-Output ('  a:tc opening tags:')
  foreach ($m in [regex]::Matches($saved, '<a:tc(?: [^>]*)?>')) { Write-Output ('    ' + $m.Value) }
  Write-Output ('  a:tcPr blocks:')
  foreach ($m in [regex]::Matches($saved, '<a:tcPr(?:/>|[^>]*>[\s\S]*?</a:tcPr>)')) { Write-Output ('    ' + $m.Value) }
  Write-Output ('  @id survived        : {0}' -f ([regex]::Matches($saved, '<a:tc id="Header')).Count)
  Write-Output ('  a:headers survived  : {0}' -f ([regex]::Matches($saved, '<a:header ')).Count)
  Write-Output ('  a:cell3D survived   : {0}' -f ([regex]::Matches($saved, '<a:cell3D')).Count)
  Write-Output ('  a:bevel survived    : {0}' -f ([regex]::Matches($saved, '<a:bevel')).Count)
  Write-Output ('  a:lightRig survived : {0}' -f ([regex]::Matches($saved, '<a:lightRig')).Count)
  Write-Output ('SAVED: {0}' -f $resaved)
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
