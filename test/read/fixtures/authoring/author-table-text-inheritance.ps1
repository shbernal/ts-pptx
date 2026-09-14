param(
  # Where to write the deck. Defaults to the fixture this recipe produces, beside `authoring/`.
  [string]$Out = (Join-Path $PSScriptRoot '..\table-text-inheritance.pptx'),
  # Where the read-back and the PNG export land. Not part of the fixture.
  [string]$ProbeDir = (Join-Path $PSScriptRoot '..\..\..\..\.tmp\table-text-inheritance')
)
# table-text-inheritance.pptx: where the text of a table cell that sets nothing of its own takes
# its size, face and colour from.
#
# Every candidate tier is given a value no other tier has, so whichever one wins is readable:
#   p:defaultTextStyle lvl1 (presentation.xml)   28pt, C00000, Courier New
#   p:otherStyle lvl1 (the slide master)         14pt, 0070C0, Georgia, italic
#   the theme                                    dk1 7030A0, minor Latin font Verdana
#   Medium Style 2 - Accent 1 tcTxStyle          no size; minor font; dk1 body, lt1 bold header
# The theme edits are what separate a table's own defaults from the theme: a hard-coded black or
# Aptos would survive them, a dk1 or minor-font reference would not.
#
#   TextBox       a text box, the control: the repo reads it through p:defaultTextStyle
#   StyledTable   2x2, the default table style (Medium Style 2 - Accent 1)
#   NoGridTable   2x2, "No Style, No Grid"
#   NoStyleTable  2x2 whose a:tableStyleId is removed, so no table style applies at all
#
# Neither text style has a COM surface, and neither does removing a table's style, so those
# three edits are injected into the saved package and the deck is handed back to PowerPoint to
# reopen and re-save: the committed bytes are PowerPoint's own. The recipe then reopens the
# fixture read-only, prints each run's resolved Font over COM and exports slide 1 to PNG.
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
$ProbeDir = [IO.Path]::GetFullPath($ProbeDir)
$tmp = [IO.Path]::ChangeExtension($Out, '.base.pptx')
foreach ($f in @($Out, $tmp)) { if (Test-Path $f) { Remove-Item $f -Force } }
New-Item -ItemType Directory -Force $ProbeDir | Out-Null

# The faces the two injected tiers name must really be installed, or PowerPoint substitutes at
# render time while still writing the requested name.
Add-Type -AssemblyName System.Drawing
foreach ($face in 'Courier New', 'Georgia', 'Verdana') {
  $f = New-Object System.Drawing.Font($face, 18); $resolved = $f.Name; $f.Dispose()
  if ($resolved -ne $face) { throw "FONT SUBSTITUTED: $face -> $resolved" }
}

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$noStyleNoGrid = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

# Rewrite one zip entry in place. $edit takes the entry's text and returns the new text.
function Update-ZipEntry([string]$zipPath, [string]$entryName, [scriptblock]$edit) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Update')
  try {
    $entry = $zip.GetEntry($entryName)
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $text = $reader.ReadToEnd()
    $reader.Close()
    $text = & $edit $text
    $entry.Delete()
    $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($entryName).Open())
    $writer.Write($text)
    $writer.Close()
  }
  finally { $zip.Dispose() }
}

# Replace exactly one match of $pattern, or fail naming what was not found.
function Replace-Once([string]$text, [string]$pattern, [string]$replacement, [string]$what) {
  $count = ([regex]::Matches($text, $pattern)).Count
  if ($count -ne 1) { throw "expected exactly one $what, found $count" }
  return ([regex]$pattern).Replace($text, $replacement, 1)
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1

  # --- author ---------------------------------------------------------------------------
  $pres = $pp.Presentations.Add(0)
  $slide = $pres.Slides.Add(1, 12)            # ppLayoutBlank

  $box = $slide.Shapes.AddTextbox(1, 40, 20, 600, 60)
  $box.Name = 'TextBox'
  $box.TextFrame.TextRange.Text = 'Box text'

  $tables = @(
    @{ Name = 'StyledTable';  Top = 100; Style = $null },
    @{ Name = 'NoGridTable';  Top = 250; Style = $noStyleNoGrid },
    @{ Name = 'NoStyleTable'; Top = 400; Style = $null }
  )
  foreach ($t in $tables) {
    $shape = $slide.Shapes.AddTable(2, 2, 40, $t.Top, 600, 100)
    $shape.Name = $t.Name
    if ($t.Style) { $shape.Table.ApplyStyle($t.Style, $true) }
    $shape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'Head A'
    $shape.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = 'Head B'
    $shape.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = 'Body A'
    $shape.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = 'Body B'
  }
  $pres.SaveAs($tmp)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # --- inject the three edits COM cannot make --------------------------------------------
  $plainLvl1 = '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="\+mn-lt"/>'
  Update-ZipEntry $tmp 'ppt/presentation.xml' {
    param($xml)
    Replace-Once $xml "(<p:defaultTextStyle><a:defPPr><a:defRPr lang=`"en-US`"/></a:defPPr><a:lvl1pPr[^>]*>)$plainLvl1" `
      '$1<a:defRPr sz="2800" kern="1200"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="Courier New"/>' `
      'default text style lvl1 defRPr'
  }
  Update-ZipEntry $tmp 'ppt/slideMasters/slideMaster1.xml' {
    param($xml)
    Replace-Once $xml "(<p:otherStyle><a:defPPr><a:defRPr lang=`"en-US`"/></a:defPPr><a:lvl1pPr[^>]*>)$plainLvl1" `
      '$1<a:defRPr sz="1400" i="1" kern="1200"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill><a:latin typeface="Georgia"/>' `
      'master otherStyle lvl1 defRPr'
  }
  Update-ZipEntry $tmp 'ppt/theme/theme1.xml' {
    param($xml)
    $xml = Replace-Once $xml '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' `
      '<a:dk1><a:srgbClr val="7030A0"/></a:dk1>' 'theme dk1'
    Replace-Once $xml '<a:minorFont><a:latin typeface="Aptos" panose="[0-9]+"/>' `
      '<a:minorFont><a:latin typeface="Verdana"/>' 'theme minor Latin font'
  }
  Update-ZipEntry $tmp 'ppt/slides/slide1.xml' {
    param($xml)
    # NoStyleTable is the last graphic frame, so its style id is the last one on the slide.
    $ids = [regex]::Matches($xml, '<a:tableStyleId>[^<]*</a:tableStyleId>')
    if ($ids.Count -ne 3) { throw "expected three a:tableStyleId, found $($ids.Count)" }
    $last = $ids[$ids.Count - 1]
    $xml.Remove($last.Index, $last.Length)
  }

  # --- hand it back to PowerPoint to reopen and re-save ----------------------------------
  $pres = $pp.Presentations.Open($tmp, 0, 0, 0)
  $pres.SaveAs($Out)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null
  Remove-Item $tmp -Force

  # --- read back and render what was saved -----------------------------------------------
  $pres = $pp.Presentations.Open($Out, -1, 0, 0)
  $s = $pres.Slides.Item(1)
  function Out-Font([string]$label, $range) {
    $font = $range.Runs(1).Font
    Write-Output ('FONT' + "`t" + $label + "`t" + $font.Size + "`t" + $font.Name + "`t" + ('{0:X6}' -f $font.Color.RGB) + "`t" + $font.Bold)
  }
  Out-Font 'TextBox' $s.Shapes.Item('TextBox').TextFrame.TextRange
  foreach ($t in $tables) {
    $table = $s.Shapes.Item($t.Name).Table
    foreach ($r in 1..2) { foreach ($c in 1..2) {
      Out-Font ("{0}[{1},{2}]" -f $t.Name, $r, $c) $table.Cell($r, $c).Shape.TextFrame.TextRange
    } }
  }
  $png = Join-Path $ProbeDir 'slide1.png'
  $s.Export($png, 'PNG', 1280, 720)
  Write-Output ('PNG' + "`t" + $png)
  $pres.Close()
  $pres = $null
  $pp.Quit()
  Write-Output ('SAVED: ' + $Out)
}
finally {
  if ($pres -ne $null) { try { $pres.Close() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
