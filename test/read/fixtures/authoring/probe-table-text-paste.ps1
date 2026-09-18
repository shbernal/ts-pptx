param(
  # Where the pasted decks, their PNGs and the read-back land. Not a fixture.
  [string]$OutDir = (Join-Path $PSScriptRoot '..\..\..\..\.tmp\table-text-paste')
)
# Probe; produces no committed fixture. What PowerPoint does to a table cell's text, which takes its
# size, italic and colour from the source master's p:otherStyle, when its slide is pasted into a
# deck whose p:otherStyle differs.
#
# Source: ../table-text-inheritance.pptx, whose p:otherStyle lvl1 is 14pt, italic, 0070C0 and
# Georgia, and whose three tables (StyledTable, NoGridTable, NoStyleTable) set nothing on their
# runs. Destination: ../default-text-style.pptx, whose p:otherStyle lvl1 is 18pt, tx1 and the minor
# font, not italic.
#   dest.pptx  Slides.Paste, "Use Destination Theme"
#   keep.pptx  ExecuteMso PasteSourceFormatting, "Keep Source Formatting"
# Each pasted table's first header and first body run is read back over COM, the slide is exported
# to PNG, and the saved slide's run properties are what to read for what PowerPoint wrote.
$ErrorActionPreference = 'Stop'
$OutDir = [IO.Path]::GetFullPath($OutDir)
$fixtures = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
New-Item -ItemType Directory -Force $OutDir | Out-Null
foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1                            # the source-formatting paste is a ribbon command
  $src = $pp.Presentations.Open((Join-Path $fixtures 'table-text-inheritance.pptx'), -1, 0, 0)

  foreach ($mode in 'dest', 'keep') {
    $dst = $pp.Presentations.Open((Join-Path $fixtures 'default-text-style.pptx'), -1, 0, -1)
    $src.Slides.Item(1).Copy()
    if ($mode -eq 'dest') {
      [void]$dst.Slides.Paste(2)
    } else {
      $win = $dst.Windows.Item(1)
      $win.Activate()
      $win.ViewType = 7                       # ppViewSlideSorter, so the paste lands after slide 1
      $dst.Slides.Item(1).Select()
      $pp.CommandBars.ExecuteMso('PasteSourceFormatting')
      for ($i = 0; $i -lt 40 -and $dst.Slides.Count -lt 2; $i++) { Start-Sleep -Milliseconds 250 }
    }
    if ($dst.Slides.Count -lt 2) { throw "the $mode paste added no slide" }
    $dst.SaveAs((Join-Path $OutDir "$mode.pptx"))
    $pasted = $dst.Slides.Item(2)
    foreach ($name in 'StyledTable', 'NoGridTable', 'NoStyleTable') {
      $table = $pasted.Shapes.Item($name).Table
      foreach ($row in 1, 2) {
        $font = $table.Cell($row, 1).Shape.TextFrame.TextRange.Runs(1).Font
        Write-Output ('FONT' + "`t" + $mode + "`t" + $name + "`t" + $row + "`t" + $font.Size + "`t" + $font.Italic + "`t" + $font.Bold + "`t" + $font.Name + "`t" + ('{0:X6}' -f $font.Color.RGB))
      }
    }
    $pasted.Export((Join-Path $OutDir "$mode.png"), 'PNG', 1280, 720)
    $dst.Saved = $true
    $dst.Close()
  }
  $src.Close()
  $pp.Quit()
}
finally {
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
