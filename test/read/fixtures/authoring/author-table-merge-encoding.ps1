param(
  # Where to write the deck. Defaults to the fixture this recipe produces, beside `authoring/`.
  [string]$Out = (Join-Path $PSScriptRoot '..\table-merge-encoding.pptx')
)
# table-merge-encoding.pptx: which span attributes PowerPoint writes on the covered cells of a
# merged region, and how it keeps them in step when a row or a column goes through the region.
#   slide 1  merge-2x2        3x3, cells (1,1)-(2,2) merged
#   slide 2  merge-inserted   the same merge, then a row inserted before row 2 and a column
#                             inserted before column 2, both through the region
#   slide 3  merge-deleted    4x4, cells (1,1)-(3,3) merged, then row 2 and column 2 deleted
#                             through the region, which leaves a 2x2 merge
#   slide 4  merge-1x2        3x3, cells (1,1)-(1,2) merged: a horizontal span only
#   slide 5  merge-2x1        3x3, cells (1,1)-(2,1) merged: a vertical span only
# Every cell holds its own "r,c" label before the merge, so the saved text also shows what
# PowerPoint does with the content of the cells a merge covers.
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
if (Test-Path $Out) { Remove-Item $Out -Force }

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

# A table of $size x $size on a new slide, each cell labelled with its 1-based "r,c".
function New-LabelledTable($pres, [int]$index, [string]$name, [int]$size) {
  $slide = $pres.Slides.Add($index, 12)       # ppLayoutBlank
  $shape = $slide.Shapes.AddTable($size, $size, 60, 60, 120 * $size, 40 * $size)
  $shape.Name = $name
  foreach ($r in 1..$size) { foreach ($c in 1..$size) {
    $shape.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text = "$r,$c"
  } }
  return $shape.Table
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(0)

  $t = New-LabelledTable $pres 1 'merge-2x2' 3
  $t.Cell(1, 1).Merge($t.Cell(2, 2))

  $t = New-LabelledTable $pres 2 'merge-inserted' 3
  $t.Cell(1, 1).Merge($t.Cell(2, 2))
  [void]$t.Rows.Add(2)
  [void]$t.Columns.Add(2)

  $t = New-LabelledTable $pres 3 'merge-deleted' 4
  $t.Cell(1, 1).Merge($t.Cell(3, 3))
  $t.Rows.Item(2).Delete()
  $t.Columns.Item(2).Delete()

  $t = New-LabelledTable $pres 4 'merge-1x2' 3
  $t.Cell(1, 1).Merge($t.Cell(1, 2))

  $t = New-LabelledTable $pres 5 'merge-2x1' 3
  $t.Cell(1, 1).Merge($t.Cell(2, 1))

  $pres.SaveAs($Out)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # Reopen what was saved, and print each slide's cells: row, column, span attributes, text.
  $pres = $pp.Presentations.Open($Out, -1, 0, 0)
  foreach ($slide in $pres.Slides) {
    $table = $slide.Shapes.Item(1).Table
    Write-Output ("SLIDE {0} {1}: {2} rows x {3} columns" -f $slide.SlideIndex, $slide.Shapes.Item(1).Name, $table.Rows.Count, $table.Columns.Count)
  }
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
