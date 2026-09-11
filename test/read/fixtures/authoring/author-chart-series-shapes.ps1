param(
  # Where to write the deck. Defaults to the fixture this recipe produces, beside `authoring/`.
  [string]$Out = (Join-Path $PSScriptRoot '..\chart-series-shapes.pptx')
)
# chart-series-shapes.pptx: the chart constructs the read model has to see and did not.
#   slide 1  scatter-chart         c:scatterChart  -> c:xVal + c:yVal
#   slide 2  bubble-chart          c:bubbleChart   -> c:xVal + c:yVal + c:bubbleSize
#   slide 3  multilevel-bar-chart  c:barChart      -> c:cat/c:multiLvlStrRef (region > quarter)
#   slide 4  pie-chart             c:pieChart      -> data labels showing percent and category name
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
if (Test-Path $Out) { Remove-Item $Out -Force }

# Snapshot pre-existing PIDs so the reap at the end only kills the servers this run spawned:
# PowerPoint, and the Excel that each chart's embedded workbook opens.
$preexisting = @(Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

# Run a COM call until it returns something, for the calls that answer before Excel is ready:
# right after `ChartData.Activate()` the workbook can still be null while Excel starts.
function Wait-Com([scriptblock]$call, [string]$what) {
  for ($attempt = 1; $attempt -le 80; $attempt++) {
    try {
      $value = & $call
      if ($null -ne $value) { return $value }
    } catch {}
    Start-Sleep -Milliseconds 250
  }
  throw "timed out waiting for $what"
}

# Fill a chart's embedded workbook from row arrays (row 1 first, column A first; $null leaves a
# cell empty), point the chart at the range, and close the workbook.
function Set-ChartData($chart, [object[]]$rows, [string]$range) {
  $chart.ChartData.Activate()
  $wb = Wait-Com { $chart.ChartData.Workbook } 'the chart workbook'
  $ws = Wait-Com { $wb.Worksheets.Item(1) } 'the chart worksheet'
  [void]$ws.Cells.Clear()
  # Text and numbers go through separate assignments. One `.Value2 = $value` site that first
  # receives a string fails on the next number with "Unable to cast object of type 'System.Int32'
  # to type 'System.String'": PowerShell's COM binder appears to keep the conversion from the
  # first call. Numbers are written as doubles, so they stay numbers in the chart's cache.
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $row = $rows[$r]
    for ($c = 0; $c -lt $row.Count; $c++) {
      $value = $row[$c]
      if ($null -eq $value) { continue }
      $cell = $ws.Range([string][char](65 + $c) + ($r + 1))
      if ($value -is [string]) { $cell.Value2 = [string]$value }
      else { $cell.Value2 = [double]$value }
    }
  }
  $chart.SetSourceData("'" + $ws.Name + "'!" + $range)
  $wb.Close()
}

$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)            # msoTrue: with window
  $pres.PageSetup.SlideWidth  = 960
  $pres.PageSetup.SlideHeight = 540

  # --- slide 1: scatter (xlXYScatter = -4169), X in column A, Y in column B ---
  $slide = $pres.Slides.Add(1, 12)            # ppLayoutBlank
  $shape = $slide.Shapes.AddChart2(-1, -4169, 60, 60, 840, 420)
  $shape.Name = 'scatter-chart'
  Set-ChartData $shape.Chart @(
    @('X', 'Y'),
    @(1, 2.5),
    @(2, 4),
    @(3, 3.5),
    @(4, 6)
  ) '$A$1:$B$5'

  # --- slide 2: bubble (xlBubble = 15), X in A, Y in B, bubble size in C ---
  # The range starts below the header row. A bubble chart does not take its first row as
  # headers the way the scatter above does: given A1:C5 it plotted the header as a point, which
  # made the X range a text cache starting with "X" and gave Y and size a leading 0.
  $slide = $pres.Slides.Add(2, 12)
  $shape = $slide.Shapes.AddChart2(-1, 15, 60, 60, 840, 420)
  $shape.Name = 'bubble-chart'
  Set-ChartData $shape.Chart @(
    @('X', 'Y', 'Size'),
    @(1, 2, 5),
    @(2, 4, 10),
    @(3, 3, 7),
    @(4, 6, 12)
  ) '$A$2:$C$5'

  # --- slide 3: clustered column (xlColumnClustered = 51) over two category levels ---
  # Two text columns left of the values, with the header cells above them empty, are what Excel
  # reads as a region > quarter category hierarchy.
  $slide = $pres.Slides.Add(3, 12)
  $shape = $slide.Shapes.AddChart2(-1, 51, 60, 60, 840, 420)
  $shape.Name = 'multilevel-bar-chart'
  Set-ChartData $shape.Chart @(
    @($null, $null, 'Revenue'),
    @('North', 'Q1', 10),
    @($null, 'Q2', 14),
    @('South', 'Q1', 8),
    @($null, 'Q2', 12)
  ) '$A$1:$C$5'

  # --- slide 4: pie (xlPie = 5) with data labels showing percent and category name ---
  $slide = $pres.Slides.Add(4, 12)
  $shape = $slide.Shapes.AddChart2(-1, 5, 60, 60, 840, 420)
  $shape.Name = 'pie-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'Share'),
    @('North', 35),
    @('South', 25),
    @('East', 22),
    @('West', 18)
  ) '$A$1:$B$5'
  $series = $shape.Chart.SeriesCollection(1)
  $series.HasDataLabels = $true
  $labels = $series.DataLabels()
  $labels.ShowValue = $false
  $labels.ShowPercentage = $true
  $labels.ShowCategoryName = $true

  $pres.SaveAs($Out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output ('SAVED: ' + $Out)
}
catch {
  Write-Output ('FAIL: ' + $_.Exception.Message)
  if ($pres -ne $null) { try { $pres.Close() } catch {} }
  if ($pp   -ne $null) { try { $pp.Quit()   } catch {} }
  throw
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp   -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  # Quit() can leave the automation servers lingering; reap only PIDs this run created.
  Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
