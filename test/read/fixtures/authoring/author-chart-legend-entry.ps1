param(
  # Where to write the deck. Defaults to the fixture this recipe produces, beside `authoring/`.
  [string]$Out = (Join-Path $PSScriptRoot '..\chart-legend-entry.pptx'),
  # Where the PNG exports land. Not part of the fixture.
  [string]$ProbeDir = (Join-Path $PSScriptRoot '..\..\..\..\.tmp\chart-legend-entry')
)
# chart-legend-entry.pptx: what `c:legendEntry/c:idx` counts when PowerPoint deletes one entry from
# a combo chart's legend, the series' `c:idx` or the entry's position in the legend.
#
# Both charts start as three clustered-column series, A, B and C (c:idx 0, 1, 2), and then have
# series A moved into a second chart group, so the legend no longer lists the series in c:idx order:
#   slide 1  combo-line-chart     A becomes a line
#   slide 2  combo-scatter-chart  A becomes an XY scatter
# The slide is exported to PNG before and after the deletion, so which label went is visible, and
# the first legend entry is the one deleted.
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
$ProbeDir = [IO.Path]::GetFullPath($ProbeDir)
if (Test-Path $Out) { Remove-Item $Out -Force }
New-Item -ItemType Directory -Force $ProbeDir | Out-Null

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$preexisting = @(Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

# Run a COM call until it returns something: right after ChartData.Activate() the workbook can
# still be null while Excel starts.
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

# Fill the chart's workbook with a category column and the three series, point the chart at it,
# close the workbook, and wait for the Excel this run started to exit before the next chart.
function Set-ThreeSeries($chart) {
  $chart.ChartData.Activate()
  $wb = Wait-Com { $chart.ChartData.Workbook } 'the chart workbook'
  $ws = Wait-Com { $wb.Worksheets.Item(1) } 'the chart worksheet'
  [void]$ws.Cells.Clear()
  $rows = @(
    @($null, 'A', 'B', 'C'),
    @('Q1', 1, 4, 7),
    @('Q2', 2, 5, 8),
    @('Q3', 3, 6, 9)
  )
  for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $rows[$r].Count; $c++) {
      $value = $rows[$r][$c]
      if ($null -eq $value) { continue }
      $cell = $ws.Range([string][char](65 + $c) + ($r + 1))
      if ($value -is [string]) { $cell.Value2 = [string]$value } else { $cell.Value2 = [double]$value }
    }
  }
  $chart.SetSourceData("'" + $ws.Name + "'!`$A`$1:`$D`$4")
  $wb.Close()
  for ($attempt = 1; $attempt -le 80; $attempt++) {
    if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $preexisting -notcontains $_.Id })) { break }
    Start-Sleep -Milliseconds 250
  }
}

$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(1)            # with a window: chart data editing wants one

  $cases = @(
    @{ Name = 'combo-line-chart';    Type = 4 },       # xlLine
    @{ Name = 'combo-scatter-chart'; Type = -4169 }    # xlXYScatter
  )
  $index = 0
  foreach ($case in $cases) {
    $index++
    $slide = $pres.Slides.Add($index, 12)     # ppLayoutBlank
    $shape = $slide.Shapes.AddChart2(-1, 51, 60, 60, 840, 420)   # xlColumnClustered
    $shape.Name = $case.Name
    $chart = $shape.Chart
    Set-ThreeSeries $chart
    $chart.HasLegend = $true
    $chart.SeriesCollection(1).ChartType = $case.Type
    $slide.Export((Join-Path $ProbeDir ("slide{0}-before.png" -f $index)), 'PNG', 1280, 720)
    Write-Output ("LEGEND {0} before: {1} entries" -f $case.Name, $chart.Legend.LegendEntries().Count)
    $chart.Legend.LegendEntries(1).Delete()
    Write-Output ("LEGEND {0} after: {1} entries" -f $case.Name, $chart.Legend.LegendEntries().Count)
    $slide.Export((Join-Path $ProbeDir ("slide{0}-after.png" -f $index)), 'PNG', 1280, 720)
  }

  $pres.SaveAs($Out)
  $pres.Saved = $true
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
  Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
