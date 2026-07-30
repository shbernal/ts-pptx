$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'bar-chart-data-labels.pptx')
if (Test-Path $out) { Remove-Item $out -Force }

$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)            # msoTrue: with window
  $pres.PageSetup.SlideWidth  = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)            # ppLayoutBlank

  # xlColumnClustered = 51, default style (-1) -> CT_BarSer (barDir=col)
  $shape = $slide.Shapes.AddChart2(-1, 51, 60, 60, 840, 420)
  $shape.Name = 'bar-chart-727'
  $chart = $shape.Chart

  # --- embedded workbook: one series, four categories ---
  $chart.ChartData.Activate()
  $wb = $chart.ChartData.Workbook
  $ws = $wb.Worksheets.Item(1)

  $ws.Range('A2').Value2 = 'Q1'
  $ws.Range('A3').Value2 = 'Q2'
  $ws.Range('A4').Value2 = 'Q3'
  $ws.Range('A5').Value2 = 'Q4'
  $ws.Range('B1').Value2 = 'Revenue'
  $ws.Range('B2').Value2 = 10
  $ws.Range('B3').Value2 = 25
  $ws.Range('B4').Value2 = 18
  $ws.Range('B5').Value2 = 30
  # drop the two extra sample series so the cache holds exactly one series
  $ws.Range('C1:D5').Clear()

  $src = "'" + $ws.Name + "'!" + '$A$1:$B$5'
  $chart.SetSourceData($src)

  # --- per-point fills (c:dPt) + custom per-point label text (c:dLbl) ---
  $series = $chart.SeriesCollection(1)
  $series.HasDataLabels = $true

  # VBA RGB(r,g,b) = r + g*256 + b*65536
  $rgb = @{
    1 = 255            # FF0000 red
    2 = 5287936        # 00B050 green  (0 + 176*256 + 80*65536)
    3 = 12611584       # 0070C0 blue   (0 + 112*256 + 192*65536)
    4 = 49407          # FFC000 amber  (255 + 192*256 + 0)
  }
  $labels = @{ 1 = 'Low'; 2 = 'Mid'; 3 = 'High'; 4 = 'Peak' }

  for ($i = 1; $i -le 4; $i++) {
    $pt = $series.Points($i)
    $pt.Format.Fill.Solid()
    $pt.Format.Fill.ForeColor.RGB = $rgb[$i]
    $pt.HasDataLabel = $true
    $pt.DataLabel.Text = $labels[$i]
  }

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output ('SAVED: ' + $out)
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
}
