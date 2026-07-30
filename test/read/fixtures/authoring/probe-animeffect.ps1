$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $SCRATCH 'animeffect-probe.pptx'
$map = Join-Path $SCRATCH 'animeffect-order.json'
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $slide = $pres.Slides.Add(1, 12)
  $seq = $slide.TimeLine.MainSequence

  # One shape + one entrance-trigger effect per candidate MsoAnimEffect id.
  # Shape name encodes the effect id so we can map name -> id after parsing.
  $applied = @()
  foreach ($id in 1..150) {
    $tb = $null
    try {
      $tb = $slide.Shapes.AddTextbox(1, 10, 10, 80, 20)
      $tb.Name = ('eff{0}' -f $id)
      $tb.TextFrame.TextRange.Text = 'x'
      $e = $seq.AddEffect($tb, $id, 0, 2)   # WithPrevious so they don't chain clicks
      $applied += $id
    } catch {
      if ($tb -ne $null) { $tb.Delete() }
    }
  }
  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  $applied | ConvertTo-Json -Compress | Set-Content -Path $map -Encoding UTF8
  Write-Host ("APPLIED " + $applied.Count + " effect ids; saved " + $out)
}
catch { Write-Host "ERROR: $($_.Exception.Message)"; throw }
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
