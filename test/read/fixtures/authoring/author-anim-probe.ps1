$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $SCRATCH 'anim-probe.pptx'
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)

  # ---- Slide 1: transition + entrance animations ----
  $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

  # Two named textboxes to animate.
  $tb1 = $slide.Shapes.AddTextbox(1, 100, 100, 400, 80) # msoTextOrientationHorizontal
  $tb1.Name = 'anim-fade'
  $tb1.TextFrame.TextRange.Text = 'Fade in on click'

  $tb2 = $slide.Shapes.AddTextbox(1, 100, 250, 400, 80)
  $tb2.Name = 'anim-fly'
  $tb2.TextFrame.TextRange.Text = 'Fly in after previous'

  # --- Slide transition ---
  # ppEffectFade = 1537 (0x601); advance on click; set duration/speed.
  $slide.SlideShowTransition.EntryEffect = 1537
  $slide.SlideShowTransition.Duration = 1.5
  $slide.SlideShowTransition.AdvanceOnClick = -1  # msoTrue

  # --- Entrance animations via TimeLine.MainSequence.AddEffect ---
  # MsoAnimEffect: msoAnimEffectFade=10, msoAnimEffectFly=2
  # MsoAnimateByLevel: msoAnimateLevelNone=0
  # MsoAnimTriggerType: msoAnimTriggerOnClick=1, msoAnimTriggerAfterPrevious=3
  $seq = $slide.TimeLine.MainSequence
  $e1 = $seq.AddEffect($tb1, 10, 0, 1)
  $e2 = $seq.AddEffect($tb2, 2, 0, 3)

  Write-Host ("Effects on slide: " + $seq.Count)
  Write-Host ("e1 EffectType=" + $e1.EffectType + " Shape=" + $e1.Shape.Name)
  Write-Host ("e2 EffectType=" + $e2.EffectType + " Shape=" + $e2.Shape.Name)

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Host "SAVED: $out"
}
catch {
  Write-Host "ERROR: $($_.Exception.Message)"
  throw
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
