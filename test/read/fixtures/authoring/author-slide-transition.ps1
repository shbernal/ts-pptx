$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $FIX 'slide-transition.pptx'
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null

# PpTransitionSpeed: Slow=1, Medium=2, Fast=3.  msoTrue=-1, msoFalse=0.
function New-LabeledSlide($pres, $text) {
  $s = $pres.Slides.Add($pres.Slides.Count + 1, 12) # ppLayoutBlank
  $tb = $s.Shapes.AddTextbox(1, 60, 230, 840, 80)
  $tb.Name = 'label'
  $tb.TextFrame.TextRange.Text = $text
  return $s
}

try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540

  # Slide 1 — fade, FAST speed bucket only (bare p:transition, no p14:dur), advance on click.
  $s = New-LabeledSlide $pres 'fade / fast / click'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 3849            # p:fade
  $t.Speed = 3                     # fast bucket
  $t.AdvanceOnClick = -1

  # Slide 2 — push DOWN, exact duration (forces mc:AlternateContent + p14:dur), advance on click.
  $s = New-LabeledSlide $pres 'push down / 1.25s exact / click'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 3852            # p:push dir="d"
  $t.Duration = 1.25               # off-bucket -> p14:dur="1250"
  $t.AdvanceOnClick = -1

  # Slide 3 — wipe UP, MEDIUM speed bucket, advance on click.
  $s = New-LabeledSlide $pres 'wipe up / medium / click'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 2818            # p:wipe dir="u"
  $t.Speed = 2                     # medium bucket
  $t.AdvanceOnClick = -1

  # Slide 4 — cut, FAST bucket, advance on click.
  $s = New-LabeledSlide $pres 'cut / fast / click'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 257             # p:cut
  $t.Speed = 3                     # fast bucket
  $t.AdvanceOnClick = -1

  # Slide 5 — dissolve, exact duration (mc:AlternateContent + p14:dur), SLOW.
  $s = New-LabeledSlide $pres 'dissolve / 2.0s exact / click'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 1537            # p:dissolve
  $t.Duration = 2.0                # -> p14:dur="2000", spd="slow"
  $t.AdvanceOnClick = -1

  # Slide 6 — fade, TIMED auto-advance (advTm), advance-on-click OFF.
  $s = New-LabeledSlide $pres 'fade / timed 3s / no click'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 3849            # p:fade
  $t.Speed = 2                     # medium bucket
  $t.AdvanceOnClick = 0
  $t.AdvanceOnTime = -1
  $t.AdvanceTime = 3.0             # -> advTm="3000"

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Host "SAVED: $out ($($pres -ne $null))"
}
catch {
  Write-Host "ERROR: $($_.Exception.Message)"
  throw
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
