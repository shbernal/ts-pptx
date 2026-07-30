$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $FIX 'slide-animation-rich.pptx'
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540

  $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

  function Add-Box($slide, $name, $y, $text) {
    $tb = $slide.Shapes.AddTextbox(1, 120, $y, 720, 70)
    $tb.Name = $name
    $tb.TextFrame.TextRange.Text = $text
    return $tb
  }

  $b1 = Add-Box $slide 'ent-fade-click'  60  'Entrance: fade (on click)'
  $b2 = Add-Box $slide 'ent-fly-after'   170 'Entrance: fly in (after previous)'
  $b3 = Add-Box $slide 'emph-grow-with'  280 'Emphasis: grow/shrink (with previous)'
  $b4 = Add-Box $slide 'exit-fade-click' 390 'Exit: fade (on click)'

  # MsoAnimEffect: Fade=10, Fly=2. MsoAnimEffect=59 -> emphasis Grow/Shrink (presetClass="emph"
  # presetID=6), probed; the entrance-block value 36 was an entrance, not an emphasis.
  # MsoAnimateByLevel: None=0.  MsoAnimTriggerType: OnClick=1, WithPrevious=2, AfterPrevious=3.
  $seq = $slide.TimeLine.MainSequence

  $e1 = $seq.AddEffect($b1, 10, 0, 1)                 # entrance fade, on click
  $e2 = $seq.AddEffect($b2, 2,  0, 3)                 # entrance fly, after previous
  $e3 = $seq.AddEffect($b3, 59, 0, 2)                 # emphasis grow/shrink, with previous
  $e4 = $seq.AddEffect($b4, 10, 0, 1)                 # fade, on click ...
  $e4.Exit = -1                                       # ... converted to an EXIT effect

  foreach ($pair in @(@('e1',$e1),@('e2',$e2),@('e3',$e3),@('e4',$e4))) {
    $e = $pair[1]
    Write-Host ("{0}: EffectType={1} Exit={2} Shape={3}" -f $pair[0], $e.EffectType, $e.Exit, $e.Shape.Name)
  }
  Write-Host ("MainSequence count = " + $seq.Count)

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
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
