$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $FIX 'slide-animation-presets.pptx'
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
    $tb = $slide.Shapes.AddTextbox(1, 120, $y, 720, 44)
    $tb.Name = $name
    $tb.TextFrame.TextRange.Text = $text
    return $tb
  }

  # One labeled text box per preset, one effect each, all on click (each its own click group)
  # so every preset template is isolated. Covers the downstream-recommended set across all 3 classes.
  $b1 = Add-Box $slide 'entr-fadeIn'  20  'Entrance: fade in'
  $b2 = Add-Box $slide 'entr-flyIn'   80  'Entrance: fly in'
  $b3 = Add-Box $slide 'entr-appear' 140  'Entrance: appear'
  $b4 = Add-Box $slide 'entr-wipe'   200  'Entrance: wipe'
  $b5 = Add-Box $slide 'emph-grow'   260  'Emphasis: grow/shrink'
  $b6 = Add-Box $slide 'emph-spin'   320  'Emphasis: spin'
  $b7 = Add-Box $slide 'exit-fadeOut' 380 'Exit: fade out'
  $b8 = Add-Box $slide 'exit-flyOut'  440 'Exit: fly out'

  # MsoAnimEffect: Appear=1, Fly=2, Fade=10, Wipe=22, GrowShrink=59 (emph presetID 6),
  #   Spin=61 (emph presetID 8). Probed empirically; see .tmp/animeffect-order.json.
  # MsoAnimTriggerType: OnClick=1.  Exit effects: AddEffect then Effect.Exit = msoTrue(-1).
  $seq = $slide.TimeLine.MainSequence

  $null = $seq.AddEffect($b1, 10, 0, 1)        # entrance fade in
  $null = $seq.AddEffect($b2, 2,  0, 1)        # entrance fly in
  $null = $seq.AddEffect($b3, 1,  0, 1)        # entrance appear
  $null = $seq.AddEffect($b4, 22, 0, 1)        # entrance wipe
  $null = $seq.AddEffect($b5, 59, 0, 1)        # emphasis grow/shrink
  $null = $seq.AddEffect($b6, 61, 0, 1)        # emphasis spin
  $e7 = $seq.AddEffect($b7, 10, 0, 1); $e7.Exit = -1   # exit fade out
  $e8 = $seq.AddEffect($b8, 2,  0, 1); $e8.Exit = -1   # exit fly out

  for ($i = 1; $i -le $seq.Count; $i++) {
    $e = $seq.Item($i)
    Write-Host ("#{0}: EffectType={1} Exit={2} Shape={3}" -f $i, $e.EffectType, $e.Exit, $e.Shape.Name)
  }

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
