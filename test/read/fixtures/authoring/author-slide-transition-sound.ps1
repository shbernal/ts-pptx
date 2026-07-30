$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $FIX 'slide-transition-sound.pptx'
$wav = (Join-Path $ASSETS 'ding.wav')
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null

# PpSoundEffectType: ppSoundNone=0, ppSoundStopPrevious=1, ppSoundFile=2. msoTrue=-1.
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

  # Built-in transition sounds (e.g. "Applause") were probed and embed IDENTICALLY
  # to a custom import (same audio rel + audio/x-wav Default + sndAc/stSnd/snd), so
  # the committed fixture stays license-clean of Microsoft's bundled audio and uses
  # only our own generated WAV. See the oracle notes.

  # Slide 1 — fade + EMBEDDED custom WAV (the basic stSnd case).
  $s = New-LabeledSlide $pres 'fade + embedded transition sound'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 3849            # p:fade
  $t.SoundEffect.ImportFromFile($wav)

  # Slide 2 — fade + EMBEDDED custom WAV, looped until next sound (stSnd loop="1").
  $s = New-LabeledSlide $pres 'fade + embedded sound, loop until next'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 3849
  $t.SoundEffect.ImportFromFile($wav)
  $t.LoopSoundUntilNext = -1       # -> stSnd loop="1"

  # Slide 3 — fade + STOP PREVIOUS sound (endSnd form, no rel, no media part).
  $s = New-LabeledSlide $pres 'fade + stop previous sound'
  $t = $s.SlideShowTransition
  $t.EntryEffect = 3849
  $t.SoundEffect.Type = 1          # ppSoundStopPrevious

  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $se = $pres.Slides.Item($i).SlideShowTransition.SoundEffect
    Write-Host ("slide {0}: SoundEffect.Type={1} Name='{2}' Loop={3}" -f $i, $se.Type, $se.Name, $pres.Slides.Item($i).SlideShowTransition.LoopSoundUntilNext)
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
