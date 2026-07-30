$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $FIX 'import-animation-merge.pptx'
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540

  # Slide 1 — one shape "Source" with an entrance Fade-on-click (mirrors slide-animation-basic).
  $s1 = $pres.Slides.Add(1, 12)
  $src = $s1.Shapes.AddTextbox(1, 120, 120, 480, 70)
  $src.Name = 'Source'
  $src.TextFrame.TextRange.Text = 'Source (fade in)'
  $null = $s1.TimeLine.MainSequence.AddEffect($src, 10, 0, 1)   # entrance fade, on click

  # Slide 2 — one shape "HostExisting" with its own entrance Fly-on-click.
  $s2 = $pres.Slides.Add(2, 12)
  $hostShape = $s2.Shapes.AddTextbox(1, 120, 320, 480, 70)
  $hostShape.Name = 'HostExisting'
  $hostShape.TextFrame.TextRange.Text = 'Host existing (fly in)'
  $null = $s2.TimeLine.MainSequence.AddEffect($hostShape, 2, 0, 1)   # entrance fly, on click

  # Copy "Source" (with its animation) from slide 1 and paste onto slide 2.
  # PowerPoint renumbers the pasted shape's spid and merges its build into slide 2's timing.
  $s1.Shapes.Item('Source').Copy()
  Start-Sleep -Milliseconds 400
  $pasted = $s2.Shapes.Paste()
  $pasted.Item(1).Name = 'Source'   # keep the stable name for the oracle

  Write-Host ("slide2 shapes: " + (($s2.Shapes | ForEach-Object { $_.Name }) -join ', '))
  Write-Host ("slide2 mainSeq effects: " + $s2.TimeLine.MainSequence.Count)
  for ($i = 1; $i -le $s2.TimeLine.MainSequence.Count; $i++) {
    $e = $s2.TimeLine.MainSequence.Item($i)
    Write-Host ("  effect {0}: type={1} shape={2}" -f $i, $e.EffectType, $e.Shape.Name)
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
