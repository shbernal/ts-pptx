$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out    = (Join-Path $FIX 'online-video.pptx')
$src    = (Join-Path $FIX 'media\tiny.mp4')
# Neutral public target so the committed fixture's external link carries no username path.
$target = 'C:\Users\Public\online-video-sample.mp4'
Copy-Item $src $target -Force
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth  = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)           # ppLayoutBlank

  # External-link video: LinkToFile=msoTrue (-1), SaveWithDocument=msoFalse (0).
  # Linked-not-embedded -> a:videoFile r:link + external video rel (no media part);
  # PowerPoint auto-generates the poster frame, so no explicit cover is needed.
  $vid = $slide.Shapes.AddMediaObject2($target, -1, 0, 180, 120, 600, 300)
  $vid.Name = 'online-video'

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output "SAVED: $out"
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp   -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
  Remove-Item $target -Force -ErrorAction SilentlyContinue
}
