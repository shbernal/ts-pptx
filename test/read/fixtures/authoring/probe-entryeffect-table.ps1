$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = Join-Path $SCRATCH 'entryeffect-table.pptx'
$map = Join-Path $SCRATCH 'entryeffect-order.json'
if (Test-Path $out) { Remove-Item $out -Force }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)

  # Discover the valid PpEntryEffect ints on a scratch slide.
  $probe = $pres.Slides.Add(1, 12) # ppLayoutBlank
  $defaultEffect = $probe.SlideShowTransition.EntryEffect
  $valid = New-Object System.Collections.Generic.List[int]
  foreach ($v in 0..0x1000) {
    try { $probe.SlideShowTransition.EntryEffect = $v; $valid.Add($v) } catch {}
  }
  Write-Host ("default EntryEffect = " + $defaultEffect)
  Write-Host ("VALID COUNT (0..4096): " + $valid.Count)

  # One slide per valid int. Slide 1 is the scratch probe; real slides start at 2,
  # so committed slide index = arrayIndex + 2. Set a non-default Duration (1.0s) so
  # PowerPoint bakes the mc:AlternateContent (p14:dur) form and we capture p14 variants.
  $order = @()
  foreach ($v in $valid) {
    $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)
    $t = $s.SlideShowTransition
    $t.EntryEffect = $v
    $t.Duration = 1.0
    $t.AdvanceOnClick = -1  # msoTrue
    $order += [int]$v
  }

  # Remove the scratch probe slide so slide N (1-based) == valid[N-1].
  $pres.Slides.Item(1).Delete()

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()

  $order | ConvertTo-Json -Compress | Set-Content -Path $map -Encoding UTF8
  Write-Host ("SAVED: $out  (" + $order.Count + " transition slides)")
  Write-Host ("ORDER: $map")
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
