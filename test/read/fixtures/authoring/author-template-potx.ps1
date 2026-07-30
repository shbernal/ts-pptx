$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'template.potx')
# Snapshot pre-existing PIDs so the reap at the end only kills the server we spawn.
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(0)        # msoFalse: no window
  # 16:9 widescreen, the modern PowerPoint default (13.333in x 7.5in = 960 x 540 pt)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540

  # A genuine .potx typically carries zero sample slides. Strip any default slide
  # so the template is a pure master/layout/theme shell.
  while ($pres.Slides.Count -gt 0) { $pres.Slides.Item(1).Delete() }

  Write-Host ("Slides: {0}" -f $pres.Slides.Count)
  Write-Host ("Masters: {0}" -f $pres.Designs.Count)
  $master = $pres.SlideMaster
  Write-Host ("Layouts: {0}" -f $master.CustomLayouts.Count)
  for ($i = 1; $i -le $master.CustomLayouts.Count; $i++) {
    Write-Host ("  layout[{0}] = {1}" -f $i, $master.CustomLayouts.Item($i).Name)
  }

  # ppSaveAsOpenXMLTemplate = 26 -> .potx (presentationml.template.main+xml).
  # NB: 27 is ppSaveAsOpenXMLTemplateMacroEnabled (.potm) — the wrong content-type.
  $pres.SaveAs($out, 26)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
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
