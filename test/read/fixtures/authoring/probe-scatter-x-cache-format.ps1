# Probe; produces no committed fixture. Builds the scatter sensitivity trio with
# probe-scatter-x-cache-format.mjs (it needs a built dist/) and exports each to PNG in PowerPoint,
# so the X axis labels can be read off the picture. The findings are recorded in the .mjs header.
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\..'))
$outDir = Join-Path $repo '.tmp\scatter-x-cache-format'
& node (Join-Path $PSScriptRoot 'probe-scatter-x-cache-format.mjs')
if ($LASTEXITCODE -ne 0) { throw 'building the decks failed' }

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}
$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  foreach ($name in 'as-written', 'cache-patched', 'axis-unlinked') {
    $deck = Join-Path $outDir "$name.pptx"
    $pres = $pp.Presentations.Open($deck, -1, 0, 0)
    $pres.Slides.Item(1).Export((Join-Path $outDir "$name.png"), 'PNG', 1280, 720)
    $pres.Close()
    Write-Output ('RENDERED ' + (Join-Path $outDir "$name.png"))
  }
  $pp.Quit()
}
finally {
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
