# Probe; produces no committed fixture. Builds the combo charts with probe-combo-legend-order.mjs
# (it needs a built dist/) and exports each to PNG in PowerPoint, so the legend order can be read
# off the picture. The findings are recorded in the .mjs header.
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\..'))
$outDir = Join-Path $repo '.tmp\combo-legend-order'
& node (Join-Path $PSScriptRoot 'probe-combo-legend-order.mjs')
if ($LASTEXITCODE -ne 0) { throw 'building the decks failed' }

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}
$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  foreach ($deck in Get-ChildItem $outDir -Filter '*.pptx' | Sort-Object Name) {
    $pres = $pp.Presentations.Open($deck.FullName, -1, 0, 0)
    $pres.Slides.Item(1).Export([IO.Path]::ChangeExtension($deck.FullName, '.png'), 'PNG', 1280, 720)
    $pres.Close()
    Write-Output ('RENDERED ' + [IO.Path]::ChangeExtension($deck.FullName, '.png'))
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
