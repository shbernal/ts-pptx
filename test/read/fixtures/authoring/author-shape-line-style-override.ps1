param(
  # Where to write the deck. Defaults to the fixture this recipe produces, beside `authoring/`.
  [string]$Out = (Join-Path $PSScriptRoot '..\shape-line-style-override.pptx'),
  # Where the probe's PNG export and read-back land. Not part of the fixture.
  [string]$ProbeDir = (Join-Path $PSScriptRoot '..\..\..\..\.tmp\shape-line-style-override')
)
# shape-line-style-override.pptx: what PowerPoint writes when one property of a styled shape's
# outline is changed, and so what an `a:ln` beside `p:style/a:lnRef` means.
#   StyleOnly   the preset shape style and nothing else          -> no spPr/a:ln
#   WeightOnly  the same style, then Line.Weight = 6pt           -> a:ln @w only?
#   DashOnly    the same style, then Line.DashStyle = msoLineDash -> a:ln/a:prstDash only?
#   ColorOnly   the same style, then Line.ForeColor.RGB = red    -> a:ln/a:solidFill (control)
# After saving, the deck is reopened read-only and slide 1 exported to PNG, and each shape's
# Line.ForeColor/Weight is printed: the pixels are what says whether the style colour survives.
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
$ProbeDir = [IO.Path]::GetFullPath($ProbeDir)
if (Test-Path $Out) { Remove-Item $Out -Force }
New-Item -ItemType Directory -Force $ProbeDir | Out-Null

# Recovery entries a force-killed PowerPoint leaves behind keep the next launch modal.
foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(0)            # msoFalse: no window
  $slide = $pres.Slides.Add(1, 12)            # ppLayoutBlank

  # msoShapeStylePreset10: a coloured fill with a darker outline of the same accent, so the
  # outline colour differs from both the fill and black.
  $cases = @(
    @{ Name = 'StyleOnly';  Left = 60;  Top = 60 },
    @{ Name = 'WeightOnly'; Left = 500; Top = 60 },
    @{ Name = 'DashOnly';   Left = 60;  Top = 300 },
    @{ Name = 'ColorOnly';  Left = 500; Top = 300 }
  )
  foreach ($case in $cases) {
    $shape = $slide.Shapes.AddShape(1, $case.Left, $case.Top, 360, 180)   # msoShapeRectangle
    $shape.Name = $case.Name
    $shape.ShapeStyle = 10
    $shape.TextFrame.TextRange.Text = $case.Name
  }
  $slide.Shapes.Item('WeightOnly').Line.Weight = 6
  $slide.Shapes.Item('DashOnly').Line.DashStyle = 4          # msoLineDash
  $slide.Shapes.Item('ColorOnly').Line.ForeColor.RGB = 255   # 0x0000FF little-endian: red

  $pres.SaveAs($Out)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # Reopen what was saved, so the read-back and the pixels describe the file rather than the session.
  $pres = $pp.Presentations.Open($Out, -1, 0, 0)            # ReadOnly, Untitled=false, WithWindow=false
  foreach ($case in $cases) {
    $line = $pres.Slides.Item(1).Shapes.Item($case.Name).Line
    Write-Output ('READBACK' + "`t" + $case.Name + "`t" + ('{0:X6}' -f $line.ForeColor.RGB) + "`t" + $line.Weight + "`t" + $line.DashStyle)
  }
  $png = Join-Path $ProbeDir 'slide1.png'
  $pres.Slides.Item(1).Export($png, 'PNG', 1280, 720)
  Write-Output ('PNG' + "`t" + $png)
  $pres.Close()
  $pres = $null
  $pp.Quit()
  Write-Output ('SAVED: ' + $Out)
}
finally {
  if ($pres -ne $null) { try { $pres.Close() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
