$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'picture-media.pptx')
$png = (Join-Path $SCRATCH 'media\pic.png')
$svg = (Join-Path $SCRATCH 'media\pic.svg')
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
$msoFalse = 0
$msoTrue = -1
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

  # G6: SVG picture (Insert-a-picture of an .svg). Test whether PowerPoint
  # writes svg-only or a raster+svg 'both'.
  $svgPic = $slide.Shapes.AddPicture($svg, $msoFalse, $msoTrue, 40, 40, 120, 120)
  $svgPic.Name = 'SvgPic'

  # G7: cropped raster picture + G8 alt text (description).
  $pic = $slide.Shapes.AddPicture($png, $msoFalse, $msoTrue, 240, 40, 260, 195)
  $pic.Name = 'CroppedPic'
  $pic.PictureFormat.CropLeft = 30
  $pic.PictureFormat.CropRight = 15
  $pic.PictureFormat.CropTop = 20
  $pic.PictureFormat.CropBottom = 10
  $pic.AlternativeText = 'A cropped stopwatch photo'

  # G8: decorative autoshape.
  $rect = $slide.Shapes.AddShape(1, 40, 320, 160, 90) # msoShapeRectangle
  $rect.Name = 'DecoRect'
  try {
    $rect.Decorative = $msoTrue
    Write-Host 'Decorative set via Shape.Decorative'
  } catch {
    Write-Host ("Shape.Decorative FAILED: " + $_.Exception.Message)
  }

  # G8: a shape with a plain description (no crop), to test description alone.
  $rect2 = $slide.Shapes.AddShape(1, 240, 320, 160, 90)
  $rect2.Name = 'DescRect'
  $rect2.AlternativeText = 'A described rectangle'

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Host 'SAVED OK'
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
