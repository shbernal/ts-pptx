param([string]$OutPath)
$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$out = if ($OutPath) { $OutPath } else { Join-Path $FIX 'picture-custgeom.pptx' }
# Brand-free raster drawn in code; run make-assets.ps1 first.
$png = Join-Path $SCRATCH 'media\photo.png'
if (-not (Test-Path $png)) { throw "missing $png; run make-assets.ps1 first" }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
$msoFalse = 0
$msoTrue = -1
$msoEditingCorner = 1
$msoSegmentLine = 0
$msoSegmentCurve = 1

# Merge Shapes has no late-bindable COM method (its enum argument refuses to marshal), so the
# ribbon command runs on a selection instead, the way custgeom.pptx's hole was made. The picture is
# selected first: the first shape in the selection decides what the merged result is.
function Merge-Intersect($slide, $picture, $clip, $name) {
  $picture.Select($msoTrue)
  $clip.Select($msoFalse)
  $pp.CommandBars.ExecuteMso('ShapesIntersect')
  $merged = $pp.ActiveWindow.Selection.ShapeRange.Item(1)
  $merged.Name = $name
  Write-Host ("{0}: Type={1} (13 = msoPicture)" -f $name, $merged.Type)
}

try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  # With a window: ExecuteMso acts on the active window's selection.
  $pres = $pp.Presentations.Add($msoTrue)
  $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank
  $pp.ActiveWindow.View.GotoSlide(1)

  # A picture clipped to a straight-edged freeform: moveTo, lnTo, close.
  $pic = $slide.Shapes.AddPicture($png, $msoFalse, $msoTrue, 40, 40, 260, 195)
  $ff = $slide.Shapes.BuildFreeform($msoEditingCorner, 60, 50)
  $ff.AddNodes($msoSegmentLine, $msoEditingCorner, 280, 50)
  $ff.AddNodes($msoSegmentLine, $msoEditingCorner, 170, 225)
  $ff.AddNodes($msoSegmentLine, $msoEditingCorner, 60, 50)
  $triangle = $ff.ConvertToShape()
  Merge-Intersect $slide $pic $triangle 'pic-clip-lines'

  # A picture clipped to a freeform with a curved edge: cubicBezTo control points.
  $pic = $slide.Shapes.AddPicture($png, $msoFalse, $msoTrue, 340, 40, 260, 195)
  $ff = $slide.Shapes.BuildFreeform($msoEditingCorner, 360, 120)
  $ff.AddNodes($msoSegmentCurve, $msoEditingCorner, 400, 40, 560, 40, 580, 140)
  $ff.AddNodes($msoSegmentLine, $msoEditingCorner, 580, 225)
  $ff.AddNodes($msoSegmentLine, $msoEditingCorner, 360, 225)
  $ff.AddNodes($msoSegmentLine, $msoEditingCorner, 360, 120)
  $arch = $ff.ConvertToShape()
  Merge-Intersect $slide $pic $arch 'pic-clip-curve'

  # Negative controls: a picture cropped to a preset shape, and a plain picture.
  $pic = $slide.Shapes.AddPicture($png, $msoFalse, $msoTrue, 640, 40, 260, 195)
  $pic.Name = 'pic-preset-oval'
  $pic.AutoShapeType = 9 # msoShapeOval
  $pic = $slide.Shapes.AddPicture($png, $msoFalse, $msoTrue, 40, 300, 200, 150)
  $pic.Name = 'pic-plain'

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Host "SAVED $out"
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
