# Authors test/read/fixtures/default-text-style.pptx via PowerPoint COM.
# Goal: a minimal deck that exercises the two lowest run-resolution tiers:
#   - PlainBox   : a plain text box whose run sets no own font -> must fall back
#                  to the presentation p:defaultTextStyle (size/colour/face).
#   - StyledRect : a rectangle with a theme shape style (-> p:style/a:fontRef),
#                  whose run sets no own font -> colour/face come from the fontRef.
$ErrorActionPreference = 'Stop'

# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX  = Join-Path $REPO 'test\read\fixtures'
$out  = Join-Path $FIX 'default-text-style.pptx'
if (Test-Path $out) { Remove-Item $out -Force }

$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Add($false)   # msoFalse = no window
    # ppLayoutBlank = 12
    $slide = $pres.Slides.Add(1, 12)

    # --- PlainBox: plain text box, no explicit run font ---
    # msoTextOrientationHorizontal=1; left, top, width, height in points
    $tb = $slide.Shapes.AddTextbox(1, 60, 60, 360, 60)
    $tb.Name = 'PlainBox'
    $tb.TextFrame.TextRange.Text = 'Plain default text'

    # --- StyledRect: rectangle with a theme shape style (carries p:style/fontRef) ---
    # msoShapeRectangle = 1
    $rect = $slide.Shapes.AddShape(1, 60, 180, 360, 90)
    $rect.Name = 'StyledRect'
    # Apply a preset theme shape style so PowerPoint writes a p:style with an
    # a:fontRef (accent-coloured fill + contrasting font colour). 60 is a mid
    # "Colored Fill - Accent" style in the shape-style gallery.
    $rect.ShapeStyle = 60
    $rect.TextFrame.TextRange.Text = 'Styled rect text'

    $pres.SaveAs($out)          # 24 = ppSaveAsOpenXMLPresentation (default by extension)
    $pres.Close()
    Write-Host "WROTE $out"
}
finally {
    $pp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
