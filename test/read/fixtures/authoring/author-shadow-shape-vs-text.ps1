$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX  = Join-Path $REPO 'test\read\fixtures'
$out = (Join-Path $FIX 'shadow-shape-vs-text.pptx')
if (Test-Path $out) { Remove-Item $out -Force }
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(0)  # msoFalse: no window
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)  # ppLayoutBlank

  # Three text boxes, one per state. PowerPoint offers a shape shadow (Shape Effects) and a text
  # shadow (Text Effects) as two separate gestures; the fixture's job is to show that they land in
  # two separate places and that neither one produces the other.
  #
  # `.Shadow` is the SHAPE's ShadowFormat; `.TextFrame2.TextRange.Font.Shadow` is the runs'. Both
  # take the same preset (msoShadow21 -> type 1 here) so the two effects differ only in where they
  # are written, which is the whole comparison.
  $a = $slide.Shapes.AddTextbox(1, 40, 40, 300, 60)
  $a.Name = 'ShapeShadowOnly'
  $a.TextFrame.TextRange.Text = 'shape shadow'
  $a.Shadow.Visible = -1
  $a.Shadow.Type = 1

  $b = $slide.Shapes.AddTextbox(1, 40, 140, 300, 60)
  $b.Name = 'TextShadowOnly'
  $b.TextFrame.TextRange.Text = 'text shadow'
  $b.TextFrame2.TextRange.Font.Shadow.Visible = -1
  $b.TextFrame2.TextRange.Font.Shadow.Type = 1

  # Both at once takes TWO user actions and is the state a single `shadow` option must not produce.
  $c = $slide.Shapes.AddTextbox(1, 40, 240, 300, 60)
  $c.Name = 'BothShadows'
  $c.TextFrame.TextRange.Text = 'both shadows'
  $c.Shadow.Visible = -1
  $c.Shadow.Type = 1
  $c.TextFrame2.TextRange.Font.Shadow.Visible = -1
  $c.TextFrame2.TextRange.Font.Shadow.Type = 1

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Host "SAVED $out"
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
