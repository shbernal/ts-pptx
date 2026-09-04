$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX  = Join-Path $REPO 'test\read\fixtures'
$out = (Join-Path $FIX 'placeholder-override.pptx')
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

  # The LAYOUT's body placeholder states three things: a frame, a bottom anchor, and a
  # left inset far from the 0.1in default. Every one of them is a property a slide
  # placeholder could also state, which is the whole point of the fixture.
  $layout = $pres.SlideMaster.CustomLayouts(2)   # Title and Content
  $lb = $layout.Shapes.Placeholders(2)
  $lb.TextFrame2.VerticalAnchor = 4              # msoAnchorBottom
  $lb.TextFrame2.MarginLeft = 72                 # 1in
  $lb.Left = 100; $lb.Top = 100; $lb.Width = 300; $lb.Height = 200

  # The SLIDE's body placeholder overrides two of them -- the anchor and the frame -- and
  # says nothing about the inset. What PowerPoint writes here is the evidence: only the
  # overridden properties appear on the slide shape, and the inset is simply absent.
  $slide = $pres.Slides.AddSlide(1, $layout)
  $slide.Shapes.Title.TextFrame.TextRange.Text = "Placeholder override"
  $sb = $slide.Shapes.Placeholders(2)
  $sb.TextFrame.TextRange.Text = "slide body"
  $sb.TextFrame2.VerticalAnchor = 1              # msoAnchorTop -- overrides the layout's 'b'
  $sb.Left = 400                                 # and a frame the layout also states

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
