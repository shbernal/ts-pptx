$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'placeholder-footer-trio.pptx')
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

  $slide = $pres.Slides.Add(1, 11)  # ppLayoutTitleOnly

  # Move the master's date/footer/slide-number placeholders to three visibly
  # distinct boxes (points). The layout keeps NO own a:xfrm for these, so the
  # slide placeholders inherit geometry all the way through to these boxes.
  # Type ids: Date=16, Footer=15, SlideNumber=13.
  $master = $pres.SlideMaster
  foreach ($ph in $master.Shapes.Placeholders) {
    switch ($ph.PlaceholderFormat.Type) {
      16 { $ph.Left = 40;  $ph.Top = 480; $ph.Width = 200; $ph.Height = 40 }   # Date
      15 { $ph.Left = 300; $ph.Top = 500; $ph.Width = 360; $ph.Height = 30 }   # Footer
      13 { $ph.Left = 780; $ph.Top = 460; $ph.Width = 140; $ph.Height = 50 }   # SlideNumber
    }
  }

  # Enable the footer trio on the slide so it materialises the three placeholder
  # shapes (with no own a:xfrm).
  $hf = $slide.HeadersFooters
  $hf.Footer.Visible = $true
  $hf.Footer.Text = "TRIO FOOTER"
  $hf.SlideNumber.Visible = $true
  $hf.DateAndTime.Visible = $true
  $hf.DateAndTime.UseFormat = $true

  # A visible title so the slide is not blank on inspection.
  $slide.Shapes.Title.TextFrame.TextRange.Text = "Footer trio geometry inheritance"

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
