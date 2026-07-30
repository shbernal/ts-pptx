$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$deck = (Join-Path $FIX 'multi-theme.pptx')

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Open($deck)       # already Ion-themed; new slide inherits Ion
  $slide = $pres.Slides.Add(3, 12)            # ppLayoutBlank, after existing 2 slides

  # (a) Rectangle filled with a LITERAL srgbClr equal to Ion accent1 (B01513).
  $r1 = $slide.Shapes.AddShape(1, 60, 60, 200, 120)   # msoShapeRectangle
  $r1.Name = 'literal-accent1'
  $r1.Fill.Solid()
  $r1.Fill.ForeColor.RGB = 1250736            # BGR of B01513 -> <a:srgbClr val="B01513"/>
  $r1.Line.Visible = 0

  # Negative control: literal srgbClr matching NO Ion theme slot (123456).
  $r2 = $slide.Shapes.AddShape(1, 300, 60, 200, 120)
  $r2.Name = 'literal-nonaccent'
  $r2.Fill.Solid()
  $r2.Fill.ForeColor.RGB = 5649426            # BGR of 123456
  $r2.Line.Visible = 0

  # (b) Table with a NON-DEFAULT built-in table style (Medium Style 2 - Accent 1).
  $t = $slide.Shapes.AddTable(2, 2, 60, 240, 400, 160)
  $t.Name = 'styled-table'
  $t.Table.ApplyStyle('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}')

  $pres.Save()                                # in place — preserve existing parts where possible
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output "SAVED: $deck"
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp   -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
