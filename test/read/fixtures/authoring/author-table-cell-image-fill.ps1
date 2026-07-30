$ErrorActionPreference = 'Stop'

# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX  = Join-Path $REPO 'test\read\fixtures'
$out  = Join-Path $FIX 'table-cell-image-fill.pptx'
$img  = Join-Path $REPO 'demos\common\images\cc_logo.jpg'

if (-not (Test-Path $img)) { throw "source image not found: $img" }
if (Test-Path $out) { Remove-Item $out -Force }

# PpBorderType
$ppBorderTop = 1; $ppBorderLeft = 2; $ppBorderBottom = 3; $ppBorderRight = 4
# "No Style, No Grid" — keeps the table style from contributing fills, so every
# fill in the saved XML is one this fixture set explicitly.
$noStyleNoGrid = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank

  # 4 rows x 2 cols, deterministic geometry.
  $shape = $slide.Shapes.AddTable(4, 2, 60, 60, 840, 400)
  $shape.Name = 'CellImageFillTable'
  $table = $shape.Table
  $table.ApplyStyle($noStyleNoGrid, $true)

  # --- Row 1: A1 picture fill (stretched) | A2 solid fill -------------------
  $table.Cell(1,1).Shape.TextFrame.TextRange.Text = 'A1 picture'
  $table.Cell(1,1).Shape.Fill.UserPicture($img)

  $table.Cell(1,2).Shape.TextFrame.TextRange.Text = 'A2 solid'
  $table.Cell(1,2).Shape.Fill.Solid()
  $table.Cell(1,2).Shape.Fill.ForeColor.RGB = 255   # 0x0000FF BGR = red

  # --- Row 2: B1 picture fill + explicit borders | B2 no fill ---------------
  $table.Cell(2,1).Shape.TextFrame.TextRange.Text = 'B1 picture+borders'
  $table.Cell(2,1).Shape.Fill.UserPicture($img)
  foreach ($side in @($ppBorderTop, $ppBorderLeft, $ppBorderBottom, $ppBorderRight)) {
    $b = $table.Cell(2,1).Borders($side)
    $b.ForeColor.RGB = 16711680   # 0xFF0000 BGR = blue
    $b.Weight = 3
    $b.Visible = -1               # msoTrue
  }

  $table.Cell(2,2).Shape.TextFrame.TextRange.Text = 'B2 plain'

  # --- Row 3: merged across both columns, picture fill ----------------------
  $table.Cell(3,1).Shape.TextFrame.TextRange.Text = 'C merged picture'
  $table.Cell(3,1).Merge($table.Cell(3,2))
  $table.Cell(3,1).Shape.Fill.UserPicture($img)

  # --- Row 4: D1 TILED picture fill | D2 plain ------------------------------
  $table.Cell(4,1).Shape.TextFrame.TextRange.Text = 'D1 tiled'
  $table.Cell(4,1).Shape.Fill.UserPicture($img)
  $table.Cell(4,1).Shape.Fill.TextureTile = -1   # msoTrue -> tile rather than stretch

  $table.Cell(4,2).Shape.TextFrame.TextRange.Text = 'D2 plain'

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output "SAVED: $out"
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
