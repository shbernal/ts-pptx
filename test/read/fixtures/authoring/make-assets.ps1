$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
Add-Type -AssemblyName System.Drawing
$dir = (Join-Path $SCRATCH 'media')
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

# A brand-free photo-ish raster (gradient + shapes) so an artistic effect has detail to encode into a .wdp.
$w = 320; $h = 240
$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, [System.Drawing.Color]::FromArgb(40,90,160), [System.Drawing.Color]::FromArgb(200,120,40), 45.0)
$g.FillRectangle($brush, $rect)
for ($i = 0; $i -lt 24; $i++) {
  $c = [System.Drawing.Color]::FromArgb(180, (30*$i)%255, (70+9*$i)%255, (200-5*$i)%255)
  $b2 = New-Object System.Drawing.SolidBrush($c)
  $g.FillEllipse($b2, 10+$i*11, 10+($i*7)%180, 60, 45)
  $b2.Dispose()
}
$g.Dispose()
$bmp.Save((Join-Path $dir 'photo.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

# A small flat PNG (for a plain / recolor target).
$bmp2 = New-Object System.Drawing.Bitmap(64, 64)
$g2 = [System.Drawing.Graphics]::FromImage($bmp2)
$g2.Clear([System.Drawing.Color]::FromArgb(90, 90, 90))
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 6)
$g2.DrawLine($pen, 8, 8, 56, 56); $g2.DrawLine($pen, 8, 56, 56, 8)
$g2.Dispose()
$bmp2.Save((Join-Path $dir 'mark.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp2.Dispose()

Get-ChildItem $dir | Select-Object Name, Length | Format-Table -AutoSize
