$ErrorActionPreference = 'Stop'

# Produces ../slide-background.pptx — the read-side oracle for a background that belongs to a
# SLIDE rather than to its layout or master. Three slides, one construct each:
#
#   1  p:bg/p:bgPr/a:blipFill   a picture background (Format Background > Picture or texture
#                               fill, applied to one slide only)
#   2  p:bg/p:bgRef             a theme background-style reference, slide-scoped
#   3  p:bg/p:bgPr/a:solidFill  a solid with an a:alpha, i.e. a transparent background
#
# The corpus had none of these at slide scope: every `source: 'slide'` background across the
# other fixtures is `solid` or `none`, so the converter's slide arm could claim a picture
# background was "not expressible through the write API" with nothing to contradict it.
#
# Slides 1 and 3 are plain COM. Slide 2 is not: `p:bgRef` is what the Design > Variants >
# Background Styles gallery writes, and a gallery has no `ExecuteMso` surface and no
# `Slide.Background` equivalent — `Slide.Background.Fill` reaches `p:bgPr` and only `p:bgPr`.
# So slide 2 takes the inject-then-reopen route `author-table-cell-horzoverflow.ps1` uses:
# inject the element, hand the deck back to PowerPoint, and let it re-serialize. PowerPoint
# keeps the `p:bgRef` verbatim, which is itself the evidence that the construct is one it
# accepts at slide scope (contrast `probe-table-cell-wrap.ps1`, where an injected
# `<a:bodyPr wrap="none"/>` is silently dropped on the same round trip).
#
# `idx="1001"` is the first `a:bgFillStyleLst` entry, which in the stock Office theme is
# `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` — so the reference resolves to a flat
# accent2, and the `themeRef` leg's *solid* path is what gets exercised. The image comes from
# `make-assets.ps1`, so the fixture's only external input is license-clean by construction.

# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$SCRATCH = Join-Path $REPO '.tmp'

$out = Join-Path $FIX 'slide-background.pptx'
$tmp = Join-Path $FIX 'slide-background.base.pptx'
$png = Join-Path $SCRATCH 'media\photo.png'
foreach ($f in @($out, $tmp)) { if (Test-Path $f) { Remove-Item $f -Force } }
if (-not (Test-Path $png)) { & (Join-Path $PSScriptRoot 'make-assets.ps1') }

$BGREF = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent2"/></p:bgRef></p:bg>'

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
$msoFalse = 0
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1

  # --- author: slides 1 and 3, the two COM can express -------------------------
  $pres = $pp.Presentations.Add(1)

  $s1 = $pres.Slides.Add(1, 12)   # ppLayoutBlank
  $s1.FollowMasterBackground = $msoFalse
  $s1.Background.Fill.UserPicture($png)

  # Slide 2 keeps the master's background for now; the injection below gives it its own.
  [void]$pres.Slides.Add(2, 12)

  $s3 = $pres.Slides.Add(3, 12)
  $s3.FollowMasterBackground = $msoFalse
  $s3.Background.Fill.Solid()
  $s3.Background.Fill.ForeColor.RGB = 0x283CC8   # 0x00BBGGRR -> C83C28
  $s3.Background.Fill.Transparency = 0.4

  $pres.SaveAs($tmp)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # --- inject: slide 2's own p:bgRef, ahead of its p:spTree --------------------
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::Open($tmp, 'Update')
  try {
    $entry = $zip.GetEntry('ppt/slides/slide2.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()

    if (([regex]::Matches($xml, '<p:cSld><p:spTree>')).Count -ne 1) { throw 'slide 2 does not open with a bare p:cSld/p:spTree' }
    $xml = $xml.Replace('<p:cSld><p:spTree>', ('<p:cSld>' + $BGREF + '<p:spTree>'))

    $entry.Delete()
    $new = $zip.CreateEntry('ppt/slides/slide2.xml')
    $writer = New-Object System.IO.StreamWriter($new.Open())
    $writer.Write($xml)
    $writer.Close()
  }
  finally { $zip.Dispose() }

  # --- hand it back to PowerPoint: it must open clean and re-serialize the bgRef -
  $pres = $pp.Presentations.Open($tmp, 0, 0, 0)   # throws on a repair prompt
  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null
  $pp.Quit()
  Remove-Item $tmp -Force

  # --- verify what PowerPoint wrote --------------------------------------------
  $zip = [System.IO.Compression.ZipFile]::OpenRead($out)
  try {
    foreach ($n in 1..3) {
      $entry = $zip.GetEntry("ppt/slides/slide$n.xml")
      $reader = New-Object System.IO.StreamReader($entry.Open())
      $saved = $reader.ReadToEnd()
      $reader.Close()
      $m = [regex]::Match($saved, '<p:bg>.*?</p:bg>')
      Write-Output ("  slide{0}: {1}" -f $n, $(if ($m.Success) { $m.Value } else { '(no p:bg)' }))
    }
  }
  finally { $zip.Dispose() }

  Write-Output ("SAVED: {0}" -f $out)
  Write-Output ("SHA256: {0}" -f (Get-FileHash $out -Algorithm SHA256).Hash.ToLower())
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
