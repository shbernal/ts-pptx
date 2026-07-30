$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'modern-comments.pptx')
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $s1 = $pres.Slides.Add(1, 12) # ppLayoutBlank (no comments)
  $s2 = $pres.Slides.Add(2, 12) # ppLayoutBlank (threaded comment)

  # Give slide 2 a named shape so the comment can (potentially) anchor to it.
  $box = $s2.Shapes.AddTextbox(1, 120, 80, 400, 60)
  $box.Name = 'Headline'
  $box.TextFrame.TextRange.Text = 'Draft headline'

  # Modern threaded comment: Add2(Left, Top, Author, AuthorInitials, Text, providerId, userId)
  $c = $s2.Comments.Add2(120, 80, 'Ada Lovelace', 'AL', 'Tighten this headline', 'Windows Live', 'ada@example.com')
  Write-Host "comment Add2 OK"

  # Reply by a different author -> exercises the reply thread + a second p188:author
  $r = $c.Replies.Add2(0, 0, 'Grace Hopper', 'GH', 'Agreed, will do', 'Windows Live', 'grace@example.com')
  Write-Host "reply Add2 OK"

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
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
Write-Host "SAVED: $out"
