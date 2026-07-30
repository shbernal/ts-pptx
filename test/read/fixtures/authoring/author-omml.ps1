$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'math-omml.pptx')
if (Test-Path $out) { Remove-Item $out -Force }

$w = $null; $doc = $null; $pp = $null; $pres = $null
try {
  # --- 1. Build a real equation in Word (genuine OMML, not hand-typed) ---
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $doc = $w.Documents.Add()
  $rng = $doc.Content
  $rng.Text = 'x^2+1=y'                 # linear math
  $null = $doc.OMaths.Add($rng)
  $om = $doc.OMaths.Item(1)
  $null = $om.BuildUp()                 # linear -> professional OMML
  $om.Range.Select()
  $w.Selection.Copy()                   # equation now on the clipboard

  # --- 2. Paste it into a PowerPoint text box; PowerPoint re-serialises on save ---
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth  = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)      # ppLayoutBlank

  # msoTextOrientationHorizontal = 1
  $tb = $slide.Shapes.AddTextbox(1, 120, 200, 700, 120)
  $tb.Name = 'equation-box'
  $tr = $tb.TextFrame.TextRange
  $tr.Text = ''
  $null = $tr.Paste()                   # paste clipboard equation into the text body

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()

  $doc.Close(0)                         # wdDoNotSaveChanges
  $w.Quit()
  Write-Output ('SAVED: ' + $out)
}
catch {
  Write-Output ('FAIL: ' + $_.Exception.Message)
  if ($pres -ne $null) { try { $pres.Close() } catch {} }
  if ($pp   -ne $null) { try { $pp.Quit()   } catch {} }
  if ($doc  -ne $null) { try { $doc.Close(0) } catch {} }
  if ($w    -ne $null) { try { $w.Quit()    } catch {} }
  throw
}
finally {
  foreach ($o in @($pres,$pp,$doc,$w)) { if ($o -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
