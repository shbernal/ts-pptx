$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'math-omml-inline.pptx')
if (Test-Path $out) { Remove-Item $out -Force }

# Snapshot pre-existing automation servers so the reap only kills ones we spawn.
$prePP = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$preW  = @(Get-Process WINWORD  -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$w = $null; $doc = $null; $pp = $null; $pres = $null
try {
  # --- 1. Build an INLINE equation in Word, mid-sentence between plain text ---
  # A math zone embedded within a paragraph that also carries ordinary text is
  # an *inline* (not display) equation: Word emits <m:oMath> with no <m:oMathPara>.
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $doc = $w.Documents.Add()
  $rng = $doc.Content
  $rng.Text = 'where x^2+1=y holds'   # one paragraph: text + linear math + text

  # Wrap just the "x^2+1=y" substring (6 chars in, 7 chars long) as a math zone.
  $eqStart = $doc.Content.Start + 6
  $eqEnd   = $eqStart + 7
  $eqRange = $doc.Range($eqStart, $eqEnd)
  $null = $doc.OMaths.Add($eqRange)
  $om = $doc.OMaths.Item(1)
  $null = $om.BuildUp()               # linear -> professional OMML
  Write-Output ('Word OMath.Type (0=display,1=inline): ' + $om.Type)

  # Copy the WHOLE paragraph (surrounding text runs + inline equation).
  $doc.Paragraphs.Item(1).Range.Select()
  $w.Selection.Copy()

  # --- 2. Paste into a PowerPoint text box; PowerPoint re-serialises on save ---
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth  = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)    # ppLayoutBlank

  # msoTextOrientationHorizontal = 1
  $tb = $slide.Shapes.AddTextbox(1, 120, 200, 700, 120)
  $tb.Name = 'inline-equation-box'
  $tr = $tb.TextFrame.TextRange
  $tr.Text = ''
  $null = $tr.Paste()                 # paste the mixed text+inline-equation paragraph

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()

  $doc.Close(0)                       # wdDoNotSaveChanges
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
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $prePP -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process WINWORD  -ErrorAction SilentlyContinue | Where-Object { $preW  -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
}
