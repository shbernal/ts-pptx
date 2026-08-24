$ErrorActionPreference = 'Stop'
# Authors ../autofit-cjk-wrap.pptx plus its ../autofit-cjk-wrap.oracle.json sidecar:
# the PowerPoint-authored ground truth for how the app breaks lines in CJK text.
#
# Every case is one wrap=square text box of a fixed width, pinned before the text
# goes in and switched to msoAutoSizeShapeToFitText last, so PowerPoint bakes
# <a:spAutoFit/> and a fitted ext.cy (the authoring skill bake-on-save contract).
# Two things are recorded per case:
#   * lines / lineWidthsPt - read back over COM from TextRange.Lines(), i.e. the
#     breaks PowerPoint actually laid out. Nothing in the saved OOXML records
#     where a line broke, so this column only exists because it is captured here.
#   * bakedHeightPt - the shape height PowerPoint fitted, which IS in the package
#     (a:ext/@cy) and is re-checked against the committed deck by
#     test/read/cjk-line-breaking-oracle.test.js on every run.
#
# One font, Malgun Gothic: a Windows-standard plain .ttf (not a .ttc, which the
# metrics parser cannot open) that covers Han, Kana, Hangul, the fullwidth forms
# and CJK punctuation in one face, so a single set of advances explains every case.

# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX = Join-Path $REPO 'test\read\fixtures'
$out = Join-Path $FIX 'autofit-cjk-wrap.pptx'
$oracleOut = Join-Path $FIX 'autofit-cjk-wrap.oracle.json'

$FACE = 'Malgun Gothic'
$SIZE = 18.0
# Box width chosen so the discriminating cases land far from the wrap boundary:
# inner width = 164.4 - 2 x 7.2pt inset = 150pt.
$BOXW = 164.4

# --- font-presence guard: PowerPoint writes the requested typeface into the XML
# even when it substituted at render time, so check GDI, not the OOXML. ---
Add-Type -AssemblyName System.Drawing
$probe = New-Object System.Drawing.Font($FACE, $SIZE)
$resolved = $probe.Name
$probe.Dispose()
if ($resolved -ne $FACE) { throw "FONT SUBSTITUTED: $FACE -> $resolved" }

# Build strings from code points so this recipe stays pure ASCII on disk.
function U([int[]]$cps) { -join ($cps | ForEach-Object { [char]$_ }) }
function Astral([int[]]$cps) { -join ($cps | ForEach-Object { [char]::ConvertFromUtf32($_) }) }

$HAN5 = U @(0x4E00, 0x4E8C, 0x4E09, 0x56DB, 0x4E94)
$KANA5 = U @(0x3042, 0x3044, 0x3046, 0x3048, 0x304A)
$HANGUL5 = U @(0xAC00, 0xB098, 0xB2E4, 0xB77C, 0xB9C8)
$HANGUL14 = U @(0xAC00, 0xB098, 0xB2E4, 0xB77C, 0xB9C8, 0xBC14, 0xC0AC, 0xC544, 0xC790, 0xCC28, 0xCE74, 0xD0C0, 0xD30C, 0xD558)
$FULLW5 = U @(0xFF21, 0xFF22, 0xFF23, 0xFF24, 0xFF25)
$HALFK10 = U @(0xFF76, 0xFF77, 0xFF78, 0xFF79, 0xFF7A, 0xFF76, 0xFF77, 0xFF78, 0xFF79, 0xFF7A)
$EXTB5 = Astral @(0x20000, 0x20001, 0x20002, 0x20003, 0x20004)
$KANAPUN = U @(0x3042, 0x3044, 0x3001, 0x3046, 0x3048)

$cases = @(
  @{ id = 'cjk__han_between_words'
    note = 'Han run between two Latin words. Per-character breaking fills line 1; treating the run as one unbreakable word would move it wholesale and cost a third line.'
    text = "aaaaaaaa $HAN5 bbbbbbbb" }

  @{ id = 'cjk__kana_between_words'
    note = 'Same shape in Hiragana: Kana breaks per character exactly as Han does.'
    text = "aaaaaaaa $KANA5 bbbbbbbb" }

  @{ id = 'cjk__hangul_between_words'
    note = 'COUNTER-CASE. Same shape in Hangul syllables. PowerPoint does NOT break Hangul per syllable: the run moves to the next line whole, so this lays out one line taller than the Han/Kana cases. Korean writes spaces between words and the app breaks it like Latin.'
    text = "aaaaaaaa $HANGUL5 bbbbbbbb" }

  @{ id = 'cjk__hangul_run_alone'
    note = 'A Hangul run longer than the line still has to break somewhere, and does so between syllables. That is the over-long-token fallback, not word-internal breaking: it is what the previous case rules out.'
    text = $HANGUL14 }

  @{ id = 'cjk__han_run_alone'
    note = 'A 20-character Han run with no spaces at all: breaks per character.'
    text = ($HAN5 + $HAN5 + $HAN5 + $HAN5) }

  @{ id = 'cjk__han_latin_no_space'
    note = 'Script boundary with no whitespace anywhere: the break lands inside the Han run, and the trailing Latin word is not split.'
    text = "aaaaaaaa$HAN5" + 'bbbbbbbb' }

  @{ id = 'cjk__fullwidth_latin'
    note = 'Fullwidth Latin (U+FF21..) breaks per character like Han, confirming the Halfwidth and Fullwidth Forms block belongs in the break set.'
    text = "aaaaaaaa $FULLW5 bbbbbbbb" }

  @{ id = 'cjk__halfwidth_kana'
    note = 'Halfwidth Katakana (U+FF66..), also in the Halfwidth and Fullwidth Forms block but half an em wide. Ten of them, so per-character breaking splits the run and word breaking would not.'
    text = "aaaaaaaa $HALFK10 bbbbbbbb" }

  @{ id = 'cjk__ext_b_astral'
    note = 'CJK Extension B (Plane 2) breaks per character. Each is a surrogate pair, so this also pins that the tokenizer iterates code points, not UTF-16 units.'
    text = "aaaaaaaa $EXTB5 bbbbbbbb" }

  @{ id = 'cjk__kinsoku_hanging_comma'
    note = 'KNOWN GAP. The ideographic comma U+3001 may not start a line, so PowerPoint hangs it past the right inset rather than moving it down. The model has no kinsoku rules and breaks before it instead: same line count here, narrower widest line.'
    text = "aaaaaaaaaaa $KANAPUN" }

  @{ id = 'latin__control'
    note = 'All-Latin control at the same width and size: breaks only at whitespace, unchanged by any of the above.'
    text = 'aaaaaaaa bbbbbbbb cccccccc' }
)

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
$rows = @()
try {
  if (Test-Path $out) { Remove-Item $out -Force }

  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank

  $top = 12.0
  foreach ($c in $cases) {
    $sh = $slide.Shapes.AddTextbox(1, 24.0, $top, $BOXW, 40.0)   # msoTextOrientationHorizontal
    $sh.Name = $c.id

    # Pin the box first, text second, AutoSize last: that ordering is what makes
    # PowerPoint bake a fitted ext.cy instead of a bare <a:spAutoFit/>.
    $tf2 = $sh.TextFrame2
    $tf2.AutoSize = 0        # msoAutoSizeNone
    $tf2.WordWrap = -1       # msoTrue
    $sh.Width = $BOXW
    $sh.Height = 40.0

    $tr = $sh.TextFrame.TextRange
    $tr.Text = $c.text
    $tr.Font.Name = $FACE
    $tr.Font.NameFarEast = $FACE
    $tr.Font.Size = $SIZE

    $tf2.AutoSize = 1        # msoAutoSizeShapeToFitText

    $lines = $tr.Lines()
    $lineTexts = @()
    $lineWidths = @()
    for ($i = 1; $i -le $lines.Count; $i++) {
      $ln = $tr.Lines($i, 1)
      $lineTexts += $ln.Text
      $lineWidths += [Math]::Round([double]$ln.BoundWidth, 4)
    }

    $rows += [ordered]@{
      id = $c.id
      note = $c.note
      text = $c.text
      fontFace = $FACE
      sizePt = $SIZE
      boxWidthPt = [Math]::Round([double]$sh.Width, 4)
      insetLeftPt = [Math]::Round([double]$tf2.MarginLeft, 4)
      insetRightPt = [Math]::Round([double]$tf2.MarginRight, 4)
      insetTopPt = [Math]::Round([double]$tf2.MarginTop, 4)
      insetBottomPt = [Math]::Round([double]$tf2.MarginBottom, 4)
      bakedHeightPt = [Math]::Round([double]$sh.Height, 4)
      lineCount = [int]$lines.Count
      lines = $lineTexts
      lineWidthsPt = $lineWidths
    }
    $top += 44.0
  }

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

$oracle = [ordered]@{
  deck = 'autofit-cjk-wrap.pptx'
  recipe = 'test/read/fixtures/authoring/author-cjk-wrap.ps1'
  authoredBy = 'desktop Microsoft PowerPoint (COM)'
  notes = 'PowerPoint line breaking for CJK text. lines/lineWidthsPt come from TextRange.Lines() at authoring time (the package does not record where a line broke); bakedHeightPt is the fitted a:ext/@cy and is re-verified against the committed deck by the test.'
  fontFace = $FACE
  sizePt = $SIZE
  cases = $rows
}
$json = $oracle | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($oracleOut, $json, (New-Object Text.UTF8Encoding($false)))

Write-Output "wrote $out"
Write-Output "wrote $oracleOut"
$rows | ForEach-Object { '{0,-32} {1} line(s)  {2}' -f $_.id, $_.lineCount, ($_.lines -join ' | ') }
