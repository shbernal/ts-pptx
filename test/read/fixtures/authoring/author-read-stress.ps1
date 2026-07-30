# Authors test/read/fixtures/read-stress.pptx — a brand-free, PowerPoint-authored
# deck that combines, in ONE package, the read-model stress dimensions we found in
# a complex real-world deck: two live slide masters each with its own theme,
# nested groups, styled tables (styleId with no own cell fill + inline scheme fill
# with lumMod), an SVG picture (raster+svg dual-rel), grayscale/biLevel recolor,
# multi-typeface embedded fonts, threaded modern comments, and speaker notes.
#
# All content is synthetic (lorem-ish), uses only Microsoft built-in theme/style
# GUIDs, and carries no third-party brand assets.
$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
foreach ($k in 'DocumentRecovery','StartupItems') {
  $p = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$k"
  if (Test-Path $p) { Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue }
}

$out   = (Join-Path $FIX 'read-stress.pptx')
$svg   = (Join-Path $SCRATCH 'media\gear.svg')
$png   = (Join-Path $SCRATCH 'media\mark.png')
$thmxB = 'C:\Program Files\Microsoft Office\root\Document Themes 16\Facet.thmx'
$msoFalse = 0; $msoTrue = -1

# Built-in fill-bearing table style GUID (Microsoft's own — not a brand asset):
# Medium Style 2 - Accent 3. Chosen because it materialises real cell fills into
# ppt/tableStyles.xml, so a cell with no own fill resolves its fill from the style.
$STYLE_GUID = '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}'

$pre = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540

  # ============================ SLIDE 1 — master 1 (default theme) ============================
  # ppLayoutText (2) => title + body placeholders (exercises placeholder inheritance on master 1).
  $s1 = $pres.Slides.Add(1, 2)
  $s1.Shapes.Title.TextFrame.TextRange.Text = 'Combined read stress'
  # Body placeholder note text
  try { $s1.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Master one body placeholder' } catch {}

  # Embedded-font text in three distinct installed families.
  $f1 = $s1.Shapes.AddTextbox(1, 40, 300, 260, 30); $f1.Name = 'FontGeorgia'
  $f1.TextFrame.TextRange.Text = 'Georgia typeface'; $f1.TextFrame.TextRange.Font.Name = 'Georgia'
  $f2 = $s1.Shapes.AddTextbox(1, 40, 335, 260, 30); $f2.Name = 'FontConsolas'
  $f2.TextFrame.TextRange.Text = 'Consolas typeface'; $f2.TextFrame.TextRange.Font.Name = 'Consolas'
  $f3 = $s1.Shapes.AddTextbox(1, 40, 370, 260, 30); $f3.Name = 'FontTrebuchet'
  $f3.TextFrame.TextRange.Text = 'Trebuchet typeface'; $f3.TextFrame.TextRange.Font.Name = 'Trebuchet MS'

  # SVG picture -> raster+svg dual-rel (mediaKind 'both').
  $svgPic = $s1.Shapes.AddPicture($svg, $msoFalse, $msoTrue, 330, 300, 64, 64); $svgPic.Name = 'SvgIcon'

  # Grayscale + biLevel recolor (real PowerPoint-authored recolor the reader recognises).
  $gp = $s1.Shapes.AddPicture($png, $msoFalse, $msoTrue, 410, 300, 64, 64); $gp.Name = 'GrayPic'
  $gp.PictureFormat.ColorType = 2   # ppPictureGrayscale -> <a:grayscl/>
  $bp = $s1.Shapes.AddPicture($png, $msoFalse, $msoTrue, 485, 300, 64, 64); $bp.Name = 'BiLevelPic'
  $bp.PictureFormat.ColorType = 3   # ppPictureBlackAndWhite -> <a:biLevel/>

  # Nested groups: (RectA + RectB) grouped, then that group grouped with RectC.
  $ra = $s1.Shapes.AddShape(1, 560, 300, 60, 40); $ra.Name = 'RectA'
  $rb = $s1.Shapes.AddShape(1, 630, 300, 60, 40); $rb.Name = 'RectB'
  $inner = $s1.Shapes.Range(@('RectA','RectB')).Group(); $inner.Name = 'InnerGroup'
  $rc = $s1.Shapes.AddShape(1, 700, 300, 60, 40); $rc.Name = 'RectC'
  $outer = $s1.Shapes.Range(@('InnerGroup','RectC')).Group(); $outer.Name = 'OuterGroup'
  $outer.Rotation = 8

  # Styled table with NO own cell fill -> cell fill must be resolved from the table style.
  $st = $s1.Shapes.AddTable(3, 3, 40, 410, 380, 90); $st.Name = 'StyledTable'
  $st.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = 'H1'
  $st.Table.Cell(2,1).Shape.TextFrame.TextRange.Text = 'body'
  $st.Table.ApplyStyle($STYLE_GUID, $true)

  # Table with an inline scheme-colour fill + brightness -> <a:schemeClr><a:lumMod/><a:lumOff/>.
  $lt = $s1.Shapes.AddTable(2, 2, 450, 410, 300, 90); $lt.Name = 'LumModTable'
  $c11 = $lt.Table.Cell(1,1).Shape
  $c11.TextFrame.TextRange.Text = 'tinted'
  $c11.Fill.ForeColor.ObjectThemeColor = 8    # msoThemeColorAccent4
  $c11.Fill.ForeColor.Brightness = -0.5       # darker -> lumMod 50%

  # Speaker notes on slide 1 (notes dimension).
  $s1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Speaker note: synthetic stress fixture, slide one.'

  # ============================ SLIDE 2 — master 2 (Facet theme) ============================
  $s2 = $pres.Slides.Add(2, 2)   # ppLayoutText
  $s2.ApplyTheme($thmxB)         # -> second design => second slideMaster + theme
  $s2.Shapes.Title.TextFrame.TextRange.Text = 'Second master, second theme'
  try { $s2.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Master two body placeholder' } catch {}

  $box = $s2.Shapes.AddTextbox(1, 60, 300, 400, 40); $box.Name = 'Headline'
  $box.TextFrame.TextRange.Text = 'Draft headline on master two'

  # Another styled table on master 2 so table-style resolution is exercised against a 2nd theme.
  $st2 = $s2.Shapes.AddTable(2, 3, 60, 360, 400, 80); $st2.Name = 'StyledTable2'
  $st2.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = 'A'
  $st2.Table.ApplyStyle($STYLE_GUID, $true)

  # Threaded modern comments with two authors (2018 comment schema).
  $cm = $s2.Comments.Add2(60, 300, 'Ada Lovelace', 'AL', 'Tighten this headline', 'Windows Live', 'ada@example.com')
  $null = $cm.Replies.Add2(0, 0, 'Grace Hopper', 'GH', 'Agreed, will revise', 'Windows Live', 'grace@example.com')
  $null = $cm.Replies.Add2(0, 0, 'Ada Lovelace', 'AL', 'Done in v2', 'Windows Live', 'ada@example.com')
  $null = $s2.Comments.Add2(470, 300, 'Grace Hopper', 'GH', 'Second standalone note', 'Windows Live', 'grace@example.com')

  if (Test-Path $out) { Remove-Item $out -Force }
  $pres.SaveAs($out, 24, $msoTrue)   # 24=pptx, EmbedTrueTypeFonts=msoTrue
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Host "SAVED: $out"
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
}
