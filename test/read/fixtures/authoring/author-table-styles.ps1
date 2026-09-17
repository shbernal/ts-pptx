# Authors test/read/fixtures/table-styles.pptx — a brand-free, PowerPoint-authored
# deck whose ppt/tableStyles.xml carries REAL style definitions, for testing the
# importSlideMasters table-style merge.
#
# Applying a built-in table style makes PowerPoint materialise that style's full
# definition into ppt/tableStyles.xml (verified: 4 defs, ~9.4 KB). Only Microsoft
# built-in style GUIDs are used, so the fixture carries no brand content.
#
# NOTE: the tblStyleLst@def (default table style) CANNOT be set via COM — PowerPoint
# exposes no equivalent of Word's SetDefaultTableStyle / Excel's DefaultTableStyle
# (Table has only ApplyStyle + a read-only Style). It used to be a one-click manual
# step, which meant re-running this recipe silently reset the default to Office's own
# Accent 1 and the fixture stopped being what the import tests read. The @def is now
# rewritten in the saved package below, so the recipe reproduces the fixture whole.
$ErrorActionPreference = 'Stop'

# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$out = (Join-Path $FIX 'table-styles.pptx')

# Built-in table styles (Microsoft's own GUIDs — not brand assets).
$STYLES = @(
    @{ Name = 'tbl-medium2-accent3'; Guid = '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}'; Label = 'Medium Style 2 - Accent 3' },
    @{ Name = 'tbl-medium4-accent4'; Guid = '{C4B1156A-380E-4F78-BDF5-A606A8083BF9}'; Label = 'Medium Style 4 - Accent 4' },
    @{ Name = 'tbl-light2-accent1';  Guid = '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}'; Label = 'Light Style 2 - Accent 1' }
)

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.DisplayAlerts = 1
    $pres = $pp.Presentations.Add(1)
    # 960x540pt = 12192000x6858000 EMU — matches empty.pptx, so importSlideMasters'
    # equal-size guard passes without an override.
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank
    $top = 40
    foreach ($s in $STYLES) {
        $shape = $slide.Shapes.AddTable(3, 3, 60, $top, 480, 120)
        $shape.Name = $s.Name
        $tbl = $shape.Table
        $tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = $s.Label
        $tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = 'body'
        $tbl.ApplyStyle($s.Guid, $true)
        Write-Output ("  {0,-22} <- {1}  {2}" -f $s.Name, $s.Guid, $s.Label)
        $top += 150
    }

    # A fourth table with every Table Style Options checkbox ticked, so `a:tblPr` carries all six
    # region flags. The other three tables give only firstRow and bandRow, which is what the read
    # model happened to expose; the four that were unreadable had no fixture to be read from.
    $shape = $slide.Shapes.AddTable(4, 4, 60, $top, 480, 140)
    $shape.Name = 'tbl-all-look-flags'
    $tbl = $shape.Table
    $tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'all six regions'
    $tbl.ApplyStyle('{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}', $true)
    $tbl.FirstRow = $true       # a:tblPr/@firstRow
    $tbl.LastRow = $true        # a:tblPr/@lastRow
    $tbl.FirstCol = $true       # a:tblPr/@firstCol
    $tbl.LastCol = $true        # a:tblPr/@lastCol
    $tbl.HorizBanding = $true   # a:tblPr/@bandRow
    $tbl.VertBanding = $true    # a:tblPr/@bandCol
    Write-Output "  tbl-all-look-flags     <- every Table Style Options checkbox"

    if (Test-Path $out) { Remove-Item $out -Force }
    $pres.SaveAs($out)
    $pres.Saved = $true
    $pres.Close()
    $pp.Quit()
    $pp = $null

    # --- set tblStyleLst@def, which COM cannot ---------------------------------
    # A non-Office default is the point of the fixture: `importSlideMasters({ tableStyles })`
    # has to carry the SOURCE deck's default across, and a fixture defaulting to Office's own
    # Accent 1 cannot tell a carried-over default from an untouched one. The attribute names a
    # style this part already defines, so nothing here invents a reference.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($out, 'Update')
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'ppt/tableStyles.xml' }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $xml = [regex]::Replace($xml, 'def="\{[0-9A-Fa-f-]+\}"', ('def="' + $STYLES[0].Guid + '"'))
        $stream = $entry.Open()
        $stream.SetLength(0)
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        $writer.Write($xml)
        $writer.Flush()
        $writer.Close()
        Write-Output ("  tblStyleLst@def        <- {0}" -f $STYLES[0].Guid)
    } finally { $zip.Dispose() }
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
Write-Output "`nSaved: $out"
