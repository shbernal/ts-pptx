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
# (Table has only ApplyStyle + a read-only Style). Setting a non-Accent-1 default is
# therefore a one-click manual step; see test/read/fixtures/README.md.
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

    if (Test-Path $out) { Remove-Item $out -Force }
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
Write-Output "`nSaved: $out"
