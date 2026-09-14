param(
  # Where the pasted decks, their PNGs and the read-back land. Not a fixture.
  [string]$OutDir = (Join-Path $PSScriptRoot '..\..\..\..\.tmp\default-text-style-paste')
)
# Probe; produces no committed fixture. What PowerPoint does to text that inherits from the source
# deck's p:defaultTextStyle when its slide is pasted into a deck whose default text style differs.
#
# Source: ../default-text-style.pptx, whose default text style is 18pt, tx1 and the minor font.
# PlainBox's run inherits all three from it; StyledRect's inherits its size and takes its colour
# from a p:style fontRef. Destination: ../table-text-inheritance.pptx, whose default text style is
# 28pt, C00000 and Courier New.
#   dest.pptx  Slides.Paste, "Use Destination Theme"
#   keep.pptx  ExecuteMso PasteSourceFormatting, "Keep Source Formatting"
# Each pasted slide's two runs are read back over COM and the slide is exported to PNG.
#
# Measured 2026-09-14: with the destination theme both runs re-resolve to 28pt, and PlainBox turns
# Courier New C00000. Keeping source formatting, both stay 18pt Aptos, PlainBox black and StyledRect
# white, and PowerPoint gets there by writing the resolved values onto each run: sz="1800", the
# colour, and the minor font as a literal a:latin, besides copying the source master. That is the
# evidence for the preserve import baking what a run takes from the source p:defaultTextStyle.
$ErrorActionPreference = 'Stop'
$OutDir = [IO.Path]::GetFullPath($OutDir)
$fixtures = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
New-Item -ItemType Directory -Force $OutDir | Out-Null
foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1                            # the source-formatting paste is a ribbon command
  $src = $pp.Presentations.Open((Join-Path $fixtures 'default-text-style.pptx'), -1, 0, 0)

  foreach ($mode in 'dest', 'keep') {
    $dst = $pp.Presentations.Open((Join-Path $fixtures 'table-text-inheritance.pptx'), -1, 0, -1)
    $src.Slides.Item(1).Copy()
    if ($mode -eq 'dest') {
      [void]$dst.Slides.Paste(2)
    } else {
      $win = $dst.Windows.Item(1)
      $win.Activate()
      $win.ViewType = 7                       # ppViewSlideSorter, so the paste lands after slide 1
      $dst.Slides.Item(1).Select()
      $pp.CommandBars.ExecuteMso('PasteSourceFormatting')
      for ($i = 0; $i -lt 40 -and $dst.Slides.Count -lt 2; $i++) { Start-Sleep -Milliseconds 250 }
    }
    if ($dst.Slides.Count -lt 2) { throw "the $mode paste added no slide" }
    $dst.SaveAs((Join-Path $OutDir "$mode.pptx"))
    $pasted = $dst.Slides.Item(2)
    foreach ($name in 'PlainBox', 'StyledRect') {
      $font = $pasted.Shapes.Item($name).TextFrame.TextRange.Runs(1).Font
      Write-Output ('FONT' + "`t" + $mode + "`t" + $name + "`t" + $font.Size + "`t" + $font.Name + "`t" + ('{0:X6}' -f $font.Color.RGB))
    }
    $pasted.Export((Join-Path $OutDir "$mode.png"), 'PNG', 1280, 720)
    $dst.Saved = $true
    $dst.Close()
  }
  $src.Close()
  $pp.Quit()
}
finally {
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
