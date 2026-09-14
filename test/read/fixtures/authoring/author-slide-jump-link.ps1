param(
  # The deck before the deletion. Defaults to the fixture beside `authoring/`.
  [string]$Before = (Join-Path $PSScriptRoot '..\slide-jump-link.pptx'),
  # The same deck after PowerPoint deletes the slide the links jump to.
  [string]$After = (Join-Path $PSScriptRoot '..\slide-jump-link-target-deleted.pptx')
)
# slide-jump-link.pptx and slide-jump-link-target-deleted.pptx: what PowerPoint does to the links,
# sections and custom show that name a slide when that slide is deleted.
#
# The before deck:
#   slide 1  Target     the slide everything below names
#   slide 2  Referrer   LinkedText, a run that jumps to Target
#                       LinkedShape, a rectangle whose click action jumps to Target
#                       ControlText, a run that jumps to slide 3, which is not deleted
#   slide 3  Other
#   sections "First" (slide 1) and "Rest" (slides 2 and 3)
#   custom show "Tour" listing slides 1 and 2
#
# The after deck is the before deck with slide 1 deleted through Slide.Delete() and saved. A test
# runs removeSlide(0) on the before deck and compares what it leaves with what PowerPoint left.
$ErrorActionPreference = 'Stop'
$Before = [IO.Path]::GetFullPath($Before)
$After = [IO.Path]::GetFullPath($After)
foreach ($f in @($Before, $After)) { if (Test-Path $f) { Remove-Item $f -Force } }

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

# A text box on $slide named $name holding $text, returned.
function Add-Label($slide, [string]$name, [string]$text, [int]$top) {
  $box = $slide.Shapes.AddTextbox(1, 60, $top, 600, 50)
  $box.Name = $name
  $box.TextFrame.TextRange.Text = $text
  return $box
}

# Make $actionSettings jump to $target: ppActionHyperlink with the "SlideID,SlideIndex,Title"
# sub-address PowerPoint writes for a slide in the same deck.
function Set-SlideJump($actionSettings, $target, [string]$title) {
  $actionSettings.Action = 7                  # ppActionHyperlink
  $actionSettings.Hyperlink.SubAddress = ('{0},{1},{2}' -f $target.SlideID, $target.SlideIndex, $title)
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(0)

  $target = $pres.Slides.Add(1, 12)           # ppLayoutBlank
  [void](Add-Label $target 'TargetLabel' 'Target' 60)
  $referrer = $pres.Slides.Add(2, 12)
  $other = $pres.Slides.Add(3, 12)
  [void](Add-Label $other 'OtherLabel' 'Other' 60)

  $text = Add-Label $referrer 'LinkedText' 'Go to Target' 60
  Set-SlideJump $text.TextFrame.TextRange.ActionSettings.Item(1) $target 'Target'

  $shape = $referrer.Shapes.AddShape(1, 60, 160, 240, 80)   # msoShapeRectangle
  $shape.Name = 'LinkedShape'
  Set-SlideJump $shape.ActionSettings.Item(1) $target 'Target'

  $control = Add-Label $referrer 'ControlText' 'Go to Other' 300
  Set-SlideJump $control.TextFrame.TextRange.ActionSettings.Item(1) $other 'Other'

  [void]$pres.SectionProperties.AddBeforeSlide(1, 'First')
  [void]$pres.SectionProperties.AddBeforeSlide(2, 'Rest')
  [void]$pres.SlideShowSettings.NamedSlideShows.Add('Tour', [int[]]@($target.SlideID, $referrer.SlideID))

  $pres.SaveAs($Before)
  $pres.Slides.Item(1).Delete()
  $pres.SaveAs($After)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null

  # Reopen the after deck and ask PowerPoint where each surviving action on the referrer points.
  $pres = $pp.Presentations.Open($After, -1, 0, 0)
  Write-Output ('SECTIONS' + "`t" + $pres.SectionProperties.Count)
  $shows = $pres.SlideShowSettings.NamedSlideShows
  Write-Output ('SHOWS' + "`t" + $shows.Count + "`t" + $(if ($shows.Count -gt 0) { $shows.Item(1).Count } else { '' }))
  foreach ($s in $pres.Slides.Item(1).Shapes) {
    $action = if ($s.HasTextFrame) { $s.TextFrame.TextRange.ActionSettings.Item(1) } else { $s.ActionSettings.Item(1) }
    Write-Output ('ACTION' + "`t" + $s.Name + "`t" + $action.Action + "`t" + $action.Hyperlink.SubAddress)
  }
  $pres.Close()
  $pres = $null
  $pp.Quit()
  Write-Output ('SAVED: ' + $Before)
  Write-Output ('SAVED: ' + $After)
}
finally {
  if ($pres -ne $null) { try { $pres.Close() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
