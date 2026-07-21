<#
.SYNOPSIS
  Drive a desktop-PowerPoint ribbon action end to end: foreground the app,
  open a ribbon command via sequential KeyTips, invoke the resulting popup
  menu item via UI Automation, optionally dump/toggle/invoke dialog controls,
  and take a window-scoped screenshot to confirm the result.

.DESCRIPTION
  Generalized from the working procedure used to author the Zoom
  (Slide/Section/Summary Zoom) fixture, whose insertion has no COM/VBA
  surface at all. Reach for this only after confirming the feature truly has
  no COM equivalent (see the powerpoint-fixture-authoring skill first,
  including its ExecuteMso fallback for enum-marshalling failures) - this is
  slower and less deterministic than COM.

.PARAMETER KeyTips
  Semicolon-separated SendKeys tokens sent SEQUENTIALLY (not chorded) to walk
  the ribbon's KeyTip overlay, e.g. '{ESC};{ESC};%;N;Y' = clear stray state,
  then Alt, then N (Insert tab), then Y (a group/button whose KeyTip is Y).
  Discover the letters by screenshotting after each Alt press - KeyTips are
  drawn as small overlay badges on the ribbon.

.PARAMETER UiaMenuItem
  Name of a MenuItem control to Invoke via UI Automation after the KeyTips
  open its popup - use this for submenu entries that keyboard KeyTips do not
  reach (confirmed: Office ribbon split-button dropdowns are on a separate
  input queue from SendKeys).

.PARAMETER DialogTitleLike
  Wildcard pattern to scope the FOLLOW-ON dialog (e.g. 'Insert *Zoom*').
  Required if you pass -Dump, -Toggle, or -InvokeButton, so control lookups
  don't enumerate the whole desktop (privacy + reliability).

.PARAMETER Dump
  Print the real UIA accessible names of every CheckBox/Button/MenuItem/
  ListItem/RadioButton in the dialog scope. Always run this once before
  guessing -Toggle/-InvokeButton names from what's visible on screen.

.PARAMETER Toggle
  Semicolon list of checkbox accessible Names to toggle ON.

.PARAMETER InvokeButton
  Accessible Name of a button to Invoke at the end (e.g. 'Insert', 'OK').

.EXAMPLE
  # Discover Insert-tab KeyTips
  & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%' -Shot shot-keytips.png

.EXAMPLE
  # Open Insert > Zoom, invoke "Slide Zoom", dump the resulting dialog
  & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%;N;Y' -UiaMenuItem 'Slide Zoom' `
      -DialogTitleLike 'Insert *Zoom*' -Dump -Shot shot-dlg.png

.EXAMPLE
  # Toggle a checkbox by its REAL accessible name (found via -Dump above) and submit
  & drive-ribbon.ps1 -DialogTitleLike 'Insert *Zoom*' -Toggle 'Slide 2 Alpha 1' `
      -InvokeButton 'Insert' -Shot shot-done.png
#>
param(
  [string]$KeyTips = "",
  [int]$KeyTipGapMs = 700,
  [string]$UiaMenuItem = "",
  [string]$DialogTitleLike = "",
  [switch]$Dump,
  [string]$Toggle = "",
  [string]$InvokeButton = "",
  [int]$SettleMs = 1300,
  [string]$Shot = "$PSScriptRoot\shot-drive-ribbon.png"
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\ppt-window-lib.ps1"

$hwnd = Set-PptForeground
Write-Host "Foregrounded PowerPoint hwnd=$hwnd"

if ($KeyTips) {
  Send-KeyTipSequence -Keys ($KeyTips -split ';') -GapMs $KeyTipGapMs
}

if ($UiaMenuItem -or $Dump -or $Toggle -or $InvokeButton) {
  . "$PSScriptRoot\uia-lib.ps1"

  if ($UiaMenuItem) {
    Start-Sleep -Milliseconds 500
    Invoke-UiaElement -Name $UiaMenuItem -ControlType $script:CT::MenuItem | Out-Null
    Start-Sleep -Milliseconds 1500
  }

  $scope = $null
  if ($DialogTitleLike) {
    $dlg = Find-UiaDialog -TitleLike $DialogTitleLike
    if ($dlg) { Write-Host "DIALOG SCOPE: '$($dlg.Current.Name)'"; $scope = $dlg }
    else { Write-Host "DIALOG SCOPE: '$DialogTitleLike' NOT FOUND - falling back to full desktop (unscoped)" }
  }

  if ($Dump) { Get-UiaControlDump -Scope $scope }

  if ($Toggle) {
    foreach ($nm in ($Toggle -split ';')) {
      if ($nm.Trim()) { Set-UiaToggleOn -Name $nm.Trim() -Scope $scope | Out-Null; Start-Sleep -Milliseconds 250 }
    }
  }

  if ($InvokeButton) {
    Invoke-UiaElement -Name $InvokeButton -ControlType $script:CT::Button -Scope $scope | Out-Null
  }
}

Start-Sleep -Milliseconds $SettleMs
$size = Save-WindowScreenshot -Hwnd $hwnd -Path $Shot
Write-Host "SHOT ($size): $Shot"
