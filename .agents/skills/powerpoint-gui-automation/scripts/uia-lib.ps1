# Dot-source this file: . '...\uia-lib.ps1'
# UI Automation helpers for controls that keyboard SendKeys and synthetic
# mouse (SetCursorPos + mouse_event) BOTH fail to activate - verified
# 2026-07-21 on Office ribbon dropdown popups and their follow-on dialogs,
# which appear to run on a separate input queue (UIPI-isolated from the
# calling process). InvokePattern/TogglePattern operate on the accessibility
# tree directly and bypass the input queue entirely, so they work where both
# keyboard and mouse input were silently swallowed.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
# NOTE: loading these assemblies flips the process DPI-aware, which can jump
# screenshot resolution mid-session (e.g. 1280x800 -> 1920x1200). Harmless for
# UIA itself (it's coordinate-free), but take your "before" screenshot before
# dot-sourcing this file if you need a stable-resolution comparison.

$script:AE = [System.Windows.Automation.AutomationElement]
$script:TS = [System.Windows.Automation.TreeScope]
$script:CT = [System.Windows.Automation.ControlType]

function Find-UiaDialog {
  # Scope subsequent searches to ONE dialog window instead of the whole
  # desktop. Unscoped RootElement.FindAll(Descendants) enumerates every open
  # window on the machine, including unrelated apps (this can and did
  # incidentally capture content from other open windows) - always scope by
  # dialog title before dumping/invoking.
  param([Parameter(Mandatory)][string]$TitleLike)
  $wcond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)
  foreach ($w in $AE::RootElement.FindAll($TS::Children, $wcond)) {
    if ($w.Current.Name -like $TitleLike) { return $w }
  }
  foreach ($w in $AE::RootElement.FindAll($TS::Descendants, $wcond)) {
    if ($w.Current.Name -like $TitleLike) { return $w }
  }
  return $null
}

function Find-UiaElementByName {
  param([Parameter(Mandatory)][string]$Name, $ControlType = $null, $Scope = $null)
  if (-not $Scope) { $Scope = $AE::RootElement }
  $conds = @(New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name))
  if ($ControlType) { $conds += New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ControlType) }
  $cond = if ($conds.Count -gt 1) { New-Object System.Windows.Automation.AndCondition($conds) } else { $conds[0] }
  return $Scope.FindFirst($TS::Descendants, $cond)
}

function Get-UiaControlDump {
  # Print the REAL accessible Name of every interactive control in scope.
  # Always run this before guessing a control's name/label from what's
  # visible on screen: displayed text and the accessible Name frequently
  # differ (e.g. a checkbox showing "2. Alpha 1" on screen was actually
  # named "Slide 2 Alpha 1" in the accessibility tree - guessing the visible
  # label found nothing and Invoke silently no-op'd with 0 selections).
  param($Scope = $null, [ControlType[]]$ControlTypes = @($script:CT::CheckBox, $script:CT::Button, $script:CT::MenuItem, $script:CT::ListItem, $script:CT::RadioButton))
  if (-not $Scope) { $Scope = $AE::RootElement }
  foreach ($ct in $ControlTypes) {
    $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ct)
    foreach ($e in $Scope.FindAll($TS::Descendants, $c)) {
      if ($e.Current.Name) { Write-Host ("{0}: '{1}'" -f $ct.ProgrammaticName, $e.Current.Name) }
    }
  }
}

function Invoke-UiaElement {
  # Works for menu items, buttons, split-button dropdown entries - anything
  # exposing InvokePattern - regardless of whether keyboard/mouse reach it.
  param([Parameter(Mandatory)][string]$Name, $ControlType = $null, $Scope = $null)
  $el = Find-UiaElementByName -Name $Name -ControlType $ControlType -Scope $Scope
  if ($null -eq $el) { Write-Host "UIA INVOKE: '$Name' NOT FOUND"; return $false }
  $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
  Write-Host "UIA INVOKE: '$Name' OK"
  return $true
}

function Set-UiaToggleOn {
  param([Parameter(Mandatory)][string]$Name, $Scope = $null)
  $el = Find-UiaElementByName -Name $Name -ControlType $script:CT::CheckBox -Scope $Scope
  if ($null -eq $el) { Write-Host "UIA TOGGLE: '$Name' NOT FOUND"; return $false }
  $tp = $el.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
  if ($tp.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) { $tp.Toggle() }
  Write-Host "UIA TOGGLE: '$Name' -> $($tp.Current.ToggleState)"
  return $true
}
