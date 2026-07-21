<#
.SYNOPSIS
  Foreground the running desktop PowerPoint window and take a window-scoped
  screenshot. Use this as the first sanity check for this skill (confirms a
  visible PowerPoint window exists and can be foregrounded from this shell)
  and any time you just want to "look" at current state without driving an
  action.
#>
param([string]$Shot = "$PSScriptRoot\shot.png")
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\ppt-window-lib.ps1"

$hwnd = Set-PptForeground
$size = Save-WindowScreenshot -Hwnd $hwnd -Path $Shot
Write-Host "Foregrounded hwnd=$hwnd; SHOT ($size): $Shot"
