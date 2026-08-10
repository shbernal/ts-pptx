#requires -Version 7
<#
.SYNOPSIS
  Open each .pptx in desktop Microsoft PowerPoint (via COM) and report whether it
  loads cleanly. Catches OOXML corruption the Node test suite cannot see.

.DESCRIPTION
  The project's supported bar (AGENTS.md) is "opens cleanly in Microsoft PowerPoint".
  A structurally invalid package makes PowerPoint raise 0x80070570 (ERROR_FILE_CORRUPT)
  or "PowerPoint could not open the file" on Presentations.Open — neither of which CI,
  which is Node-only, can detect.

  Each open runs in a background job guarded by a timeout so that a modal "repair?"
  dialog blocks only that job instead of hanging the caller. Only PowerPoint processes
  spawned by this run are reaped, so an interactive PowerPoint with unsaved work is
  never touched.

  Exit code is 0 when every deck opens, 1 when any deck fails or hangs — usable as a
  release / pre-commit gate.

.PARAMETER Path
  One or more .pptx paths or globs (e.g. demos/showcases/output/*.pptx).

.PARAMETER TimeoutSec
  Per-deck open timeout. Default 90.

.EXAMPLE
  & .agents/skills/powerpoint-desktop-smoke/scripts/smoke-open.ps1 -Path demos/showcases/output/*.pptx
#>
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string[]]$Path,

  [int]$TimeoutSec = 90
)

$ErrorActionPreference = 'Stop'

# Clear PowerPoint Resiliency so a prior crash's disabled-items / recovery state does
# not mask a fresh repair. Best-effort: absence or a locked key is fine.
try {
  $resil = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency'
  if (Test-Path $resil) { Remove-Item -LiteralPath $resil -Recurse -Force -Confirm:$false -ErrorAction Stop }
} catch {}

# Runs in a child pwsh (Start-Job) so a modal dialog cannot hang the caller.
$openScript = {
  param($deck)
  $resolved = (Resolve-Path -LiteralPath $deck).Path
  # Snapshot pre-existing PIDs so we only ever reap the server we spawn below.
  $preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $pp = $null; $pres = $null
  try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.DisplayAlerts = 1
    # Open(FileName, ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoFalse)
    $pres = $pp.Presentations.Open($resolved, -1, 0, 0)
    $slides = $pres.Slides.Count
    $pres.Close()
    $pp.Quit()
    [pscustomobject]@{ ok = $true; slides = $slides; error = $null }
  }
  catch {
    [pscustomobject]@{ ok = $false; slides = $null; error = $_.Exception.Message }
  }
  finally {
    if ($null -ne $pres) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
    if ($null -ne $pp) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # COM Quit() can leave the automation server lingering; reap only our own PID(s).
    Get-Process POWERPNT -ErrorAction SilentlyContinue |
      Where-Object { $preexisting -notcontains $_.Id } |
      ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
  }
}

$decks = @()
foreach ($p in $Path) {
  # -Path (not -LiteralPath) so globs like output/*.pptx expand.
  $decks += @(Resolve-Path -Path $p -ErrorAction SilentlyContinue | ForEach-Object { $_.Path })
}
if ($decks.Count -eq 0) { Write-Error "No .pptx files matched: $($Path -join ', ')"; exit 2 }

$fail = 0
foreach ($deck in $decks) {
  $name = Split-Path $deck -Leaf
  $job = Start-Job -ScriptBlock $openScript -ArgumentList $deck
  if (Wait-Job $job -Timeout $TimeoutSec) {
    $r = Receive-Job $job
    if ($r.ok) {
      Write-Host ('PASS  {0,-44} slides={1}' -f $name, $r.slides) -ForegroundColor Green
    }
    else {
      Write-Host ('FAIL  {0,-44} {1}' -f $name, $r.error) -ForegroundColor Red
      $fail++
    }
  }
  else {
    Write-Host ('HANG  {0,-44} no response in {1}s (modal repair dialog?)' -f $name, $TimeoutSec) -ForegroundColor Yellow
    Stop-Job $job
    $fail++
  }
  Remove-Job $job -Force -ErrorAction SilentlyContinue
}

if ($fail -gt 0) {
  Write-Host ''
  Write-Host "$fail of $($decks.Count) deck(s) did NOT open cleanly." -ForegroundColor Red
  exit 1
}
Write-Host ''
Write-Host "All $($decks.Count) deck(s) opened cleanly." -ForegroundColor Green
exit 0
