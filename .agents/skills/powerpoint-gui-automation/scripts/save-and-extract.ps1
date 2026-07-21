<#
.SYNOPSIS
  Save the presentation you've been GUI-driving (via its already-running COM
  instance) and extract its .pptx package to a fresh directory, so you can
  read the actual generated OOXML rather than trusting the screenshot alone.
  A screenshot proves the UI reacted; it does not prove the XML is correct
  or wired up (relationships, ids) - always close the loop by reading the
  parts this script extracts.

.PARAMETER NameLike
  Wildcard to find the right open Presentations item by its .Name (useful
  when more than one presentation is open).

.PARAMETER DestDir
  Extraction target. If it already exists, a fresh "<DestDir>-<random>" is
  used instead of deleting anything - do not delete-and-recreate: this
  sandbox's Bash/PowerShell guard can false-positive-block Remove-Item when
  the surrounding command text also contains regex-like substrings (e.g. a
  literal "r:" from an XML namespace or "\w+"), so a fresh directory per run
  is more reliable than reuse.
#>
param(
  [Parameter(Mandatory)][string]$NameLike,
  [Parameter(Mandatory)][string]$DestDir
)
$ErrorActionPreference = 'Stop'

$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
$pres = $null
for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
  if ($ppt.Presentations.Item($i).Name -like $NameLike) { $pres = $ppt.Presentations.Item($i); break }
}
if ($null -eq $pres) { throw "No open presentation matching Name -like '$NameLike'" }

$pptxPath = $pres.FullName
$pres.Save()
Write-Host "SAVED: $pptxPath (Slides=$($pres.Slides.Count))"

$zip = "$DestDir.zip"
Copy-Item -Path $pptxPath -Destination $zip -Force
$dir = $DestDir
if (Test-Path $dir) { $dir = "$DestDir-$(Get-Random)" }
Expand-Archive -Path $zip -DestinationPath $dir -Force
Write-Host "EXTRACT DIR: $dir"

Get-ChildItem -Path (Join-Path $dir "ppt\slides") -Filter *.xml -ErrorAction SilentlyContinue |
  ForEach-Object { Write-Host "SLIDE: $($_.Name) ($($_.Length) bytes)" }
Get-ChildItem -Path (Join-Path $dir "ppt\slides\_rels") -Filter *.rels -ErrorAction SilentlyContinue |
  ForEach-Object { Write-Host "RELS: $($_.Name)" }
