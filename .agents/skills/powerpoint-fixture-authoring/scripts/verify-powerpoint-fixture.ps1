param(
  [Parameter(Mandatory = $true)]
  [string]$Path,

  [switch]$InspectGroups
)

$ErrorActionPreference = 'Stop'

$resolved = (Resolve-Path -LiteralPath $Path).Path
$pp = $null
$pres = $null
$openInfo = $null

# Snapshot PIDs that already exist so we only ever reap the automation server
# we spawn below — never a user's interactive PowerPoint with unsaved work.
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Open($resolved, -1, 0, 0)
  $openInfo = [ordered]@{
    opened = $true
    slides = $pres.Slides.Count
    name = $pres.Name
    saved = [bool]$pres.Saved
    powerPointVersion = $pp.Version
  }
  $pres.Close()
  $pp.Quit()
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

# COM Quit() can leave the automation server lingering. Reap only PIDs that
# appeared during this run, so callers never have to kill a process by hand.
# Wait for exit so the powerPointProcesses snapshot below is accurate.
$reapedProcessIds = @(
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    ForEach-Object {
      $proc = $_
      Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
      try { $proc.WaitForExit(5000) | Out-Null } catch {}
      $proc.Id
    }
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryText {
  param(
    [Parameter(Mandatory = $true)]
    [System.IO.Compression.ZipArchive]$Zip,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $entry = $Zip.GetEntry($Name)
  if ($null -eq $entry) { return $null }

  $stream = $entry.Open()
  try {
    $reader = [System.IO.StreamReader]::new($stream)
    try {
      return $reader.ReadToEnd()
    }
    finally {
      $reader.Dispose()
    }
  }
  finally {
    $stream.Dispose()
  }
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($resolved)
try {
  $appXml = Read-ZipEntryText -Zip $zip -Name 'docProps/app.xml'
  $application = if ($appXml -match '<Application>(.*?)</Application>') { $Matches[1] } else { $null }
  $appVersion = if ($appXml -match '<AppVersion>(.*?)</AppVersion>') { $Matches[1] } else { $null }

  $groups = @()
  if ($InspectGroups) {
    $slideEntries = $zip.Entries |
      Where-Object { $_.FullName -match '^ppt/slides/slide[0-9]+\.xml$' } |
      Sort-Object FullName

    foreach ($entry in $slideEntries) {
      $xml = Read-ZipEntryText -Zip $zip -Name $entry.FullName
      $matches = [regex]::Matches(
        $xml,
        '<p:grpSp>[\s\S]*?<p:nvGrpSpPr>[\s\S]*?<p:cNvPr id="([^"]+)" name="([^"]+)"[\s\S]*?<p:grpSpPr>[\s\S]*?<a:xfrm([^>]*)>',
        [System.Text.RegularExpressions.RegexOptions]::Singleline
      )
      foreach ($match in $matches) {
        $groups += [ordered]@{
          slide = $entry.FullName
          id = $match.Groups[1].Value
          name = $match.Groups[2].Value
          xfrm = $match.Groups[3].Value.Trim()
        }
      }
    }
  }
}
finally {
  $zip.Dispose()
}

$hash = Get-FileHash -LiteralPath $resolved -Algorithm SHA256
$powerPointProcesses = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, MainWindowTitle)

[ordered]@{
  path = $resolved
  sha256 = $hash.Hash.ToLowerInvariant()
  application = $application
  appVersion = $appVersion
  open = $openInfo
  groups = $groups
  reapedProcessIds = $reapedProcessIds
  powerPointProcesses = $powerPointProcesses
} | ConvertTo-Json -Depth 6
