# Readiness guard for the autofit-calibration fixture matrix.
# Certifies (1) the five fonts actually resolve (GDI), not just that PowerPoint
# would write the requested typeface, and (2) LibreOffice soffice is reachable.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$fonts = 'Aptos','Aptos SemiBold','Calibri','Tahoma','Arial'
$bad = @()
Write-Output "=== Font presence (GDI resolution) ==="
foreach ($name in $fonts) {
  $f = New-Object System.Drawing.Font($name, 18)
  $resolved = $f.Name; $f.Dispose()
  $ok = ($resolved -eq $name)
  if (-not $ok) { $bad += "$name -> $resolved" }
  Write-Output ("  {0,-16} -> {1,-22} {2}" -f $name, $resolved, ($(if($ok){'OK'}else{'*** SUBSTITUTED ***'})))
}

Write-Output "=== LibreOffice ==="
$soffice = (Get-Command soffice -ErrorAction SilentlyContinue).Source
if (-not $soffice) {
  $cand = "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe",
          'C:\Program Files\LibreOffice\program\soffice.exe',
          'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
  $soffice = (Get-ChildItem $cand -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName)
}
if ($soffice) { Write-Output "  soffice: $soffice" } else { Write-Output "  soffice: *** NOT FOUND ***" }

Write-Output "=== VERDICT ==="
if ($bad.Count -eq 0 -and $soffice) {
  Write-Output "READY: all fonts resolve and LibreOffice is present."
} else {
  if ($bad.Count) { Write-Output ("NOT READY - substituted fonts: " + ($bad -join '; ')) }
  if (-not $soffice) { Write-Output "NOT READY - LibreOffice missing." }
}
