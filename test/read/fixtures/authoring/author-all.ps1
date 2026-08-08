$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$TMP = $SCRATCH
# Both live here, beside this recipe. They used to be read from $TMP, which is how the
# generator came to exist twice: the copy under scripts/ was edited, the .tmp copy this
# line actually loaded was whatever had last been staged there, and neither was the file
# a reader would open. Resolve from $PSScriptRoot like every other recipe in this
# directory, so the file that runs is the file that is tracked.
$engine = Join-Path $PSScriptRoot 'author-deck.ps1'
$measure = Join-Path $PSScriptRoot 'measure-lo.py'
$py = 'C:\Users\000023500\AppData\Local\Programs\LibreOffice\program\python.exe'

# deck -> whether to run the LibreOffice cross-measure
$decks = [ordered]@{
  'autofit-line-metrics' = $true
  'autofit-shrink'       = $false
  'autofit-resize'       = $true
  'autofit-edge'         = $true
}
foreach ($name in $decks.Keys) {
  $cases = "$FIX\$name.cases.json"
  $pptx  = "$FIX\$name.pptx"
  Write-Output "==== authoring $name ===="
  & $engine -CasesPath $cases -OutPath $pptx
  if ($decks[$name]) {
    Write-Output "---- LibreOffice measure $name ----"
    & $py $measure $pptx | Set-Content "$TMP\$name.lo.json" -Encoding utf8
    Write-Output "LO json: $((Get-Item "$TMP\$name.lo.json").Length) bytes"
  }
}
Write-Output "ALL DECKS AUTHORED"
