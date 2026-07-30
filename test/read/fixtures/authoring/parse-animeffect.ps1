$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
$ASSETS  = Join-Path $PSScriptRoot 'assets'
$pptx = Join-Path $SCRATCH 'animeffect-probe.pptx'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($pptx)
try {
  $entry = $zip.Entries | Where-Object { $_.FullName -eq 'ppt/slides/slide1.xml' }
  $sr = New-Object System.IO.StreamReader($entry.Open())
  $xmlText = $sr.ReadToEnd(); $sr.Close()
} finally { $zip.Dispose() }

$doc = New-Object System.Xml.XmlDocument
$doc.LoadXml($xmlText)
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main')
$ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')

# spid -> shape name
$spidToName = @{}
foreach ($cNvPr in $doc.SelectNodes('//p:sp/p:nvSpPr/p:cNvPr', $ns)) {
  $spidToName[$cNvPr.GetAttribute('id')] = $cNvPr.GetAttribute('name')
}

$rows = @()
foreach ($cTn in $doc.SelectNodes('//p:cTn[@presetID]', $ns)) {
  $spTgt = $cTn.SelectSingleNode('.//p:tgtEl/p:spTgt', $ns)
  $spid = if ($spTgt) { $spTgt.GetAttribute('spid') } else { '' }
  # Collect distinct behavior-node element local-names under this effect's childTnLst
  $behaviors = @()
  $childLst = $cTn.SelectSingleNode('p:childTnLst', $ns)
  if ($childLst) {
    foreach ($n in $childLst.ChildNodes) {
      if ($n.LocalName -ne 'par' -and $n.LocalName -ne 'set') { $behaviors += $n.LocalName }
      elseif ($n.LocalName -eq 'set') { $behaviors += 'set' }
    }
  }
  $rows += [pscustomobject]@{
    name        = $spidToName[$spid]
    spid        = $spid
    presetID    = $cTn.GetAttribute('presetID')
    presetClass = $cTn.GetAttribute('presetClass')
    presetSub   = $cTn.GetAttribute('presetSubtype')
    nodeType    = $cTn.GetAttribute('nodeType')
    behaviors   = ($behaviors -join ',')
  }
}
# name is effN where N is the MsoAnimEffect id
$rows = $rows | Sort-Object { [int]($_.name -replace 'eff','') }
$rows | Where-Object { $_.presetClass -eq 'emph' } | ForEach-Object {
  "{0,-8} id={1,-4} presetID={2,-4} sub={3,-3} behaviors=[{4}]" -f $_.name, ($_.name -replace 'eff',''), $_.presetID, $_.presetSub, $_.behaviors
}
Write-Host "---ATTRNAMES for emph ids 57,59,61,63---"
foreach ($wantId in 57,59,61,63) {
  $spid = ($spidToName.GetEnumerator() | Where-Object { $_.Value -eq "eff$wantId" }).Key
  $cTn = $doc.SelectNodes('//p:cTn[@presetID]', $ns) | Where-Object {
    $t = $_.SelectSingleNode('.//p:tgtEl/p:spTgt', $ns); $t -and $t.GetAttribute('spid') -eq $spid
  } | Select-Object -First 1
  $attrs = @()
  foreach ($an in $cTn.SelectNodes('.//p:attrName', $ns)) { $attrs += $an.InnerText }
  "eff$wantId presetID=$($cTn.GetAttribute('presetID')): attrNames=[$($attrs -join '; ')]"
}
Write-Host "---TOTAL $($rows.Count) effects---"
