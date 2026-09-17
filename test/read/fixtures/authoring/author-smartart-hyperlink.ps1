# Authors ../smartart-hyperlink.pptx — one SmartArt diagram whose first node's text carries a
# hyperlink, so the drawing cache carries one too.
#
# A SmartArt diagram stores its text twice: the authored nodes in `ppt/diagrams/data{N}.xml`, and
# a copy of every drawn string in `ppt/diagrams/drawing{N}.xml`. A link on a node's run therefore
# exists twice as well, and the two parts hold their own relationships — so `@r:id` in the drawing
# names a relationship of the DRAWING part, not of the data part. Reading the cached text without
# them reported the raw id and a null url, where the same run read through the point's own
# `textFrame` resolved. This fixture is the deck where those two readings can be compared.
#
# `TextRange2` has no hyperlink surface — `SmartArt.Nodes` reaches a node's `TextFrame2` and
# nothing on it sets a link — so this follows the same inject-then-reopen sequence
# `author-smartart-families.ps1` uses for its arrow labels: author with PowerPoint, inject the
# `a:hlinkClick` and its relationship into the data part, then REOPEN in PowerPoint and SaveAs the
# fixture path. The committed bytes are PowerPoint's own serialization, and the round trip is what
# proves the link is real: PowerPoint rebuilds the drawing cache from the data part on open, so a
# link that reaches `drawing1.xml` is one PowerPoint itself put there.
#
# Read .agents/skills/powerpoint-fixture-authoring/SKILL.md first — teardown/reap discipline and
# the Resiliency precondition live there.
#
#   & test\read\fixtures\authoring\author-smartart-hyperlink.ps1

$ErrorActionPreference = 'Stop'

$FIX = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$out = Join-Path $FIX 'smartart-hyperlink.pptx'
$tmp = Join-Path $FIX 'smartart-hyperlink.base.pptx'
foreach ($f in @($out, $tmp)) { if (Test-Path $f) { Remove-Item $f -Force } }

foreach ($sub in 'DocumentRecovery', 'StartupItems') {
	$key = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$sub"
	if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

$CYCLE = 'urn:microsoft.com/office/officeart/2005/8/layout/cycle2'
$LINK = 'https://example.invalid/smartart-node'
$LINKED_NODE = 'linked-node'
$PLAIN_NODE = 'plain-node'

function Get-Layout($app, $id) {
	$layouts = $app.SmartArtLayouts
	for ($i = 1; $i -le $layouts.Count; $i++) {
		if ($layouts.Item($i).Id -eq $id) { return $layouts.Item($i) }
	}
	throw "no installed SmartArt layout with id $id"
}

function Read-Entry($zip, $name) {
	$entry = $zip.GetEntry($name)
	if ($null -eq $entry) { throw "no such part: $name" }
	$reader = New-Object System.IO.StreamReader($entry.Open())
	try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

function Write-Entry($zip, $name, $text) {
	$zip.GetEntry($name).Delete()
	$writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open())
	try { $writer.Write($text) } finally { $writer.Close() }
}

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
	$pp = New-Object -ComObject PowerPoint.Application
	$pp.DisplayAlerts = 1

	# --- author: one cycle diagram, two named nodes ------------------------------------
	# cycle2 is the simplest family that maps one node to one drawn shape, so the two readings
	# being compared differ only in which part they came from.
	$pres = $pp.Presentations.Add(1)
	$slide = $pres.Slides.Add(1, 12)   # ppLayoutBlank
	$shape = $slide.Shapes.AddSmartArt((Get-Layout $pp $CYCLE), 40, 40, 880, 440)
	$shape.Name = 'linked-smartart'
	$nodes = $shape.SmartArt.Nodes
	$nodes.Item(1).TextFrame2.TextRange.Text = $LINKED_NODE
	$nodes.Item(2).TextFrame2.TextRange.Text = $PLAIN_NODE

	$pres.SaveAs($tmp)
	$pres.Saved = $true
	$pres.Close()
	$pres = $null

	# --- inject: a hyperlink on the first node's run, plus its relationship -------------
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	$dataPart = $null
	$zip = [System.IO.Compression.ZipFile]::Open($tmp, 'Update')
	try {
		$slideXml = Read-Entry $zip 'ppt/slides/slide1.xml'
		$relsXml = Read-Entry $zip 'ppt/slides/_rels/slide1.xml.rels'
		$dmId = [regex]::Match($slideXml, '<dgm:relIds[^>]*\sr:dm="([^"]+)"').Groups[1].Value
		if (-not $dmId) { throw 'slide 1 has no dgm:relIds/@r:dm' }
		$target = [regex]::Match($relsXml, "<Relationship Id=`"$dmId`"[^>]*Target=`"([^`"]+)`"").Groups[1].Value
		$dataPart = 'ppt/' + ($target -replace '^\.\./', '')
		$leaf = $dataPart.Substring($dataPart.LastIndexOf('/') + 1)
		$dataRels = $dataPart.Substring(0, $dataPart.LastIndexOf('/')) + '/_rels/' + $leaf + '.rels'

		# A freshly authored data part references nothing, so it has no relationships part at all.
		# `.rels` is covered by the package's `Default Extension="rels"`, so creating one declares
		# itself; only the relationship inside it is new.
		$relId = 'rId1'
		$rel = '<Relationship Id="' + $relId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + $LINK + '" TargetMode="External"/>'
		if ($null -eq $zip.GetEntry($dataRels)) {
			$writer = New-Object System.IO.StreamWriter($zip.CreateEntry($dataRels).Open())
			try {
				$writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					$rel + '</Relationships>')
			}
			finally { $writer.Close() }
			Write-Output ("created {0}" -f $dataRels)
		}
		else {
			$dataRelsXml = Read-Entry $zip $dataRels
			$used = [regex]::Matches($dataRelsXml, 'Id="rId(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value }
			$relId = 'rId' + (1 + ($used | Measure-Object -Maximum).Maximum)
			$rel = $rel -replace 'Id="rId1"', ('Id="' + $relId + '"')
			Write-Entry $zip $dataRels ($dataRelsXml -replace '</Relationships>', ($rel + '</Relationships>'))
		}

		# `a:hlinkClick` is the last child of `a:rPr` in CT_TextCharacterProperties, and the run
		# PowerPoint wrote carries a self-closing `a:rPr`, so it has to be opened up.
		$dataXml = Read-Entry $zip $dataPart
		$run = [regex]::Match($dataXml, '<a:r>\s*<a:rPr\b[^>]*/>\s*<a:t>' + [regex]::Escape($LINKED_NODE) + '</a:t>\s*</a:r>')
		if (-not $run.Success) { throw "no plain run holding $LINKED_NODE in $dataPart" }
		$rPr = [regex]::Match($run.Value, '<a:rPr\b[^>]*/>').Value
		$opened = ($rPr -replace '/>$', '>') + '<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="' + $relId + '"/></a:rPr>'
		Write-Entry $zip $dataPart $dataXml.Replace($run.Value, $run.Value.Replace($rPr, $opened))
		Write-Output ("injected {0} -> {1} on '{2}' in {3}" -f $relId, $LINK, $LINKED_NODE, $dataPart)
	}
	finally { $zip.Dispose() }

	# --- hand it back to PowerPoint: it rebuilds the drawing cache from the data --------
	$pres = $pp.Presentations.Open($tmp, 0, 0, 0)   # throws on a repair prompt
	$pres.SaveAs($out)
	$pres.Saved = $true
	$pres.Close()
	$pres = $null
	$pp.Quit()
	Remove-Item $tmp -Force

	# --- findings: what PowerPoint kept, not what was asked for ------------------------
	$zip = [System.IO.Compression.ZipFile]::OpenRead($out)
	try {
		foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match '^ppt/diagrams/(data|drawing)\d+\.xml$' } | Sort-Object FullName)) {
			$xml = Read-Entry $zip $entry.FullName
			$ids = [regex]::Matches($xml, '<a:hlinkClick[^>]*r:id="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
			Write-Output ('{0,-30} hlinkClick r:id: {1}' -f $entry.FullName, $(if ($ids) { $ids -join ',' } else { '(none)' }))
		}
		foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match '^ppt/diagrams/_rels/' } | Sort-Object FullName)) {
			$xml = Read-Entry $zip $entry.FullName
			foreach ($m in [regex]::Matches($xml, '<Relationship[^>]*hyperlink[^>]*>')) {
				Write-Output ('{0,-30} {1}' -f $entry.FullName, $m.Value)
			}
		}
	}
	finally { $zip.Dispose() }

	Write-Output ('SAVED: {0} ({1} bytes)' -f $out, (Get-Item $out).Length)
}
finally {
	if ($pres) { try { $pres.Close() } catch { } }
	if ($pp) { try { $pp.Quit() } catch { } }
	Start-Sleep -Milliseconds 800
	Get-Process POWERPNT -ErrorAction SilentlyContinue |
		Where-Object { $preexistingIds -notcontains $_.Id } |
		ForEach-Object { try { $_.Kill() } catch { } }
}
