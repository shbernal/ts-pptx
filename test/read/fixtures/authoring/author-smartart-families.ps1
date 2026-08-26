# Authors ../smartart-families.pptx — four SmartArt diagrams, one per layout family, each
# node carrying a distinct greppable string so the node -> drawn-shape mapping is falsifiable.
#
# The fixture exists for the drawing cache. A SmartArt diagram stores its text twice: the
# authored nodes in `ppt/diagrams/data{N}.xml`, and a copy of every drawn string in
# `ppt/diagrams/drawing{N}.xml`, which PowerPoint recomputes on open and every other renderer
# draws verbatim. Anything that re-texts a diagram has to update both, and the four families
# here are the ones that break a naive "one node, one shape" mapping in different ways:
#
#   1 orgChart1  multi-level nesting plus an `asst` node — and every node is presented twice,
#                once as a text box and once as the connector below it.
#   2 process1   arrows whose `sibTrans` points carry labels, which no other family has.
#   3 cycle2     non-linear order, one node per shape.
#   4 pList1     picture nodes, whose `pictRect` shape has a fill and no text body at all.
#
# Arrow labels have **no COM surface** — `SmartArt.Nodes` reaches nodes and assistants only,
# and text set on a node never lands on a transition point. So slide 2 follows the
# `author-table-cell-horzoverflow.ps1` sequence: author with PowerPoint -> inject the labels
# into the data part -> REOPEN in PowerPoint and SaveAs the fixture path. The committed bytes
# are PowerPoint's own serialization, and the round-trip is what proves the labels are real:
# PowerPoint strips transition text a layout has no place for, and the findings block below
# prints exactly which of the injected labels survived.
#
# Read .agents/skills/powerpoint-fixture-authoring/SKILL.md first — teardown/reap discipline
# and the Resiliency precondition live there.
#
#   & test\read\fixtures\authoring\author-smartart-families.ps1

$ErrorActionPreference = 'Stop'

$FIX = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$out = Join-Path $FIX 'smartart-families.pptx'
$tmp = Join-Path $FIX 'smartart-families.base.pptx'
foreach ($f in @($out, $tmp)) { if (Test-Path $f) { Remove-Item $f -Force } }

# Document Recovery state left by a force-killed POWERPNT makes the next launch modal, which
# surfaces as RPC_E_CALL_REJECTED on the first automation call. Clear it before every run.
foreach ($sub in 'DocumentRecovery', 'StartupItems') {
	$key = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$sub"
	if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

$ORG_CHART = 'urn:microsoft.com/office/officeart/2005/8/layout/orgChart1'
$PROCESS = 'urn:microsoft.com/office/officeart/2005/8/layout/process1'
$CYCLE = 'urn:microsoft.com/office/officeart/2005/8/layout/cycle2'
$PICTURE_LIST = 'urn:microsoft.com/office/officeart/2005/8/layout/pList1'

# `Application.SmartArtLayouts` is empty until a presentation exists — the collection is
# populated from the open document, not from the installed layout set.
function Get-Layout($app, $id) {
	$layouts = $app.SmartArtLayouts
	for ($i = 1; $i -le $layouts.Count; $i++) {
		if ($layouts.Item($i).Id -eq $id) { return $layouts.Item($i) }
	}
	throw "no installed SmartArt layout with id $id"
}

# `SmartArtNodes.Add()` appends to whichever collection it is called on, so a node's own
# `.Nodes.Add()` is how a child level is reached. There is no COM route to a transition point.
function Set-Nodes($nodes, $labels) {
	for ($i = 1; $i -le $nodes.Count; $i++) {
		$nodes.Item($i).TextFrame2.TextRange.Text = $labels[$i - 1]
	}
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

	# --- author: four diagrams, default 16:9 (960x540pt), blank layouts ----------------
	$pres = $pp.Presentations.Add(1)

	# Slide 1 — org chart. The default is a root with one assistant and three reports; a
	# grandchild under the first report is what makes the nesting more than two deep.
	$slide = $pres.Slides.Add(1, 12) # ppLayoutBlank
	$shape = $slide.Shapes.AddSmartArt((Get-Layout $pp $ORG_CHART), 40, 40, 880, 440)
	$shape.Name = 'OrgChart'
	$root = $shape.SmartArt.Nodes.Item(1)
	$root.TextFrame2.TextRange.Text = 'org-root'
	Set-Nodes $root.Nodes @('org-asst', 'org-child-1', 'org-child-2', 'org-child-3')
	$grandchild = $root.Nodes.Item(2).Nodes.Add()
	$grandchild.TextFrame2.TextRange.Text = 'org-grandchild'
	Write-Output ('slide 1 orgChart1  root children={0} asst.Type={1} grandchild.Level={2}' -f
		$root.Nodes.Count, $root.Nodes.Item(1).Type, $grandchild.Level)

	# Slide 2 — labelled process. Four nodes so there are three arrows between them, plus the
	# trailing `sibTrans` after the last node, which has no arrow to label.
	$slide = $pres.Slides.Add(2, 12)
	$shape = $slide.Shapes.AddSmartArt((Get-Layout $pp $PROCESS), 40, 120, 880, 280)
	$shape.Name = 'LabeledProcess'
	[void]$shape.SmartArt.Nodes.Add()
	Set-Nodes $shape.SmartArt.Nodes @('proc-1', 'proc-2', 'proc-3', 'proc-4')
	Write-Output ('slide 2 process1   nodes={0}' -f $shape.SmartArt.Nodes.Count)

	# Slide 3 — cycle. One node per drawn shape, and the ring order is not document order.
	$slide = $pres.Slides.Add(3, 12)
	$shape = $slide.Shapes.AddSmartArt((Get-Layout $pp $CYCLE), 40, 40, 880, 440)
	$shape.Name = 'Cycle'
	Set-Nodes $shape.SmartArt.Nodes @('cycle-1', 'cycle-2', 'cycle-3', 'cycle-4', 'cycle-5')
	Write-Output ('slide 3 cycle2     nodes={0}' -f $shape.SmartArt.Nodes.Count)

	# Slide 4 — picture list. Each node draws two shapes; the picture one stays empty on
	# purpose, so the fixture carries no image and the shape is a fill with no text body.
	$slide = $pres.Slides.Add(4, 12)
	$shape = $slide.Shapes.AddSmartArt((Get-Layout $pp $PICTURE_LIST), 40, 40, 880, 440)
	$shape.Name = 'PictureList'
	Set-Nodes $shape.SmartArt.Nodes @('pic-1', 'pic-2', 'pic-3', 'pic-4')
	Write-Output ('slide 4 pList1     nodes={0} shapes/node={1}' -f
		$shape.SmartArt.Nodes.Count, $shape.SmartArt.Nodes.Item(1).Shapes.Count)

	$pres.SaveAs($tmp)
	$pres.Saved = $true
	$pres.Close()
	$pres = $null

	# --- inject: arrow labels into slide 2's data part --------------------------------
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	$injected = 0
	$dataPart = $null
	$zip = [System.IO.Compression.ZipFile]::Open($tmp, 'Update')
	try {
		# `dgm:relIds/@r:dm` on the graphic frame names the data part, resolved against the
		# slide's own relationships.
		$slideXml = Read-Entry $zip 'ppt/slides/slide2.xml'
		$relsXml = Read-Entry $zip 'ppt/slides/_rels/slide2.xml.rels'
		$dmId = [regex]::Match($slideXml, '<dgm:relIds[^>]*\sr:dm="([^"]+)"').Groups[1].Value
		if (-not $dmId) { throw 'slide 2 has no dgm:relIds/@r:dm' }
		$target = [regex]::Match($relsXml, "<Relationship Id=`"$dmId`"[^>]*Target=`"([^`"]+)`"").Groups[1].Value
		$dataPart = 'ppt/' + ($target -replace '^\.\./', '')

		$dataXml = Read-Entry $zip $dataPart
		$dataXml = [regex]::Replace($dataXml, '<dgm:pt\b[^>]*>.*?</dgm:pt>', {
				param($m)
				$pt = $m.Value
				$type = [regex]::Match($pt, '^<dgm:pt\b[^>]*\stype="([^"]+)"').Groups[1].Value
				if ($type -ne 'parTrans' -and $type -ne 'sibTrans') { return $pt }
				$script:injected++
				$label = '{0}-{1}' -f $type, $script:injected
				# CT_Pt orders prSet, spPr, t — a transition point PowerPoint has never drawn text
				# into carries no `dgm:t` at all, so one has to be appended rather than replaced.
				$body = "<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=`"en-US`"/><a:t>$label</a:t></a:r></a:p></dgm:t>"
				$existing = [regex]::Match($pt, '<dgm:t>.*?</dgm:t>', 'Singleline')
				if ($existing.Success) { return $pt.Replace($existing.Value, $body) }
				return $pt -replace '</dgm:pt>$', "$body</dgm:pt>"
			}, 'Singleline')
		Write-Entry $zip $dataPart $dataXml
	}
	finally { $zip.Dispose() }
	Write-Output ('injected {0} transition labels into {1}' -f $injected, $dataPart)

	# --- hand it back to PowerPoint: it regenerates the drawing cache from the data ----
	$pres = $pp.Presentations.Open($tmp, 0, 0, 0)   # throws on a repair prompt
	$pres.SaveAs($out)
	$pres.Saved = $true
	$pres.Close()
	$pres = $null
	$pp.Quit()
	Remove-Item $tmp -Force

	# --- findings: what PowerPoint resolved, not what was asked for -------------------
	$zip = [System.IO.Compression.ZipFile]::OpenRead($out)
	try {
		foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match '^ppt/diagrams/data\d+\.xml$' } | Sort-Object FullName)) {
			$xml = Read-Entry $zip $entry.FullName
			$layout = [regex]::Match($xml, 'loTypeId="([^"]+)"').Groups[1].Value -replace '.*/', ''
			$counts = @{}
			foreach ($m in [regex]::Matches($xml, '<dgm:pt\b[^>]*>')) {
				$t = [regex]::Match($m.Value, '\stype="([^"]+)"').Groups[1].Value
				if (-not $t) { $t = 'node' }
				$counts[$t] = 1 + $counts[$t]
			}
			$kept = [regex]::Matches($xml, '<a:t>((?:par|sib)Trans-\d+)</a:t>') | ForEach-Object { $_.Groups[1].Value }
			Write-Output ('{0,-24} {1,-12} pts {2}  transition labels kept: {3}' -f
				$entry.FullName, $layout,
				(($counts.Keys | Sort-Object | ForEach-Object { "$_=$($counts[$_])" }) -join ' '),
				$(if ($kept) { $kept -join ',' } else { '(none)' }))
		}
	}
	finally { $zip.Dispose() }

	Write-Output ('SAVED: {0} ({1} bytes)' -f $out, (Get-Item $out).Length)
	Write-Output ('SHA256: {0}' -f (Get-FileHash $out -Algorithm SHA256).Hash.ToLower())
}
finally {
	if ($null -ne $pres) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
	if ($null -ne $pp) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
	Get-Process POWERPNT -ErrorAction SilentlyContinue |
		Where-Object { $preexistingIds -notcontains $_.Id } |
		Stop-Process -Force -ErrorAction SilentlyContinue
}
