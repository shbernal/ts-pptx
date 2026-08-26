# Probe; produces no committed fixture. Answers which SmartArt layouts let a `parTrans` or
# `sibTrans` point carry a label, which is not recorded anywhere in a package and has no COM
# surface to ask.
#
# The route is the one `author-smartart-families.ps1` uses for slide 2, run across a spread of
# layout families at once: PowerPoint authors each diagram, a label is injected into every
# transition point of every data part, and PowerPoint reopens and re-saves the deck. **It
# strips transition text a layout has no place for**, so what survives the round-trip is the
# answer. All output lands in `<repo>/.tmp/`; nothing here is committed.
#
# The findings are recorded in the `smartart-families.pptx` entry of ../README.md. Rerun this
# to re-derive them:
#
#   & test\read\fixtures\authoring\probe-smartart-transition-labels.ps1

$ErrorActionPreference = 'Stop'

$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$TMP = Join-Path $REPO '.tmp'
if (-not (Test-Path $TMP)) { [void](New-Item -ItemType Directory -Path $TMP) }
$base = Join-Path $TMP 'smartart-transition-labels.base.pptx'
$out = Join-Path $TMP 'smartart-transition-labels.pptx'
foreach ($f in @($base, $out)) { if (Test-Path $f) { Remove-Item $f -Force } }

foreach ($sub in 'DocumentRecovery', 'StartupItems') {
	$key = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$sub"
	if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

# One per family whose arrows or connectors plausibly take text, plus the two the fixture
# itself covers (`process1`, `orgChart1`) as controls at each end of the answer.
$layouts = @(
	'process1', 'hProcess3', 'hProcess6', 'chevron1', 'process5', 'bProcess3',
	'arrow1', 'arrow4', 'cycle1', 'cycle2', 'cycle3', 'cycle7',
	'hierarchy6', 'orgChart1', 'hProcess11', 'radial1'
) | ForEach-Object { 'urn:microsoft.com/office/officeart/2005/8/layout/' + $_ }

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

	# --- author: one slide per layout, distinct node text --------------------------------
	$pres = $pp.Presentations.Add(1)
	# `SmartArtLayouts` is empty until a presentation exists, so index it only now.
	$installed = @{}
	for ($i = 1; $i -le $pp.SmartArtLayouts.Count; $i++) { $installed[$pp.SmartArtLayouts.Item($i).Id] = $i }

	$slideNo = 0
	foreach ($id in $layouts) {
		$slideNo++
		if (-not $installed.ContainsKey($id)) { throw "no installed SmartArt layout with id $id" }
		$slide = $pres.Slides.Add($slideNo, 12) # ppLayoutBlank
		$sa = $slide.Shapes.AddSmartArt($pp.SmartArtLayouts.Item($installed[$id]), 40, 40, 880, 440).SmartArt
		for ($n = 1; $n -le $sa.Nodes.Count; $n++) {
			$sa.Nodes.Item($n).TextFrame2.TextRange.Text = 'S{0}N{1}' -f $slideNo, $n
		}
	}
	$pres.SaveAs($base)
	$pres.Saved = $true
	$pres.Close()
	$pres = $null

	# --- inject: a label into every transition point of every data part -------------------
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	$zip = [System.IO.Compression.ZipFile]::Open($base, 'Update')
	try {
		foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -match '^ppt/diagrams/data\d+\.xml$' })) {
			$part = $entry.FullName
			$n = [regex]::Match($part, '\d+').Value
			$script:k = 0
			$xml = [regex]::Replace((Read-Entry $zip $part), '<dgm:pt\b[^>]*>.*?</dgm:pt>', {
					param($m)
					$pt = $m.Value
					$type = [regex]::Match($pt, '^<dgm:pt\b[^>]*\stype="([^"]+)"').Groups[1].Value
					if ($type -ne 'parTrans' -and $type -ne 'sibTrans') { return $pt }
					$script:k++
					$label = 'D{0}{1}{2}' -f $n, $(if ($type -eq 'parTrans') { 'P' } else { 'S' }), $script:k
					$body = "<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=`"en-US`"/><a:t>$label</a:t></a:r></a:p></dgm:t>"
					# A transition PowerPoint has never drawn text into carries no `dgm:t` at all.
					$existing = [regex]::Match($pt, '<dgm:t>.*?</dgm:t>', 'Singleline')
					if ($existing.Success) { return $pt.Replace($existing.Value, $body) }
					return $pt -replace '</dgm:pt>$', "$body</dgm:pt>"
				}, 'Singleline')
			Write-Entry $zip $part $xml
		}
	}
	finally { $zip.Dispose() }

	# --- hand it back: PowerPoint keeps only what its layout can draw ---------------------
	$pres = $pp.Presentations.Open($base, 0, 0, 0)   # throws on a repair prompt
	$pres.SaveAs($out)
	$pres.Saved = $true
	$pres.Close()
	$pres = $null
	$pp.Quit()

	# --- report ---------------------------------------------------------------------------
	$zip = [System.IO.Compression.ZipFile]::OpenRead($out)
	try {
		$rows = foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match '^ppt/diagrams/data\d+\.xml$' })) {
			$xml = Read-Entry $zip $entry.FullName
			$kept = @([regex]::Matches($xml, '<a:t>D\d+([PS])\d+</a:t>') | ForEach-Object { $_.Groups[1].Value })
			[pscustomobject]@{
				Layout   = [regex]::Match($xml, 'loTypeId="([^"]+)"').Groups[1].Value -replace '.*/', ''
				ParTrans = @($kept | Where-Object { $_ -eq 'P' }).Count
				SibTrans = @($kept | Where-Object { $_ -eq 'S' }).Count
				Verdict  = if ($kept.Count -eq 0) { 'drops all transition text' }
				elseif (@($kept | Where-Object { $_ -eq 'P' }).Count -gt 0) { 'keeps parTrans text' }
				else { 'keeps sibTrans text' }
			}
		}
		$rows | Sort-Object Verdict, Layout | Format-Table -AutoSize | Out-String | Write-Output
	}
	finally { $zip.Dispose() }

	Write-Output ('probe deck: {0}' -f $out)
	Write-Output 'A kept count lower than the layout''s transition count is expected: the trailing'
	Write-Output 'sibTrans after the last node labels no arrow, and is dropped with the rest.'
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
