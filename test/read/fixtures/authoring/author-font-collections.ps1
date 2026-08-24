$ErrorActionPreference = 'Stop'
# Authors ../fonts/windows-collections.oracle.json: an independent reading of the
# advance widths and cmap coverage inside the TrueType collections (.ttc) that ship
# with Windows.
#
# Why an oracle at all. src/measure/font-collection.ts unwraps one member of a .ttc
# into a standalone sfnt by rewriting the table directory in place, on the strength of
# a claim about the format (a member's table records carry offsets absolute to the
# file, which is how members share one glyf). If that claim were wrong the unwrapper
# would still produce a parseable font, just one whose advances silently came from the
# wrong bytes. Nothing self-generated can catch that, so the numbers are taken from a
# DIFFERENT implementation: System.Windows.Media.GlyphTypeface, the WPF font stack,
# which reads a collection natively through a `file:///...ttc#index` URI and needs none
# of this repo's code to do it.
#
# The sidecar records, per collection member: the Win32 family name, and the advance
# width (in em, the unit GlyphTypeface reports) of a fixed set of code points chosen to
# separate members that would otherwise look alike - Latin for the proportional vs
# monospaced split (MS Gothic advances 'A' at 0.5 em, MS PGothic at 0.6328), Han and
# Kana for the CJK body, a fullwidth form, a math operator that only Cambria Math
# carries, and a Hangul syllable.
#
# A code point the member's cmap does not cover is OMITTED rather than recorded as
# zero, so the sidecar states coverage as well as width: the test asserts that
# everything listed here is covered, which would fail if the unwrapper handed back a
# different member's cmap.
#
# Consumed by test/regression/text/font-collection.test.js, which skips when the fonts
# are absent (any non-Windows machine, and CI). Re-run after a Windows font update if
# the test starts reporting a family this file does not list.

$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$outFile = Join-Path $REPO 'test\read\fixtures\fonts\windows-collections.oracle.json'

Add-Type -AssemblyName PresentationCore

# U+0041 'A', U+0061 'a', U+0030 '0', U+0020 space, U+002E '.', U+004D 'M', U+0069 'i'
# U+65E5 日, U+672C 本, U+8A9E 語, U+306E の, U+30C6 テ, U+FF21 fullwidth A
# U+2211 n-ary summation (Cambria Math, not Cambria), U+AC00 가 (Hangul)
$codePoints = @(0x41, 0x61, 0x30, 0x20, 0x2E, 0x4D, 0x69, 0x65E5, 0x672C, 0x8A9E, 0x306E, 0x30C6, 0xFF21, 0x2211, 0xAC00)

$faces = @()
foreach ($file in (Get-ChildItem (Join-Path $env:SystemRoot 'Fonts\*.ttc') | Sort-Object Name)) {
	$index = 0
	while ($true) {
		# GlyphTypeface throws once `index` is past the last member: that is how the
		# member count is discovered without parsing the ttcf header here.
		try {
			$uri = [Uri]("file:///" + $file.FullName.Replace('\', '/') + "#" + $index)
			$face = New-Object System.Windows.Media.GlyphTypeface($uri)
		} catch {
			break
		}
		$map = $face.CharacterToGlyphMap
		$advances = [ordered]@{}
		foreach ($cp in $codePoints) {
			if ($map.ContainsKey($cp)) {
				$advances[('U+{0:X4}' -f $cp)] = [math]::Round($face.AdvanceWidths[$map[$cp]], 6)
			}
		}
		$faces += [pscustomobject]@{
			file     = $file.Name
			index    = $index
			family   = $face.Win32FamilyNames.Values[0]
			advances = $advances
		}
		$index++
	}
}

[pscustomobject]@{
	source    = 'System.Windows.Media.GlyphTypeface (WPF), authored by test/read/fixtures/authoring/author-font-collections.ps1'
	unit      = 'advance widths are fractions of an em, as GlyphTypeface reports them'
	windows   = [System.Environment]::OSVersion.Version.ToString()
	faceCount = $faces.Count
	faces     = $faces
} | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $outFile

Write-Host ("wrote {0} faces from {1} collections to {2}" -f $faces.Count, ($faces | Select-Object -ExpandProperty file -Unique).Count, $outFile)
