param(
  [Parameter(Mandatory = $true)]
  [string]$Path,

  # 1-based slide number to dump (ppt/slides/slideN.xml).
  [int]$Slide = 1,

  # Emit the raw single-line XML exactly as stored, instead of indented.
  [switch]$Raw
)

$ErrorActionPreference = 'Stop'

$resolved = (Resolve-Path -LiteralPath $Path).Path
Add-Type -AssemblyName System.IO.Compression.FileSystem

$entryName = "ppt/slides/slide$Slide.xml"
$zip = [System.IO.Compression.ZipFile]::OpenRead($resolved)
try {
  $entry = $zip.GetEntry($entryName)
  if ($null -eq $entry) {
    $available = ($zip.Entries |
      Where-Object { $_.FullName -match '^ppt/slides/slide[0-9]+\.xml$' } |
      ForEach-Object { $_.FullName }) -join ', '
    throw "Entry '$entryName' not found. Available slides: $available"
  }

  $stream = $entry.Open()
  try {
    $reader = [System.IO.StreamReader]::new($stream)
    try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
  }
  finally {
    $stream.Dispose()
  }
}
finally {
  $zip.Dispose()
}

if ($Raw) {
  $xml
}
else {
  # Indent for readable inspection of the OOXML construct being pinned.
  $doc = [xml]$xml
  $sw = [System.IO.StringWriter]::new()
  $settings = [System.Xml.XmlWriterSettings]::new()
  $settings.Indent = $true
  $settings.IndentChars = '  '
  $writer = [System.Xml.XmlWriter]::Create($sw, $settings)
  try {
    $doc.Save($writer)
  }
  finally {
    $writer.Dispose()
  }
  $sw.ToString()
}
