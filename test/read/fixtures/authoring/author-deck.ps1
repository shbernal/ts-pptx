param(
  [Parameter(Mandatory=$true)][string]$CasesPath,
  [Parameter(Mandatory=$true)][string]$OutPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# ---- enums ----
$msoTextOrientationHorizontal = 1
$ppLayoutBlank = 12
$msoTrue = -1; $msoFalse = 0
$msoAutoSizeNone = 0; $msoAutoSizeShapeToFitText = 1; $msoAutoSizeTextToFitShape = 2
$msoAnchorTop = 1; $msoAnchorMiddle = 3; $msoAnchorBottom = 4
$alignMap = @{ 'l' = 1; 'ctr' = 2; 'r' = 3; 'just' = 4 }
$anchorMap = @{ 't' = $msoAnchorTop; 'ctr' = $msoAnchorMiddle; 'b' = $msoAnchorBottom }

$spec = Get-Content -Raw -Path $CasesPath | ConvertFrom-Json

# ---- hard font-presence guard (corrected substitution guard) ----
$bad = @()
foreach ($face in $spec.fontsRequired) {
  $f = New-Object System.Drawing.Font($face, 18); $resolved = $f.Name; $f.Dispose()
  if ($resolved -ne $face) { $bad += "$face -> $resolved" }
}
if ($bad.Count) { throw "FONT SUBSTITUTION DETECTED (aborting author): $($bad -join '; ')" }

function To-PPText([string]$s) {
  # \n in case text -> intra-paragraph line break (vertical tab -> <a:br/>)
  return ($s -replace "`n", [string][char]11)
}

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null; $pres = $null
if (Test-Path $OutPath) { Remove-Item $OutPath -Force }

try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth  = $spec.slideWidthPt
  $pres.PageSetup.SlideHeight = $spec.slideHeightPt

  $slideNums = @($spec.cases | ForEach-Object { [int]$_.slide } | Sort-Object -Unique)
  $slides = @{}
  foreach ($n in $slideNums) { $slides[$n] = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank) }

  foreach ($c in $spec.cases) {
    $slide = $slides[[int]$c.slide]
    $w = [double]$c.wPt; $h = [double]$c.hPt
    $s = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, [double]$c.xPt, [double]$c.yPt, $w, $h)
    $s.Name = $c.id
    $tf = $s.TextFrame2
    $tf.WordWrap = ($(if ($c.wrap) { $msoTrue } else { $msoFalse }))
    $tf.AutoSize = $msoAutoSizeNone
    $s.Width = $w; $s.Height = $h
    if ($c.insetsPt) {
      $tf.MarginLeft = [double]$c.insetsPt.l; $tf.MarginRight = [double]$c.insetsPt.r
      $tf.MarginTop  = [double]$c.insetsPt.t; $tf.MarginBottom = [double]$c.insetsPt.b
    }
    if ($c.anchor -and $anchorMap.ContainsKey($c.anchor)) { $tf.VerticalAnchor = $anchorMap[$c.anchor] }

    $tr = $tf.TextRange
    # Pass 1: insert all text; format non-empty runs inline. (A trailing empty
    # paragraph is not enumerable via .Paragraphs() until the whole text exists,
    # so empty-paragraph formatting is deferred to pass 2.)
    $pi = 0
    foreach ($para in $c.paragraphs) {
      if ($pi -gt 0) { [void]$tr.InsertAfter([string][char]13) }  # new paragraph
      foreach ($run in $para.runs) {
        $ptxt = To-PPText $run.text
        if ($ptxt.Length -eq 0) { continue }
        $rng = $tr.InsertAfter($ptxt)
        $rng.Font.Name = $run.font
        $rng.Font.Size = [double]$run.sizePt
        $rng.Font.Bold = ($(if ($run.bold) { $msoTrue } else { $msoFalse }))
        $rng.Font.Italic = ($(if ($run.italic) { $msoTrue } else { $msoFalse }))
        if ($run.PSObject.Properties.Name -contains 'charSpacingPts' -and $run.charSpacingPts -ne $null) {
          try { $rng.Font.Spacing = [double]$run.charSpacingPts } catch { Write-Output "WARN charSpacing unsupported on $($c.id): $_" }
        }
      }
      $pi++
    }
    # Pass 2: paragraph-level formatting + empty-paragraph run font (all paragraphs now exist)
    $pcount = $tr.Paragraphs().Count
    $pi = 0
    foreach ($para in $c.paragraphs) {
      if ($pi + 1 -gt $pcount) { $pi++; continue }
      $prng = $tr.Paragraphs($pi + 1, 1)
      $firstRun = $para.runs[0]
      if ((To-PPText $firstRun.text).Length -eq 0) {
        $prng.Font.Name = $firstRun.font; $prng.Font.Size = [double]$firstRun.sizePt
      }
      $pf = $prng.ParagraphFormat
      if ($para.PSObject.Properties.Name -contains 'align' -and $para.align -and $alignMap.ContainsKey($para.align)) { $pf.Alignment = $alignMap[$para.align] }
      if ($para.PSObject.Properties.Name -contains 'spaceBeforePts' -and $para.spaceBeforePts -ne $null) { $pf.LineRuleBefore = $msoFalse; $pf.SpaceBefore = [double]$para.spaceBeforePts }
      if ($para.PSObject.Properties.Name -contains 'spaceAfterPts' -and $para.spaceAfterPts -ne $null) { $pf.LineRuleAfter = $msoFalse; $pf.SpaceAfter = [double]$para.spaceAfterPts }
      if ($para.PSObject.Properties.Name -contains 'lineSpacingPct' -and $para.lineSpacingPct -ne $null) { $pf.LineRuleWithin = $msoTrue; $pf.SpaceWithin = [double]$para.lineSpacingPct / 100.0 }
      elseif ($para.PSObject.Properties.Name -contains 'lineSpacingPts' -and $para.lineSpacingPts -ne $null) { $pf.LineRuleWithin = $msoFalse; $pf.SpaceWithin = [double]$para.lineSpacingPts }
      $pi++
    }

    # AutoSize applied AFTER text with the box already pinned (None + W/H above):
    #  - shrink: box stays pinned, PowerPoint bakes <a:normAutofit fontScale lnSpcReduction>
    #  - resize: PowerPoint writes <a:spAutoFit> + fitted ext.cy
    if ($c.kind -eq 'shrink') { $tf.AutoSize = $msoAutoSizeTextToFitShape }
    elseif ($c.kind -eq 'resize') { $tf.AutoSize = $msoAutoSizeShapeToFitText }
  }

  $pres.SaveAs($OutPath)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output "AUTHORED $($spec.cases.Count) cases across $($slideNums.Count) slides -> $OutPath"
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
}
