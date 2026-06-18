---
name: powerpoint-fixture-authoring
description: Use when creating, replacing, verifying, or documenting real Microsoft PowerPoint-authored .pptx fixtures in this PptxGenJS repository, especially for read-model or OOXML bugs that need desktop PowerPoint output rather than PptxGenJS-generated packages.
---

# PowerPoint Fixture Authoring

Use this skill to create reference `.pptx` fixtures authored by desktop
Microsoft PowerPoint on Windows. These fixtures are evidence for how PowerPoint
writes OOXML; do not generate them with PptxGenJS.

## Workflow

1. Work from the repo root and inspect `git status --short`.
2. Locate an existing fixture with `Get-ChildItem -Recurse -Filter '<name>.pptx'`
   before replacing it. If the user asks to replace it, delete only that exact
   path.
3. Put curated fixtures in `test/read/fixtures/` unless the user specifies
   another target. Use `pptx-bank/` only for uncommitted exploration corpus
   files.
4. Author the deck with desktop PowerPoint COM through Windows PowerShell 5.1:
   `powershell.exe -NoProfile -ExecutionPolicy Bypass ...`. PowerShell 7 can
   instantiate COM, but Windows PowerShell 5.1 has been steadier for clean
   `Close()` / `Quit()` on this machine.
5. Keep the fixture minimal and explicit:
   - set slide size deliberately;
   - name important shapes/groups with stable names;
   - use visible labels only when they help future inspection;
   - avoid external assets unless the bug requires them;
   - prefer deterministic coordinates, colors, rotations, and flips.
6. Save with real PowerPoint using `Presentation.SaveAs()`, set
   `Presentation.Saved = $true`, close the presentation, quit PowerPoint,
   release COM objects, and verify no `POWERPNT` process remains.
7. Verify the saved package:
   - open it once through PowerPoint COM with no repair prompt;
   - inspect `docProps/app.xml` for `Microsoft Office PowerPoint` and
     `AppVersion`;
   - inspect the relevant slide XML for the OOXML construct being pinned;
   - compute SHA-256.
8. Update `test/read/fixtures/README.md` with provenance, hash, purpose, and
   the desktop PowerPoint check date.
9. Commit only the fixture and directly related documentation when asked to
   commit. Leave unrelated dirty state untouched.

## COM Authoring Pattern

Use a temporary script or inline encoded command with this shape:

```powershell
$ErrorActionPreference = 'Stop'
$out = 'C:\path\to\test\read\fixtures\example.pptx'
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

  # Add named shapes/groups that exercise the target PowerPoint behavior.

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
```

For grouped shape fixtures, create the child shapes first, group by shape names,
then set transforms on the returned group. PowerPoint writes group transforms
under `p:grpSpPr/a:xfrm`, for example `rot`, `flipH`, and `flipV`.

## Verification Helper

After creating a fixture, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\powerpoint-fixture-authoring\scripts\verify-powerpoint-fixture.ps1 -Path test\read\fixtures\<fixture>.pptx -InspectGroups
```

The helper opens the deck through PowerPoint COM, reads package metadata,
computes SHA-256, and optionally lists group `a:xfrm` attributes from
`ppt/slides/slide*.xml`.

If the helper output does not show the expected OOXML construct, fix the
PowerPoint authoring script and regenerate the fixture rather than patching the
OOXML by hand.
