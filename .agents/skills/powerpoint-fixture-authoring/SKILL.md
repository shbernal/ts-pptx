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
4. Author the deck with desktop PowerPoint COM. Write the script to a temp file
   (e.g. `.tmp/author-<name>.ps1`) and run it through the **PowerShell (pwsh 7)
   tool** with the call operator: `& '.tmp/author-<name>.ps1'`. PowerShell 7
   drives PowerPoint COM fine here. Do **not** invoke
   `powershell.exe -NoProfile -ExecutionPolicy Bypass ...`: the `-ExecutionPolicy
   Bypass` flag trips the sandbox's "Security Weaken" classifier and the call is
   denied. Plain `& '<script>.ps1'` runs under the session's existing policy and
   needs no bypass.
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
   - inspect the relevant slide XML for the OOXML construct being pinned
     (use `scripts/dump-slide-xml.ps1`, below — don't re-author a throwaway
     dump script);
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
# Snapshot pre-existing PIDs so the reap at the end only kills the server we
# spawn — never a user's interactive PowerPoint with unsaved work.
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
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
  # Quit() can leave the automation server lingering; reap only PIDs we created.
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
```

Do not chase a lingering `POWERPNT` by killing a hard-coded PID — the
snapshot/reap above (and the same logic in the verify helper) handles it
deterministically without touching a user's open PowerPoint.

For grouped shape fixtures, create the child shapes first, group by shape names,
then set transforms on the returned group. PowerPoint writes group transforms
under `p:grpSpPr/a:xfrm`, for example `rot`, `flipH`, and `flipV`.

### Freeform / boolean (`custGeom`) geometry

Build a single freeform with `BuildFreeform`/`AddNodes`/`ConvertToShape` (each
`AddNodes` appends one segment; `msoSegmentLine`=0, `msoSegmentCurve`=1 takes
control1/control2/end). Return-to-start closes the path → `a:close`.

For a hole or a boolean combination, use PowerPoint's **Merge Shapes** rather
than the COM `ShapeRange.MergeShapes(...)` method — that method's enum argument
fails under PowerShell COM late-binding (both 5.1 and 7): the direct call throws
`Exception setting "MergeShapes": Cannot convert ... to Object` and a reflection
`InvokeMember` throws `DISP_E_TYPEMISMATCH`. Drive the ribbon command instead:
open the deck **with a window** (`Presentations.Open(path, msoFalse, msoFalse,
msoTrue)`), select the shapes, and call `ExecuteMso`, which takes a plain string
and needs no enum marshalling:

```powershell
$base.Select($true)     # msoTrue: replace selection
$hole.Select($false)    # msoFalse: extend selection
$pp.CommandBars.ExecuteMso('ShapesSubtract')   # or ShapesUnion / ShapesCombine / ShapesFragment
$merged = $pp.ActiveWindow.Selection.ShapeRange.Item(1)
$merged.Name = '...'    # name the merged result
```

(This same `ExecuteMso`-on-a-selection pattern is the fallback for any COM method
whose enum/optional args refuse to late-bind in PowerShell.)

**`custGeom` output gotchas — verified 2026-06-21, note in the README so the next
author doesn't re-discover them:**

- `BuildFreeform` emits exactly one `a:path` per freeform.
- Merge Shapes (Union / Combine / Subtract, **even of disjoint shapes**)
  consolidates everything into a **single** `a:path` with multiple
  `moveTo`…`close` contours — desktop PowerPoint never writes more than one
  `a:path` per `custGeom`. A genuine multi-`a:path` `a:pathLst` is therefore **not
  authorable** via any built-in PowerPoint operation; it is schema-legal but only
  arises from other producers (e.g. SVG import) and remains unverified here.

## Autofit bake-on-save (`normAutofit` / `spAutoFit`)

PowerPoint bakes autofit results into the saved XML **non-interactively** (no
manual editing in the UI) — but only with the right COM sequence. **Verified
2026-06-21** while authoring the `autofit-*` calibration decks:

- **Pin the box first, set `AutoSize` last.** `AddTextbox`'s height argument is
  ignored, so explicitly: `TextFrame2.AutoSize = msoAutoSizeNone`, set
  `Shape.Width`/`Shape.Height`, add the text, **then** set `AutoSize`. With the
  box pinned small at that moment:
  - `msoAutoSizeTextToFitShape` (shrink) bakes
    `<a:normAutofit fontScale="…" lnSpcReduction="…"/>` and keeps `ext` pinned;
  - `msoAutoSizeShapeToFitText` (resize) bakes `<a:spAutoFit/>` and a fitted
    `ext.cy`.
- **The trigger is the box being pinned small when `AutoSize` is applied — not
  "before vs after text" per se.** An earlier theory ("set AutoSize before
  text") was wrong: a box that has already grown tall bakes only a bare
  `<a:normAutofit/>` with no scale. Pin → text → AutoSize-last is what bakes a
  real `fontScale`.
- **Two-pass text build.** A trailing empty paragraph is **not** enumerable via
  `TextRange.Paragraphs()` until the whole text exists — `Paragraphs($i+1,1)`
  throws an "index out of bounds" COM error. So: pass 1 inserts every
  paragraph's text and formats non-empty runs inline; pass 2 (once
  `Paragraphs().Count` is final) does paragraph-level formatting and the run
  font for empty paragraphs. See `.tmp/author-deck.ps1` from that session for a
  complete parameterized engine.

## Font-presence guard (substitution is invisible in the XML)

**PowerPoint writes `latin@typeface="X"` into the run XML even when font `X` is
not installed** — it substitutes only at render time. So asserting the typeface
in the saved OOXML does **not** prove the font was actually used; a fixture can
silently carry the wrong metrics. Verify **host-side via GDI** instead — the
resolved face must equal the requested name:

```powershell
Add-Type -AssemblyName System.Drawing
foreach ($face in 'Aptos','Aptos SemiBold','Calibri','Tahoma','Arial') {
  $f = New-Object System.Drawing.Font($face, 18); $resolved = $f.Name; $f.Dispose()
  if ($resolved -ne $face) { throw "FONT SUBSTITUTED: $face -> $resolved" }
}
```

Run this as a hard precondition before authoring any font-sensitive fixture
(and re-run in a **fresh** process after installing a font — a prior process's
"ready" can't be trusted). `.tmp/readiness-guard.ps1` and the guard block in
`.tmp/author-deck.ps1` are the worked examples.

### Provisioning fonts / LibreOffice without elevation

Both are installable with **no admin** when a fixture needs them — verified
2026-06-21:

- **Per-user font install (no elevation):** copy the `.ttf`s into
  `%LOCALAPPDATA%\Microsoft\Windows\Fonts` and register each under
  `HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`; GDI resolves them
  in a fresh process. Office **cloud fonts** (e.g. Aptos) are already on disk in
  `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\<Family>\` but with numeric
  names and no extension and **only visible to Office apps**, not GDI — picking
  the font in the PowerPoint dropdown is *not* enough. Identify each cached file
  by its OpenType `name` table (family/face), then per-user install the faces
  you need.
- **LibreOffice without elevation:** winget's package is a machine-scoped MSI
  that triggers UAC. Instead do an **administrative extract** into a user dir:
  `msiexec /a <msi> /qn TARGETDIR=%LOCALAPPDATA%\Programs\LibreOffice`. That
  yields a runnable `program\soffice.exe` (confirmed headless conversion works)
  with no admin. Verify the MSI's SHA-256 against winget's published hash first.

LibreOffice is useful here as an independent cross-measure: its `program\
python.exe` + `pyuno` can open a deck (LibreOffice recomputes `spAutoFit` on
load) and read each shape's fitted size via UNO — a second opinion on
PowerPoint's baked metrics. See `.tmp/measure-lo.py` for the UNO bootstrap.

## Helpers

Both scripts run through the PowerShell (pwsh 7) tool with the call operator —
no `-ExecutionPolicy Bypass`.

**Verify** — opens the deck through PowerPoint COM, reads package metadata,
computes SHA-256, optionally lists group `a:xfrm` attributes, and reaps any
automation-server process it spawned (reported as `reapedProcessIds`):

```powershell
& '.agents\skills\powerpoint-fixture-authoring\scripts\verify-powerpoint-fixture.ps1' -Path 'test\read\fixtures\<fixture>.pptx' -InspectGroups
```

**Dump slide XML** — prints a slide's XML so you can confirm the exact OOXML
construct PowerPoint emitted (indented by default; `-Raw` for the stored form):

```powershell
& '.agents\skills\powerpoint-fixture-authoring\scripts\dump-slide-xml.ps1' -Path 'test\read\fixtures\<fixture>.pptx' -Slide 1
```

If the helper output does not show the expected OOXML construct, fix the
PowerPoint authoring script and regenerate the fixture rather than patching the
OOXML by hand.
