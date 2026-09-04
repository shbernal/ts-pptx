---
name: powerpoint-desktop-smoke
description: Use on Windows with desktop Microsoft PowerPoint installed when you need to confirm that ts-pptx-generated .pptx output actually opens (the project's supported bar), to catch OOXML corruption (0x80070570) the Node test suite cannot see, or to bisect which feature emits a package PowerPoint rejects. Good as a pre-release smoke check after any change to emitted OOXML.
metadata:
  # For working *on* ts-pptx, not *with* it. `npx skills add shbernal/ts-pptx` walks
  # .claude/skills/ (a symlink to this tree) as well as the published skills/, and this flag
  # is what keeps it out of the menu a consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install
  # it anyway.
  internal: true
---

# PowerPoint Desktop Smoke Test

AGENTS.md defines the project's supported bar as **"output opens cleanly in Microsoft
PowerPoint."** CI is Node-only and cannot check that. This machine has desktop
PowerPoint installed and COM-automatable, so opening generated `.pptx` files here is the
only thing that catches structural corruption — duplicate `cNvPr` ids, dangling
relationship/`spid` references, bad content types — that produces valid-looking XML the
test suite passes but PowerPoint refuses.

**Failure signals** from `Presentations.Open`:
- `0x80070570` (ERROR_FILE_CORRUPT) — "The file or directory is corrupted and unreadable."
- "PowerPoint could not open the file."
- A hang (the script reports `HANG`) — a modal "repair?" dialog is blocking.

Use this for the two out-of-suite failure modes; do NOT use it to chase third-party
office-suite interop quirks (WPS round-trips, etc.), which AGENTS.md puts out of scope.

## Prerequisites

- Windows with desktop PowerPoint (`POWERPNT.EXE`, Office16) — verify COM is available:
  `[Type]::GetTypeFromProgID('PowerPoint.Application')` is non-null.
- Run scripts through the **PowerShell (pwsh 7) tool** with the call operator:
  `& '.agents/skills/powerpoint-desktop-smoke/scripts/smoke-open.ps1' ...`. Do not shell
  out to `powershell.exe -ExecutionPolicy Bypass`.

## Workflow

1. **Generate decks.** From the repo root, `pnpm demos:build` (it rebuilds `dist/` first
   only if stale) writes both showcases to `demos/showcases/output/`:
   `Kestrel_Q3_Business_Review.pptx` (charts, tables, groups, masters) and
   `Field_Notes_Four_Cities.pptx` (images, media, a 3D model, picture effects). Between
   them they reach most of the emitter. `pnpm demos:build quarterly-review` builds one.

   The showcases are decks, not a feature matrix — there is no per-feature generator to
   ask for a single construct. For that, write a focused deck (step 3).

2. **Open the decks in PowerPoint.** Point the smoke script at the output:
   ```
   & '.agents/skills/powerpoint-desktop-smoke/scripts/smoke-open.ps1' -Path 'demos/showcases/output/*.pptx'
   ```
   It opens each deck in a timeout-guarded background job (a modal dialog blocks only that
   job, never the session), reaps only the PowerPoint it spawned, prints `PASS`/`FAIL`/`HANG`
   per deck, and exits non-zero if any deck fails. Use it the same way on any single deck.

3. **Bisect a failure.** If a showcase fails, narrow it with a minimal repro **written
   inside `demos/showcases/`** (so the `pptx-ts` workspace dependency resolves)
   that adds just the suspect construct, and shrink it until a single `addX` call flips
   PASS→FAIL. `scripts/powerpoint-com-smoke.mjs` (`pnpm run test:com`) takes any deck with
   `--file <deck.pptx>` for the corruption-open check alone.

4. **Confirm the structural defect.** Extract the failing package and inspect the offending
   slide part. Common culprits and how to see them:
   - **Duplicate `<p:cNvPr id="…">`** on a slide (every id must be unique) — extract
     `ppt/slides/slideN.xml` and group the `id` values.
   - **Dangling `spid`/`r:embed`/`r:link`** — a `<p:spTgt spid>` or `r:embed` that targets
     no shape id / no relationship. Cross-check `ppt/slides/_rels/slideN.xml.rels`.
   - **Missing/incorrect `[Content_Types].xml` default** for a media extension.

   Extract to a **fresh** directory rather than deleting one — the agent sandbox has a
   false-positive guard that blocks `Remove-Item` when the surrounding command text also
   contains regex like `r:` or `\w+`.

5. **Fix, rebuild, and re-open** to confirm PASS. Add a regression test that reproduces the
   structural defect at the XML level (e.g. assert `cNvPr` ids are unique, or that a timing
   `spid` resolves to its own shape) so CI guards the class going forward — the PowerPoint
   check itself cannot run in CI.

## Notes

- The emitter uses `index + 2` (the slide-object index) for every shape's `<p:cNvPr>` id
  and for animation/media `spid` targets. Any id computed from a *different* space (a
  relationship id, a running counter) risks colliding or desyncing — a frequent source of
  0x80070570. When touching id/`spid` emission, smoke-test a slide that **mixes media with
  text/shapes**, not a single-object slide (where the two id spaces coincide and the bug
  hides).
- To author *reference* fixtures from real PowerPoint (rather than smoke-test ts-pptx
  output), use the `powerpoint-fixture-authoring` skill instead.
