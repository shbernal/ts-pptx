# Fixture authoring recipes

The desktop-PowerPoint COM recipes that produced the `.pptx` fixtures in the parent
directory, plus the Node scripts that derive the committed `*.oracle.json` and
`*.cases.json` sidecars from them.

These are **recipes, not tests.** Nothing here runs in CI: every `author-*.ps1` needs a
licensed desktop PowerPoint and an interactive Windows session, and `measure-lo.py` needs
a local LibreOffice. Read
[`.agents/skills/powerpoint-fixture-authoring/SKILL.md`](../../../../.agents/skills/powerpoint-fixture-authoring/SKILL.md)
before running any of them — it carries the COM ordering rules, the teardown/reap
discipline, and the autofit bake-on-save contract that these scripts depend on.

They are tracked so a committed fixture can be re-authored or extended from a clean
checkout. Nothing in this directory is published to npm (`package.json` `files: ["dist"]`).

## Running them

All scripts resolve their own paths from `$PSScriptRoot` / `import.meta.url`, so they can
be invoked from anywhere:

```powershell
& test\read\fixtures\authoring\author-read-stress.ps1        # re-authors ../read-stress.pptx
node test/read/fixtures/authoring/build-oracles.mjs          # re-derives the oracle sidecars
node test/read/fixtures/authoring/gen-cases.mjs --help       # the one here that takes flags
```

`gen-cases.mjs` and `measure-lo.py` were duplicated under `scripts/` until August 2026,
with nothing keeping the copies in step. This directory is the single home for both:
their output feeds the Windows authoring step next door, so they are recipes like
everything else here. `scripts/` keeps only `extract-autofit-calibration.mjs`, which
derives its table from the finished decks and needs no desktop app —
[`scripts/README.md`](../../../../scripts/README.md) states the boundary.

Two things to know before you commit the result:

- **Re-run Prettier.** The sidecars are committed in Prettier's format, but the builders
  emit raw `JSON.stringify(…, '\t')`. After regenerating, run
  `pnpm exec prettier --write "test/read/fixtures/*.json"` or the diff will be pure
  whitespace noise. With that done, every builder here reproduces its committed sidecar
  byte-for-byte — that round-trip is the check that a recipe is still honest.
- **A re-authored `.pptx` will not be byte-identical** to the committed one (PowerPoint
  stamps fresh ids/timestamps), so its SHA-256 changes. Update the hash and the
  provenance line in [`../README.md`](../README.md) whenever you replace a fixture, and
  re-verify with the skill's `verify-powerpoint-fixture.ps1`.

## What produces what

| Recipe                                                                                                                                                             | Produces                                                                                                                                                                                                                                                                                                                                                                                                       |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `author-deck.ps1`                                                                                                                                                  | **Engine.** Parameterized (`-CasesPath`, `-OutPath`) authoring driver used by `author-all.ps1`; the worked example the authoring skill cites for the pin-then-text-then-AutoSize ordering and the font-readiness guard block.                                                                                                                                                                                  |
| `author-all.ps1` + `gen-cases.mjs` + `measure-lo.py` + `readiness-guard.ps1`                                                                                       | The autofit calibration matrix: `autofit-{shrink,resize,edge,line-metrics}.pptx` and their 4 committed `*.cases.json`. `readiness-guard.ps1` certifies the five required faces actually resolve through GDI (PowerPoint silently substitutes otherwise) and that `soffice` is reachable; `measure-lo.py` reads LibreOffice's recomputed geometry over UNO and feeds `scripts/extract-autofit-calibration.mjs`. |
| `author-read-stress.ps1`                                                                                                                                           | `read-stress.pptx` — the multi-dimension integration fixture (two masters/themes, nested groups, styled tables, dual raster+SVG picture, recolor, embedded fonts, threaded comments, notes).                                                                                                                                                                                                                   |
| `author-table-styles.ps1`                                                                                                                                          | `table-styles.pptx`. Applying a _built-in_ style is what makes PowerPoint materialize real definitions into `ppt/tableStyles.xml`; only Microsoft built-in style GUIDs are used, so the fixture stays brand-free.                                                                                                                                                                                              |
| `author-slide-transition.ps1`, `author-slide-transition-sound.ps1`                                                                                                 | `slide-transition{,-sound}.pptx`. The sound fixture embeds `assets/ding.wav`.                                                                                                                                                                                                                                                                                                                                  |
| `author-slide-animation-{basic,presets,rich}.ps1`                                                                                                                  | The three `slide-animation-*.pptx` timing/build fixtures.                                                                                                                                                                                                                                                                                                                                                      |
| `author-import-animation-merge.ps1`                                                                                                                                | `import-animation-merge.pptx` — the cross-slide paste-with-animation spid-remap ground truth.                                                                                                                                                                                                                                                                                                                  |
| `author-anim-probe.ps1`, `probe-animeffect.ps1`, `parse-animeffect.ps1`                                                                                            | Probe decks + `animeffect-order.json` behind the `MsoAnimEffect` findings in [`docs/animations-and-transitions.md`](../../../../docs/animations-and-transitions.md).                                                                                                                                                                                                                                           |
| `author-table-cell-horzoverflow.ps1`                                                                                                                               | `table-cell-horzoverflow.pptx`. The attribute has no COM or UI surface, so this one injects it into `slide1.xml` and then has PowerPoint reopen and re-save the deck — the committed bytes are PowerPoint's own serialization. An explicit `horzOverflow="clip"` is dropped by that round-trip (it is the schema default), which is why the fixture pins only `overflow` plus a bare control.                     |
| `probe-table-cell-wrap.ps1`                                                                                                                                        | **Probe; produces no committed fixture.** Establishes that a table cell cannot be made non-wrapping (`WordWrap` is read-only over COM, and `<a:bodyPr wrap="none"/>` renders inert and is stripped on the next save) and that `a:tcPr/@horzOverflow` is a glyph clip/overflow control rather than a wrap switch. Cited from `src/gen/slide/objects/table.ts`; all output lands in `.tmp/`.                       |
| `probe-table-cell-a11y-and-3d.ps1`                                                                                                                                 | **Probe; produces no committed fixture.** Settles three `a:tc`/`a:tcPr` constructs with no COM or UI surface, by injecting all three and having PowerPoint reopen and re-save: `a:tc/@id` and `a:tcPr/a:headers` (the screen-reader header association) are **stripped**, while `a:tcPr/a:cell3D` survives verbatim. Controlled rather than circumstantial — `a:cell3D` and `a:headers` went into the *same* `a:tcPr` and only one came back — which is what makes header association a documented non-authorable construct rather than an unimplemented one. Cited from `src/read/api/table.ts` and [`docs/tables.md`](../../../../docs/tables.md); output lands in `.tmp/`. |
| `probe-entryeffect-table.ps1` → `parse-entryeffect-table.mjs` → `entryeffect-table.json`                                                                           | The probed `PpEntryEffect → {element, ns, variant, modernOnly}` table (159 accepted ints, `0` = none). **`build-oracles.mjs` reads `entryeffect-table.json` as a required input**, so it is committed here rather than regenerated per run — reproducing it needs a full 0..4096 COM sweep.                                                                                                                    |
| `build-oracles.mjs`                                                                                                                                                | `slide-transition.oracle.json` (embeds the table above) + the `slide-animation-{basic,rich}` oracles.                                                                                                                                                                                                                                                                                                          |
| `build-{sound,merge,presets}-oracle.mjs`                                                                                                                           | The `slide-transition-sound`, `import-animation-merge` and `slide-animation-presets` oracles.                                                                                                                                                                                                                                                                                                                  |
| `author-omml.ps1` / `author-omml-inline.ps1`                                                                                                                       | `math-omml.pptx` (display equation) and `math-omml-inline.pptx` (inline, no `m:oMathPara`). Both build a genuine equation in Word first rather than hand-typing OMML.                                                                                                                                                                                                                                          |
| `author-modern-comments.ps1`                                                                                                                                       | `modern-comments.pptx`. Modern (2018) comments need `Comments.Add2` with an explicit `providerId`; an empty one raises "Illegal value" and silently falls back to legacy comments.                                                                                                                                                                                                                             |
| `author-picture-media.ps1`                                                                                                                                         | `picture-media.pptx`. See the asset note below.                                                                                                                                                                                                                                                                                                                                                                |
| `make-cube-glb.mjs` → `author-model3d.ps1` | `assets/cube.glb` then `model3d.pptx` — the `am3d:model3d` graphic-frame oracle. Plain `pwsh` COM is reliable for `Shapes.Add3DModel`; no `cscript`/retry needed. The `.glb` is generated rather than downloaded so the fixture input is license-clean by construction — and PowerPoint re-exports it on insert anyway, so the stored part is Microsoft's own GLTF export of it, not the input bytes. |
| `author-bar-chart.ps1`, `author-default-text-style.ps1`, `author-footer-trio.ps1`, `author-online-video.ps1`, `author-template-potx.ps1`, `extend-multi-theme.ps1` | The correspondingly named fixtures. `author-online-video.ps1` deliberately targets a path under `C:\Users\Public` so the committed fixture's external link carries no user directory.                                                                                                                                                                                                                          |
| `make-assets.ps1`                                                                                                                                                  | Regenerates the programmatic media (`photo.png`, `mark.png`) into `<repo>/.tmp/media/` — gradient/shape rasters drawn in code, so they are license-clean by construction.                                                                                                                                                                                                                                      |

## Assets

`assets/ding.wav` is committed: 844 bytes of self-generated 16-bit PCM mono 8 kHz sine,
which is what `slide-transition-sound.oracle.json` cites for its license-clean claim.

`assets/cube.glb` is committed too, and is likewise generated rather than sourced — 1516
bytes emitted by `make-cube-glb.mjs` (one mesh, 24 vertices, one PBR material, no textures).
Regenerate it with `node test/read/fixtures/authoring/make-cube-glb.mjs`; the output is
deterministic, so a clean checkout reproduces the committed bytes.

The other source images these recipes once consumed are **not** committed. In particular
the SVG behind `picture-media.pptx`'s `SvgPic` shape was a Microsoft Office stock icon
(_Insert → Icons_) — that is fine baked inside a fixture whose provenance says so, but
not as a standalone file in this repo. To re-author that fixture, re-insert an icon the
same way rather than looking for a checked-in copy.
