# Plan: Autofit calibration fixtures (the oracle for measured text fit)

Status: **DONE (2026-06-21).** Four PowerPoint-authored decks
(`autofit-line-metrics/-shrink/-resize/-edge.pptx`, 149 cases), the derived
`test/read/fixtures/autofit-calibration.json` (with the LibreOffice cross-measure
column), the regeneration tooling (`scripts/gen-autofit-cases.mjs`,
`scripts/measure-autofit-lo.py`, `scripts/extract-autofit-calibration.mjs`), the
provenance/hashes/case-id scheme/findings note in `test/read/fixtures/README.md`,
and the backlog entry (`dn-measured-text-fit`) are all committed. See the README
"Autofit calibration oracle" section for the findings. Acceptance items 1–5 below
are satisfied; `PLAN-measured-text-fit.md` P1 is now unblocked.
Owner: (fork)
Blocks: `PLAN-measured-text-fit.md` (measurement + solvers). This is the
**fixture-gated precondition** for that work — per `docs/backlog-workflow.md`
("Fixture-Gated Work") and the OOXML rule in `CLAUDE.md`, the solver must be built
against genuine PowerPoint output, not synthetic or round-tripped XML. No solver
code lands until the oracle here exists.

## Why this comes first

The measured-fit solvers (shrink `fontScale`/`lnSpcReduction`, resize `ext.cy`)
have to reproduce PowerPoint's own layout decisions closely enough that:

- **shrink** never under-shrinks (text must not overflow in either renderer), and
- **resize** never under-grows (no safety net — see main plan).

We cannot validate "close enough" against prose or against XML we generated
ourselves. We need PowerPoint-authored boxes where **PowerPoint itself computed
and baked the fit value**, read those values back, and hold the solver to them as
a conservative regression target. Those boxes are these fixtures.

Chosen approach (from design discussion): metric reads via **opentype.js**
(cmap + hmtx advances + hhea/OS-2 vertical metrics; deliberately no GPOS/GSUB
shaping — raw advances over-estimate width, which is the conservative direction),
and a **calibrate-against-PowerPoint-fixtures loop** rather than trying to derive
PowerPoint's engine analytically.

## The hard problem these fixtures must pin down

Width-summing is easy and conservative. The fidelity risk is **vertical line
metrics**: total laid-out height = `lineCount × lineHeight`, and `lineHeight` is
where renderers diverge.

- "Single" spacing is not `1.0 × fontSize`; it is font-metric-derived.
- A font carries two ascent/descent pairs (hhea vs OS/2 win-* vs OS/2 typo-*);
  which one applies depends on the `USE_TYPO_METRICS` fsSelection bit.
- **PowerPoint and LibreOffice can pick different pairs.** downstream renders
  through headless LibreOffice but the file must also be correct in PowerPoint, so
  the solver has to be conservative against the **taller** of the two. The
  fixtures must measure both engines, not just PowerPoint.

## Design principle: separate per-font metrics from font-independent policy

A full combinatorial sweep (5 fonts × every other axis) explodes. Split the axes:

- **Per-font metric calibration** — sweep all 5 fonts on a small core of
  scenarios. This pins each family's advance widths and effective line height.
  Fonts: **Aptos, Aptos SemiBold, Calibri, Tahoma, Arial** (M365 box has Aptos
  natively — confirmed — so no substitution; the substitution guard below still
  runs to prove it).
- **Policy calibration** — sweep the *other* axes on a single anchor font
  (**Aptos**, regular + bold only). PowerPoint's fit *policy* — discrete
  fontScale steps, how `lnSpcReduction` trades against `fontScale`, how `off.y`
  moves per vertical anchor — is font-independent, so it only needs one family.

## Authoring mechanics

Follow `.agents/skills/powerpoint-fixture-authoring/SKILL.md` exactly (PowerPoint
COM via pwsh 7, `& '<script>.ps1'`, no `-ExecutionPolicy Bypass`; snapshot/reap
`POWERPNT` PIDs; `SaveAs` + `Saved=$true`; verify with the skill's helper scripts).

- **Shrink** boxes: fixed `w`/`h`, `shape.TextFrame2.AutoSize = msoAutoSizeTextToFitShape (2)`,
  then add enough text to overflow → PowerPoint bakes `<a:normAutofit fontScale=… lnSpcReduction=…/>`.
- **Resize** boxes: `shape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText (1)`,
  add text → PowerPoint writes `<a:spAutoFit/>` and the fitted `xfrm/ext/@cy`
  (+ `off/@y`).
- Deterministic geometry, colors, and a stable **case-id shape name** on the box
  under test (e.g. `shrink__aptos__sz18__b0i0__w3.00h1.00__ovr3x`). A sibling
  `*.cases.json` manifest maps case-id → full inputs (text string, insets,
  spacing, anchor, charSpacing) since not everything fits cleanly in a name.

### Precondition spike (do this before authoring the matrix)

Two things can invalidate the whole oracle; prove them out on 1–2 throwaway slides
first:

1. **Bake-on-save.** Confirm that setting `AutoSize` via COM + `SaveAs` actually
   writes a non-100 `fontScale` (and a fitted `cy`) **without interactive
   editing**. If it does not, find the reliable trigger (set text *after*
   AutoSize, nudge the shape size, or drive `ExecuteMso` on a selection per the
   skill's enum-marshalling fallback) and document the recipe. If no
   non-interactive recipe exists, that is a blocking finding to surface before
   proceeding.
2. **Substitution guard.** Dump the saved slide XML and assert the run's
   resolved `latin@typeface` equals the requested face (Aptos must not have been
   substituted). Bake a tiny assertion into the verify step so a future
   font-missing machine fails loudly instead of silently calibrating to Calibri.

## Fixture decks

Group by concern into a few multi-slide decks (one slide = one case), committed to
`test/read/fixtures/`. Multi-slide is the only sane shape for a matrix this size
and still matches repo convention (one deck per concern).

### Deck 1 — `autofit-line-metrics.pptx` (foundation: line height + advances)

Pins each font's effective single-line height and advance widths straight from
XML, with no wrapping ambiguity.

- Per font ∈ {Aptos, Aptos SemiBold, Calibri, Tahoma, Arial}, per size ∈ {12, 18, 32}:
  a `msoAutoSizeShapeToFitText` box containing **1, 2, and 3** lines of
  **non-wrapping** text (hard `\n`, single spacing, zero space-before/after,
  default insets). Back out single-line height = `(cy@K − cy@1) / (K−1)` and the
  first-line/inset overhead = `cy@1 − height`.
- A fixed-size, `AutoSizeNone` box per font/size with a **known character string**
  (e.g. all lowercase, then all uppercase, then digits) for advance-width
  cross-check against opentype.js hmtx output.
- **LibreOffice cross-measure:** render this deck headless → PNG and measure the
  text-block height per case. Record PowerPoint-cy vs LibreOffice-rendered height
  side by side; the delta is the USE_TYPO_METRICS divergence the solver must
  bound. Store both in the extracted table.

### Deck 2 — `autofit-shrink.pptx` (normAutofit calibration)

- **Per-font core** (all 5 fonts, regular; + SemiBold/bold where applicable),
  one anchor scenario: `w≈3.0in h≈1.0in`, size 18, text overflowing ~3×, single
  spacing, default insets, `wrap=square`. → per-font baked `fontScale`.
- **Policy sweep (Aptos regular+bold):**
  - Overflow-magnitude ladder: ~1 line over, ~1.5×, 2×, 3×, 5× → learn the
    discrete fontScale step set and the search range.
  - `lnSpcReduction` onset: heavy overflow cases to capture when PowerPoint trades
    line-spacing reduction *before/with* dropping `fontScale`, and the step values.
  - Line-spacing variants: single, 1.5×, exact-points → effect on fit.
  - `spaceBefore`/`spaceAfter` non-zero, multi-paragraph.
  - Multi-run paragraph (mixed bold and mixed size within one wrapped line) → line
    height = max run height; wrap measured across run boundaries.
  - Non-default insets (`lIns/rIns/tIns/bIns`) and a non-default `charSpacing`.
  - Vertical anchor t/ctr/b (should not change `fontScale`; verify and document).
  - Width-driven single-line overflow (one long title) vs height-driven multiline.

### Deck 3 — `autofit-resize.pptx` (spAutoFit / baked cy calibration)

- **Per-font core** (all 5 fonts): single-line and 3-line wrapped boxes →
  per-font `cy` (corroborates Deck 1).
- **Policy sweep (Aptos):**
  - Vertical anchor t/ctr/b → learn how PowerPoint adjusts `off.y` when growing
    (grows down / both ways / up).
  - Does spAutoFit also **shrink** the box below the authored height when text is
    short? Capture an under-filled box.
  - Line-spacing and space-before/after contribution to `cy`.
  - Insets contribution to `cy`.

### Deck 4 — `autofit-edge.pptx` (Aptos; the cases most likely to break the simulator)

- Over-long single token with no break opportunity → character-wrap behavior.
- Trailing spaces at line end (do they count toward line width?).
- Empty paragraph / blank line height contribution.
- Tab character handling.
- Leading/trailing whitespace runs and a run that is only whitespace.
- Mixed font sizes in one paragraph (line-height = max).
- (Documentation-only) one CJK + one RTL box marked **unsupported** so the gap is
  recorded as evidence rather than discovered later.

## Extraction → committed calibration table (so CI stays Node-only / Linux)

The `.pptx` files are the provenance/oracle, but the Node test suite runs headless
on Linux and must not depend on Windows or PowerPoint. So:

- A dump script (extend `scripts/dump-slide-xml.ps1` or add a small reader using
  the fork's existing read model) extracts, per case-id: **inputs** (font, size,
  bold/italic, box w/h, insets, spacing, anchor, charSpacing, text) and
  **PowerPoint outputs** (`fontScale`, `lnSpcReduction`, or `ext.cy`/`off.y`),
  plus the LibreOffice-measured height for Deck 1.
- Commit the result as `test/read/fixtures/autofit-calibration.json` (the derived
  table). Solver unit tests read this JSON and assert the solver is conservative
  against it (computed `fontScale` ≤ PowerPoint's; computed `cy` ≥ PowerPoint's
  and ≥ LibreOffice-measured height). The JSON is regenerable from the `.pptx`;
  the `.pptx` remains the source of truth.

## Deliverables / acceptance for THIS plan

1. Precondition spike resolved: documented, reliable non-interactive bake recipe
   (or an escalated blocking finding) + substitution guard wired into verify.
2. Four fixture decks authored by desktop PowerPoint, verified opening with no
   repair prompt, SHA-256 recorded.
3. `test/read/fixtures/README.md` updated with provenance, hashes, purpose, the
   case-id naming scheme, and the desktop PowerPoint check date.
4. `autofit-calibration.json` extracted and committed, with the LibreOffice
   cross-measure column for Deck 1.
5. A short findings note (in this file or the README) capturing what the oracle
   revealed: PowerPoint's fontScale step set, the lnSpcReduction trade policy, the
   per-anchor `off.y` rule, and the measured PowerPoint-vs-LibreOffice line-height
   delta per font. These findings directly parameterize the solver.

Only once 1–5 are done does `PLAN-measured-text-fit.md` P1 implementation start.

## Housekeeping

- Record this as the blocking fixture precondition in `docs/backlog.yml` against
  the measured-fit feature item (tag the relevant `constructs:` key for autofit),
  cross-linking the downstream overflow driver.
- Keep the authoring `.ps1` scripts out of commits (temp `.tmp/` per the skill);
  commit only the `.pptx`, the extracted JSON, and README/doc updates.
