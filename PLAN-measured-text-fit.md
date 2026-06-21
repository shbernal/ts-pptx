# Plan: Measured text fit (`fit` that actually fits in headless renders)

Status: proposal / not started
Owner: (fork)
Related downstream driver: downstream overflow back-and-forth (text spilling out
of cards/components; "extend the card more", "text still overlaps the icon").

## TL;DR

The `fit` autofit **markup already exists** in the fork and downstream already
uses it heavily (`fit: 'shrink'` in ~10 registry components). But the fork emits
the **bare flag with no baked result** (`<a:normAutofit/>` / `<a:spAutoFit/>`), and
a bare flag defers the fit computation to an interactive edit/resize event that a
headless render never fires. So in downstream's headless LibreOffice → PNG
pipeline (and on plain file-open) nothing recomputes and the text still overflows.
The components *think* they handle overflow; in the render they don't.

This is **not** an inherent renderer limitation — both PowerPoint and LibreOffice
*do* honor a fit that is already baked into the file. The catch is that the two
mechanisms store the fitted answer in **different places**, so the fix is not
symmetric:

| `fit` value | emits today | where the fitted answer lives | renderer on plain open |
| --- | --- | --- | --- |
| `'shrink'` | `<a:normAutofit/>` | `fontScale`/`lnSpcReduction` **attrs on the element** | applies stored scale; bare = no attrs = 100% = no shrink |
| `'resize'` | `<a:spAutoFit/>` | the shape's `a:ext/@cy` (+ `a:off/@y`) in `xfrm` | draws the box at whatever `cy` is already written |

`<a:normAutofit/>` is a no-op headless because, with no `fontScale` attribute,
both apps trust the stored scale (100%) rather than re-running layout on open.
`<a:spAutoFit/>` is really a *marker* ("this height was auto-computed, recompute
on edit") — the rendered height is not in the autofit element at all, it is in
`xfrm/ext/@cy`, so the renderer just draws the box at the `cy` already in the file.

The missing capability is therefore **measurement**: the library has no font
metrics, so it cannot compute the value to bake — a real
`fontScale`/`lnSpcReduction` (for shrink) or a fitted box height (for resize).
This plan adds serialize-time measured fit so overflow self-corrects in headless
renders without a human round-trip through PowerPoint.

This is generic PptxGenJS value — every consumer that renders or ships pptx
headlessly hits it — so it belongs upstream, not as a downstream workaround.

## Evidence (current state)

- API exists: `fit?: 'none' | 'shrink' | 'resize' | TextFitShrinkProps`
  (`src/core-interfaces.ts:1855`), with an object form
  `{ type:'shrink', fontScale?, lnSpcReduction? }` (`:1772`).
- Serialization: `src/gen-xml.ts:1582-1597`
  - `'shrink'` → bare `<a:normAutofit/>`
  - `'resize'` → `<a:spAutoFit/>`
  - object form → `genXmlNormAutofit()` (`:1511`) bakes explicit attrs.
- The fork's own comments document the limitation: *"Bare shrink does not work
  automatically — PowerPoint calculates fontScale/lnSpcReduction dynamically upon
  edit/resize"* (`gen-xml.ts:1585`) and the `fit` doc block
  (`core-interfaces.ts:1838-1854`): *"There is no way for this library to trigger
  that behavior… As a workaround, pass an object form of 'shrink' to bake explicit
  values."*
- No measurement anywhere: no `fontkit`/`opentype`/glyph-advance code, no such
  dependency (`package.json`). The object-form escape hatch requires the consumer
  to compute the scale by hand — which nobody can do reliably without metrics.
- Export pipeline is **async** (`src/pptxgen.ts:603` `exportPresentation`, awaits
  media/chart promises) → there is a legitimate place to do async font loading
  before the sync XML pass in `gen-xml.ts` consumes computed values.
- There is **no write-side font registration** today (font handling in `src/` is
  all on the read/theme-resolution side). Measurement needs a way to locate the
  actual font file for a face name.

## Goal

Given a text box (or table cell) with known runs, font face/size/bold/italic,
inner box dimensions (box minus `lIns/rIns/tIns/bIns`), wrap mode, and char
spacing, the library can:

1. **shrink** — compute the largest `fontScale` (and optional `lnSpcReduction`)
   at which the wrapped text fits the box, and bake it into
   `<a:normAutofit fontScale=… lnSpcReduction=…/>` so the render shows fitted text.
2. **resize** — compute the box height the text needs and grow the shape
   (`spPr/a:xfrm/a:ext` + offset per vertical anchor), the baked equivalent of
   `spAutoFit`.

Both must produce output that is correct in a headless LibreOffice render and on
plain file-open in PowerPoint, with no manual edit/resize.

## Design

### 1. Font metrics provider
- Add a metrics layer that loads a font file (TTF/OTF/TTC) and exposes per-glyph
  advance widths + ascent/descent/line-gap (units/em).
- **Lib choice: `opentype.js`** (decided). We only need raw cmap + hmtx advances
  and hhea/OS-2 vertical metrics, which opentype.js exposes directly and lighter
  than fontkit. Crucially we **do not** want full GPOS/GSUB shaping: kerning
  almost always narrows a line, so summing raw advances over-estimates width —
  the conservative direction we want (shrink a touch too much, never overflow).
  Revisit fontkit only if the real downstream font set hits TTC/CFF/variable
  files opentype.js cannot parse (Aptos ships as plain TTF, so unlikely).
- The genuine accuracy risk is **vertical line metrics** (hhea vs OS/2 win-* vs
  typo-*, gated by the `USE_TYPO_METRICS` bit), not width. PowerPoint and
  LibreOffice can choose different pairs; the solver must be conservative against
  the taller. This is what the calibration fixtures (below) pin down.
- Cache parsed fonts by resolved file path; key registrations by
  `(face, bold, italic)` since variant advances differ.

### 2. Font resolution / registration (new public surface)
- Consumers must tell the library where a face's file lives, since the fork does
  not embed/register fonts on the write side today. Proposed API:
  - `pptx.registerFontMetrics(face: string, source: string | Uint8Array, opts?: { bold?: boolean; italic?: boolean })`
  - Optional auto-resolution helper on the Node runtime (e.g. shell out to
    `fc-match`) kept out of core so the browser build stays clean.
- If a face has no registered metrics, measured fit **degrades gracefully**:
  fall back to the current bare-markup behavior (and optionally a heuristic
  average-advance table) and emit a single console.warn — never throw.

### 3. Line-break simulator
- Implement `wrap=square` line breaking matching PowerPoint's greedy
  word-wrap: break on whitespace, hard-break over-long tokens, honor `\n`,
  `charSpacing`, bold/italic metrics, and multi-run paragraphs.
- Output: wrapped line count + total laid-out height for a given font scale.
- Accuracy target: conservative (slightly over-estimate width and line height) so
  we shrink/grow a touch too much rather than overflow. Document known gaps
  (kerning, ligatures, complex scripts) up front.

### 3a. Fidelity strategy: calibrate against PowerPoint, don't derive it
- **Decided approach.** Rather than try to reproduce PowerPoint's layout engine
  analytically, calibrate the model against PowerPoint-authored fixtures where
  PowerPoint itself baked the fit value, and hold the solver to them as a
  conservative regression target (computed `fontScale` ≤ PowerPoint's; computed
  `cy` ≥ PowerPoint's *and* ≥ the LibreOffice-rendered height).
- This is a **fixture-gated precondition**: see `PLAN-autofit-calibration-fixtures.md`.
  No solver code lands until that oracle exists (per the OOXML / fixture-gated
  rules in `CLAUDE.md`). The fixtures also pin the PowerPoint-vs-LibreOffice
  line-height divergence the analytic model alone cannot resolve.
- Calibration font matrix: **Aptos, Aptos SemiBold, Calibri, Tahoma, Arial**
  (per-font metrics) with PowerPoint's font-independent *policy* — fontScale step
  set, `lnSpcReduction` trade, per-anchor `off.y` — swept on Aptos only.

### 4. Solvers
- **shrink**: binary-search `fontScale` in `[minScale, 100]` (e.g. minScale 25)
  until laid-out height ≤ inner height; optionally trade some `lnSpcReduction`
  before dropping font size, mirroring PowerPoint's behavior.
- **resize**: compute required inner height at full size → set `ext.cy`; adjust
  `off.y` for `anchor` (t/ctr/b) so growth direction matches PowerPoint. Note the
  render is driven entirely by the baked `ext.cy` — the `<a:spAutoFit/>` marker is
  not what the renderer applies, it only preserves autofit semantics for a later
  human edit. Unlike shrink, resize has **no safety net**: under-estimate the
  height and the text overflows (there is no text scaling fallback), so the wrap
  simulator must err tall here, not just wide.

### 5. Integration point
- Run a **measured-fit pass during async export**, before `gen-xml`'s sync body
  build, writing computed `fontScale`/`lnSpcReduction` (or new `ext`/`off`) onto
  the slide object. Keep `gen-xml.ts` consuming pre-computed values only.
- Drive it off the existing `fit` value so consumers opt in by what they already
  write. Decide one of:
  - (a) make `fit:'shrink'` / `'resize'` measure automatically **when metrics are
    registered** (zero API churn for downstream — its components already say
    `fit:'shrink'`), or
  - (b) add an explicit `{ type:'shrink', measure:true }` / `'shrink-measured'`
    to keep current bare behavior the default.
  Recommendation: (a). Per the fork's no-back-compat-obligation policy, changing
  what `'shrink'` emits when metrics are present is acceptable and is the
  least-friction path for the one real consumer; the no-metrics path is unchanged.

## API surface (proposed)

```ts
pptx.registerFontMetrics('Aptos', '/path/to/Aptos.ttf')
pptx.registerFontMetrics('Aptos', aptosBoldBytes, { bold: true })

slide.addText(runs, { x, y, w, h, fontFace: 'Aptos', fit: 'shrink' })
// → with metrics registered, exports <a:normAutofit fontScale="83000"/>
//   computed to fit; without metrics, exports bare <a:normAutofit/> + warns once.
```

## Phases

- **P0 — calibration fixtures (blocking precondition).** Author the PowerPoint
  autofit oracle and extract the calibration table. Fully specified in
  `PLAN-autofit-calibration-fixtures.md`. P1 does not start until P0 is done.
- **P1 — measurement + shrink solver (highest value).** Provider + registration +
  wrap simulator + shrink binary search + integration for text boxes. This alone
  kills the overflow-out-of-card class for downstream.
- **P2 — resize solver.** `fit:'resize'` grows the box; covers "extend the card
  more" directly.
- **P3 — table cells + heuristic fallback table.** Extend to `gen-tables` cells;
  add an average-advance fallback so unregistered fonts still improve.

## Risks / open questions

- Metric fidelity vs PowerPoint's own layout engine (kerning, ligatures, GPOS) and
  vs LibreOffice's line metrics. Mitigated by erring conservative (raw advances,
  taller-of-two line height) **and** the calibration-fixture regression target
  (`PLAN-autofit-calibration-fixtures.md`) — this is the primary fidelity control.
- Browser bundle size: keep `opentype.js` metrics on the Node path / lazy-load.
- Variable fonts & TTC face selection (lower risk with opentype.js on the known
  downstream TTF set; revisit fontkit only if a real font file fails to parse).
- CJK / RTL / complex shaping out of scope for P1 (document as unsupported).
- Performance on large decks (cache fonts; only measure boxes that opt in).
- **resize ≠ "extend the card".** Baking `ext.cy` grows only the *text box*. The
  card background rectangle and the icon next to it are separate shapes the library
  does not know are related, so resize alone will not "extend the card more" or
  un-overlap the icon — that layout coordination lives in the downstream
  component, not in `spAutoFit` baking. This makes P1 (shrink) the higher-leverage
  fix for the actual driver; P2 (resize) helps standalone text boxes but is only a
  partial answer for grouped card components.

## Tracking / housekeeping (fork conventions)

- Add a `CHANGELOG.md` `[Unreleased]` entry when P1 lands (this is a feature, with
  a behavior change to `'shrink'` when metrics are registered — note the migration).
- Record the work item in `docs/backlog.yml` (feature, not a downstream-need
  stopgap, since the proper home is the fork itself); cross-link the downstream
  driver.
- Acceptance: a fixture deck with deliberately overflowing card text renders
  fitted through the LibreOffice → PNG path (compare against a baseline PNG), and
  opens fitted in PowerPoint without an edit/resize.
