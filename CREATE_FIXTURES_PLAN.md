# Create-Fixtures Plan & Status

Status of the real PowerPoint-authored `.pptx` read fixtures and the read-model
work they unblock. Originally an analysis of fixtures to add; now updated to
record what has landed and what remains.

Scope: the curated read/round-trip fixtures in `test/read/fixtures/`. These are
genuine Microsoft PowerPoint output (see `test/read/fixtures/README.md`), used by
the `pptxgenjs/read` harness to prove the read model holds against real Office
OOXML — **not** PptxGenJS's own serializer. Author every fixture with the
`powerpoint-fixture-authoring` skill (desktop PowerPoint COM), never with
PptxGenJS, then record provenance + SHA-256 in the fixtures README.

## TL;DR — current state (2026-06-19)

All four planned fixtures plus the `group-transform` extension shipped, their
read tests are wired to read genuine PowerPoint XML (not the old write→read
round-trips), the two implied read fixes landed, and `pnpm run test:read` is
**245/245 green**.

| Fixture | Added | Tests wired | Fix landed | State |
|---|---|---|---|---|
| `theme-colors.pptx` | yes | yes | n/a (accessor already shipped) | ✅ done |
| `gradient-fill.pptx` | yes | yes | n/a | ✅ done |
| `preset-geometry.pptx` | yes | yes | n/a | ✅ done |
| `multi-theme.pptx` | yes | yes | `fix(read): resolve p:style fillRef/lnRef` (59ea7bcf) | ✅ done (fillRef/lnRef); restyle-literals still deferred |
| `group-transform.pptx` (extended) | yes | yes | `fix(read): compose group rotations in absolute frames` (6871b337) | ✅ done |

Fixtures were added in `4b1293c3 addition of new fixtures`; all five carry a
non-default **Ion** theme / PowerPoint-authored constructs and passed a Windows
desktop PowerPoint open-clean check on 2026-06-18 (README provenance + SHA-256
recorded).

The original problem this set out to fix — read-model accessors validated only by
**round-tripping through PptxGenJS's own writer** (circular evidence) or
**hand-typed oracle tables** — is resolved for color transforms, theme
resolution, gradient stops, preset-geometry adjusts, and style-matrix fill/line.
One small write→read round-trip remains (per-shape rotation/flip; see Remaining
Work C).

---

## Delivered fixtures

### `theme-colors.pptx` — ✅ done

One Ion-theme slide of named rectangles, each filled with a scheme color carrying
a distinct transform, plus a scheme-color line and a scheme-color text run.

- Shapes: `accent1-plain`, `accent1-lm60-lo40`, `accent2-lm75`, `accent3-shade50`,
  `accent4-tint40`, `lt2-lm20-lo80`, `explicit-srgb-fill`,
  `accent1-line-accent2-2pt` (2 pt scheme line), `text-accent5-run`.
- `test/read/color-transform.test.js` now reads its oracle straight from the
  fixture: `POWERPOINT_ORACLE` maps each shape name → the PowerPoint-COM effective
  RGB (`accent1-plain B01513`, `accent1-lm60-lo40 ED5654`, …), and asserts
  `applyColorTransforms(resolvedFill.hex, resolvedFill.transforms)` matches within
  ±1/channel. The old synthetic transform table is gone; only algebra edge cases
  (alpha/alphaMod, clamps, identity) remain hand-written, which is correct.
- `test/read/style-accessors.test.js`:
  - "Shape style reads — minimal real PowerPoint fixtures" → `lineWidthPt` (2 pt),
    `lineSchemeColor` (`accent2`), `resolvedLine.hex` (`EA6312`).
  - "Theme colour resolution — real PowerPoint XML (theme-colors.pptx)" →
    `resolvedFill` for scheme vs explicit srgb, base hex + raw transforms +
    `effectiveHex`, and `Run.resolvedColor` (`accent5 → 54849A`).

### `gradient-fill.pptx` — ✅ done

One slide of named gradient rectangles + a solid control.

- Shapes: `grad-linear-2` (2 srgb stops 0%/100%), `grad-linear-3-scheme` (3 stops,
  first stop `accent1` scheme with `effectiveHex B01513`), `grad-radial` (path
  gradient), `solid-control`.
- `style-accessors.test.js` "gradientStops reads PowerPoint-authored gsLst stops"
  asserts stop counts, positions (0/0.5/1), the srgb-vs-scheme color split with
  `effectiveHex`, that a radial/path gradient still exposes its stops, and that the
  solid control reports `gradientStops === null`. Replaces the former write→read
  round-trip.

### `preset-geometry.pptx` — ✅ done

One slide of PowerPoint-authored autoshapes with dragged adjust handles.

- Shapes: `roundRect-adj` (`adj = val 12000`), `chevron-adj` (`adj = val 35000`),
  `blockArc-adj1-adj2-adj3` (three handles incl. angle guides `7200000`,
  `30000000`), `rect-no-adjust` (control, `{}`).
- `style-accessors.test.js` "adjustValues exposes PowerPoint-authored avLst
  handles" asserts each handle name/value and the empty control. Replaces the
  former roundRect write→read round-trip. (The fixture went beyond the plan: a
  three-handle `blockArc` instead of a generic two-handle shape.)

### `multi-theme.pptx` — ✅ done (fillRef/lnRef); one deferred follow-up

One Ion-theme source slide whose shapes color through the theme — including a
shape with **no explicit spPr fill/line** that resolves only via the `p:style`
fillRef/lnRef matrix.

- Shapes: `style-matrix-default` (no explicit fill/line; fillRef idx 1 accent1,
  lnRef idx 2 accent1+shade), `scheme-accent1-fill` (explicit accent1 fill +
  accent2 line over a style matrix), `scheme-accent2-lighter`, `scheme-accent5-text`.
- Drove `fix(read): resolve p:style fillRef/lnRef in resolvedFill/resolvedLine`
  (59ea7bcf) — `resolvedFill`/`resolvedLine` now walk the style matrix when the
  shape carries no explicit color. `style-accessors.test.js` "Style-matrix
  fill/line resolution" asserts the fillRef resolves (`B01513`), the lnRef carries
  its `shade 15000` transform, and an explicit spPr fill/line still wins.
- Also the source deck for `import-slide-preserve.test.js` and
  `import-slide-restyle.test.js` (imported into a blank `empty`/default-theme
  target), pinning that `preserve` bakes the Ion theme to literal srgb and
  `restyle` leaves refs symbolic against a genuinely different source theme.

### `group-transform.pptx` (extended) — ✅ done

Extended with the additional rotation/scale/nesting cases and its
grouped + ungrouped-twin slides; drove `fix(read): compose group rotations in
absolute frames` (6871b337). `Shape.absoluteFrame` now composes enclosing group
`@rot`/`@flipH`/`@flipV`. Backlog `dn-group-rot-flip-frame` is **implemented**.

---

## Remaining work (the "missing stuff")

### A. Placeholder-inherited run colour in the read getter — fixture + fix needed

`dn-readmodel-style-followups` had two halves: (1) walk `p:style` fillRef/lnRef,
and (2) resolve a **placeholder-inherited run colour** (a run whose `a:rPr`
carries no own color, inheriting from the layout/master placeholder). Half (1)
shipped via 59ea7bcf + `multi-theme.pptx`. Half (2) is **not** in the read getter:
`placeholderInheritedColor` exists in `src/read/oxml/theme.ts` only on the
`importSlide` flatten path, and `Run.resolvedColor` does not consult it. No
fixture exercises it.

- **Fixture:** a small deck whose slide has a placeholder (e.g. a title/body
  placeholder) with **no run-level color**, on a layout/master whose placeholder
  defines a scheme color. Either a new `placeholder-inherited.pptx` or a second
  slide added to `multi-theme.pptx` (it already carries the Ion theme +
  master/layout). Name the placeholder stably.
- **Fix + test:** extend `Run.resolvedColor` (`src/read/api/text.ts` /
  `theme-context.ts`) to fall back to `placeholderInheritedColor`; assert the
  inherited run resolves to the master/layout color while an explicit run color
  still wins.
- **Backlog:** `dn-readmodel-style-followups` is now `partially-implemented`
  (updated 2026-06-19) — fillRef/lnRef shipped (59ea7bcf), placeholder-inherited
  run colour is the remaining open leg (`next_action:
  layer-placeholder-inherited-run-colour`, stopgap retained), with
  `multi-theme.pptx` as the available fixture base.

### B. importSlide `restyle` literal force-remap + table style — feature deferred, fixture content to add

`dn-importslide-restyle-literals` (still `deferred`) wants: (a) a force-remap mode
that rewrites literal `srgbClr` matching a source-theme slot to the destination
slot, and (b) copying the source table style into the destination `tableStyles`.
`multi-theme.pptx` enables testing (a) once built but currently has **no**
literal-srgb shape whose hex equals an Ion theme slot, and **no** table with a
custom table style.

- **Fixture additions (when the feature is scheduled):** add to `multi-theme.pptx`
  (or a `restyle-literals.pptx`) one rectangle filled with a literal `srgbClr`
  equal to an Ion accent hex (so a force-remap has something to match), and one
  table using a non-default table style (so the table-style copy has a source).
- This is net-new feature work, not a wiring gap; only add the fixture content
  alongside implementing the mode.

### C. Per-shape rotation/flip — one write→read round-trip remains (low priority)

`style-accessors.test.js` "Per-shape rotation / flip" reads `Shape.rotation`/
`flipH`/`flipV` mostly from synthetic `spWithXfrm` XML using **real PowerPoint
angle values** (e.g. `rot="2259366"` → 37.6561°, the unsigned `19216344` →
320.2724°), which is good evidence. But one case still round-trips through the
write API (`addShape({ rotate: 45, flipH: true })` → reopen). To fully
de-circularize, add two ungrouped rotated/flipped rectangles to an existing
fixture (e.g. `theme-colors.pptx` or `preset-geometry.pptx`) — a rotated shape and
a flipped shape with stable names — and replace the round-trip assertion with a
read of those. Low priority: the raw-value synthetic tests already use authentic
PowerPoint magnitudes.

### D. Freeform `custGeom` read — out of scope unless a reader lands (note only)

The replication capability-map lists `cust-geom` as a construct. There is no read
accessor for freeform path geometry today and no fixture. Add a `custGeom` fixture
only if/when a read accessor for it is implemented; tracked here so it is not
forgotten.

---

## Generation-feature fixtures — moved out (on stand-by)

The two `interesting-with-tweaks` backlog items — `upstream-pr-1447` (native
PowerPoint comments) and `upstream-pr-1431` (animation engine) — are
**write/serialization** features, not read-model work, and are **on stand-by**
(not active). Their fixture plan now lives in
[`CREATE_FIXTURES_PLAN_GENERATION.md`](CREATE_FIXTURES_PLAN_GENERATION.md).

## Not needed (already covered or out of scope)

- **Charts** — `mixed.pptx` carries a real chart and `chart.test.js` reads it; no
  minimal-chart gap is recorded in the backlog.
- **Generation-only constructs** (`bodypr-vert`/`dn-text-direction-serialization`,
  `normautofit-shrink`/`fit:'shrink'`, gradient *line* stroke
  `dn-gradient-line-stroke`, connector presets) — pinned by serialization schema
  fixtures in `test/schema.test.js` that assert "option X emits attribute Y". Per
  `dn-doc-render-caveats` these are deliberately XML-contract assertions, not
  read-model fixtures, so no PowerPoint `.pptx` is needed.

## Workflow reminder

For any new fixture or fixture addition: author with the
`powerpoint-fixture-authoring` skill, run
`scripts/verify-powerpoint-fixture.ps1`, confirm the target OOXML construct is
present, update `test/read/fixtures/README.md` (provenance table, purpose,
SHA-256, desktop-PowerPoint check date), wire the read test to the fixture instead
of a round-trip/oracle path, and run `pnpm run test:read`.
