# Create-Fixtures Plan

Analysis of which **already-implemented** features need a new real
PowerPoint-authored `.pptx` fixture, and the exact shape each fixture must take.

Scope: the curated read/round-trip fixtures in `test/read/fixtures/`. These are
genuine Microsoft PowerPoint output (see `test/read/fixtures/README.md`), used by
the `pptxgenjs/read` harness to prove the read model holds against real Office
OOXML — **not** PptxGenJS's own serializer. Author every fixture below with the
`powerpoint-fixture-authoring` skill (desktop PowerPoint COM), never with
PptxGenJS, then record provenance + SHA-256 in the fixtures README.

## Why new fixtures are needed (the core finding)

The read-model **style accessors** shipped (`Shape.adjustValues`,
`Shape.gradientStops`, `Shape.lineWidthPt`, `Shape.resolvedFill` /
`resolvedLine`, `Run.resolvedColor`, `Slide.themeContext()`, and
`applyColorTransforms`). But their tests do **not** read genuine PowerPoint XML
for the interesting constructs. They fall back to:

- **Write→read round-trip** — `test/read/style-accessors.test.js` generates a
  deck with PptxGenJS, reopens it, and reads it back. Its own header admits:
  _"geometry adjusts and gradient stops are not in the vendored fixtures, so we
  round-trip them through the write API."_ This is **circular evidence**: the
  reader is only checked against the writer, so any shared wrong assumption about
  how PowerPoint encodes the construct passes silently.
- **Hand-authored oracle tables** — `test/read/color-transform.test.js` verifies
  `applyColorTransforms` against a typed table of `source → effective` hexes
  ("verified against PowerPoint/LibreOffice output" by hand), with no fixture
  carrying the actual `<a:schemeClr>` + `<a:lumMod>/<a:lumOff>` PowerPoint wrote.
- **Synthetic spTree strings** — `style-accessors.test.js` wraps standalone XML
  it builds itself for the DOM-only accessors.

The vendored/promoted fixtures (`empty`, `textbox`, `image`, `table`, `mixed`,
`hidden`, `group-transform`) do not carry preset-geometry adjust handles,
gradient fills, scheme-color transforms on shapes, or a second theme — so none of
these accessors is currently pinned to real PowerPoint output. Closing that is
the purpose of the fixtures below.

## Current fixture coverage

| Fixture | Pins (real PowerPoint construct) |
|---|---|
| `empty.pptx` | minimal OPC baseline |
| `textbox.pptx` | text runs / text-bearing parts |
| `image.pptx` | raster + SVG media, content-type by extension |
| `table.pptx` | `a:tbl` graphic frames |
| `mixed.pptx` | connectors, nested groups, chart, SmartArt, tables (enumeration) |
| `hidden.pptx` | `show="0"` hidden-slide flag |
| `group-transform.pptx` | group `@rot`/`@flipH`/`@flipV` (drives `dn-group-rot-flip-frame`) |

Gaps with no real-PowerPoint coverage: **scheme-color transforms**, **gradient
fills**, **preset-geometry adjust handles**, **explicit line widths**, and a
**second/non-default theme** for `importSlide` theme-flattening.

---

## Fixture 1 — `theme-colors.pptx` (highest value)

**Drives / pins**
- `applyColorTransforms` (`src/read/oxml/color-transform.ts:124`) — replaces the
  hand oracle in `color-transform.test.js` with genuine Office luminance output.
- `Shape.resolvedFill` / `resolvedLine` (`src/read/api/shapes.ts:586,597`) and
  `Slide.themeContext()` (`src/read/api/theme-context.ts`) — de-circularizes the
  "Theme colour resolution" suite in `style-accessors.test.js`, which today only
  round-trips the default Office theme through the writer.
- `Shape.lineWidthPt` (`src/read/api/shapes.ts:522`).
- Backlog: the replication colour-transform blocker (audit "blocker #2", now
  implemented) and `dn-readmodel-style-followups` (base resolver shipped; the
  `fillRef`/`lnRef` + placeholder-inherited follow-up is still deferred — this
  fixture pins the part that *is* implemented and sets up the follow-up).

**Why a real fixture matters here:** PowerPoint applies `lumMod`/`lumOff`/`shade`/
`tint`/`satMod` with its own HSL rounding. The oracle table encodes a few
hand-checked values with a ±1/channel slack; a genuine deck lets the test read
the base `<a:schemeClr val="…">` + transform list *and* compare against the
exact effective color PowerPoint rendered, with no hand transcription.

**Shape of the fixture**
- One slide, deliberate `960×540` size, blank layout.
- Use a **non-default theme** (apply a built-in PowerPoint theme with distinctive
  accent colors, or edit the theme's `clrScheme`) so resolution can't accidentally
  pass by matching the Office default.
- A row of named rectangles, each **filled with a scheme color carrying a
  distinct transform**, covering the transform kinds the resolver supports:
  - `accent1` plain (no transform) — baseline.
  - `accent1` `lumMod 60000` + `lumOff 40000` (the common "lighter 40%").
  - `accent2` `lumMod 75000` (PowerPoint "darker 25%").
  - `accent3` `shade 50000`.
  - `accent4` `tint 40000`.
  - `bg2` `lumMod 20000` + `lumOff 80000`.
- One rectangle with an **explicit line**: scheme-color stroke at a known width
  (e.g. `2 pt` → `<a:ln w="25400">`) to pin `lineWidthPt` and `resolvedLine`.
- One text box whose **run color is a scheme color** (e.g. `accent5`) to pin
  `Run.resolvedColor`.

**Must contain in slide XML:** `<a:solidFill><a:schemeClr val="accentN"><a:lumMod/>
<a:lumOff/>…</a:schemeClr></a:solidFill>`, an `<a:ln w="…"><a:solidFill>
<a:schemeClr/></a:ln>`, and a non-default `ppt/theme/theme1.xml` `clrScheme`.
Name the shapes stably (e.g. `accent1-lm60-lo40`) so tests select by name.

**Verification:** open clean in desktop PowerPoint; record each shape's
PowerPoint-rendered fill hex (eyedropper / "More Colors") into the test as the
oracle, so the assertion compares the read+transform result to what PowerPoint
actually shows.

---

## Fixture 2 — `gradient-fill.pptx`

**Drives / pins**
- `Shape.gradientStops` (`src/read/api/shapes.ts:554`) — replaces the
  write→read round-trip in `style-accessors.test.js` ("gradientStops reads gsLst
  stops…") with a real `<a:gradFill><a:gsLst>` authored by PowerPoint.
- Backlog: `upstream-pr-1454` / `upstream-pr-1295` (gradient fill support,
  implemented) gain a *read-side* real-PowerPoint companion to the existing
  generation schema fixtures. Construct key `grad-fill` in the replication
  capability-map.

**Why a real fixture matters here:** PowerPoint writes `gsLst` stop positions in
1000ths of a percent and orders/normalizes stops its own way; reading its real
output (rather than our writer's) confirms the position-unit + color-split parse.

**Shape of the fixture**
- One slide, blank layout, deliberate size.
- Three named rectangles:
  - **Linear gradient**, 2 stops, `srgbClr` stops at 0% and 100% with a non-zero
    angle (`<a:lin ang=… scaled="1"/>`).
  - **Linear gradient**, 3 stops, mixing a `schemeClr` stop with `srgbClr` stops
    (exercises the color-split + transform-carry on a gradient stop).
  - **Radial/path gradient** (`<a:path path="circle">`) — confirms the reader
    doesn't assume linear.
- One solid-filled rectangle as the negative control (`gradientStops === null`).

**Must contain:** `<a:gradFill><a:gsLst><a:gs pos="…"><a:srgbClr|schemeClr/>…
</a:gsLst><a:lin|a:path/></a:gradFill>` with at least one multi-stop and one
scheme-color stop. Stable shape names (`grad-linear-2`, `grad-linear-3-scheme`,
`grad-radial`, `solid-control`).

**Verification:** open clean in desktop PowerPoint; assert stop count, the
positions in percent, and the per-stop color split (srgb hex vs scheme token +
transforms), plus `null` for the solid control.

---

## Fixture 3 — `preset-geometry.pptx`

**Drives / pins**
- `Shape.adjustValues` (`src/read/api/shapes.ts:533`) — replaces the
  `roundRect` round-trip in `style-accessors.test.js` ("adjustValues exposes a
  roundRect rectRadius as the avLst adj handle") with real PowerPoint `avLst`
  output.
- Backlog: relates to `shape-presets` regression coverage and the replication
  `cust-geom` / preset-adjust constructs; supports faithful "source-pixel replica"
  geometry transcription.

**Why a real fixture matters here:** PowerPoint names adjust handles per preset
(`adj`, `adj1`/`adj2`, …) and stores guide values in its own units (fractions of
shape size in 1000ths). Our writer may pick different defaults; reading
PowerPoint's actual `<a:avLst><a:gd name=… fmla="val …"/>` is the only way to pin
the handle names/values consumers depend on.

**Shape of the fixture**
- One slide, blank layout, deliberate size.
- Named autoshapes, each with a **non-default adjust handle dragged** in
  PowerPoint so an explicit `avLst` is written:
  - `roundRect` with a custom corner radius (single `adj`).
  - `chevron` or `homePlate` (single `adj`, arrow depth).
  - `roundRect`-family or `snip`/`bevel` shape with **two handles** (`adj1`,
    `adj2`) to exercise the multi-handle map.
  - A plain `rect` as the negative control (`adjustValues === {}`).

**Must contain:** `<p:sp><p:spPr><a:prstGeom prst="roundRect"><a:avLst>
<a:gd name="adj" fmla="val …"/></a:avLst></a:prstGeom>` for each adjusted shape,
and a `prst="rect"` with no `avLst`. Stable shape names matching the preset.

**Verification:** open clean in desktop PowerPoint; assert the handle name set and
values per shape, and `{}` for the plain rect.

---

## Fixture 4 — `multi-theme.pptx` (medium priority)

**Drives / pins**
- `Presentation.importSlide(source, index, { theme: 'preserve' })`
  (`src/read/api/presentation.ts`) — `import-slide-preserve.test.js` currently
  leans on `mixed.pptx` plus inline XML and the *default* theme. A deck whose
  slides genuinely resolve against a **non-default** theme is what proves
  `preserve` bakes scheme colors + style-matrix fills to the *correct* literal
  `srgbClr` (i.e. matching the source render, not the destination theme).
- Backlog: `dn-importslide-restyle-literals`, `dn-importslide-v1-limits` (theme
  handling), and `import-slide-restyle.test.js` (restyle leaving refs symbolic).

**Why a real fixture matters here:** "preserve flattens the source theme to
literals" is only meaningfully tested when source and destination themes differ;
with the default theme on both sides a no-op would pass. A real non-default theme
deck makes the before/after literal hex observable.

**Shape of the fixture**
- One source slide using shapes filled with **scheme colors and style-matrix
  fill references** (`<p:style><a:fillRef idx=…><a:schemeClr/></a:fillRef>`), on a
  master/layout/theme whose `clrScheme` is clearly **not** the Office default
  (distinctive accent hexes).
- Keep it small: 2–3 shapes is enough; the value is the distinctive theme, not
  shape variety.

**Must contain:** a `ppt/theme/theme1.xml` with a non-default `clrScheme`, and
slide shapes referencing it via `schemeClr` and `fillRef`/`lnRef` so a
`preserve` import has symbolic refs to flatten and a `restyle` import has refs to
rebind.

**Verification:** open clean in desktop PowerPoint; import its slide into a blank
default-theme deck with `theme: 'preserve'` and assert the emitted slide XML
carries literal `srgbClr` equal to the source theme's resolved hexes (and that
`restyle` leaves `schemeClr` symbolic). Note: this fixture also exercises the
deferred `style-fill-ref` resolver follow-up once that ships.

---

## Fixture 5 — `group-transform.pptx` extensions (additional cases)

**Drives / pins**
- `Shape.absoluteFrame` rot/flip composition (`src/read/api/shapes.ts:413`) for
  `dn-group-rot-flip-frame` — beyond the pure-rotation/flip cases the current
  fixture already pins.

**State of the existing fixture (verified 2026-06-18).** `group-transform.pptx`
ships a grouped slide (slide 1) and its PowerPoint **ungrouped twin** (slide 2),
paired by name (`"<grp> child <kind>"` ↔ `"<grp>-ungrouped child <kind>"`). Four
groups — `rot30` (30°), `flipH`, `flipV`, `rot330-flipHV` (330° + flipH + flipV) —
each with three children. A closed-form composition matches PowerPoint's baked
ungroup coordinates to **0 EMU** on all 12 children, so this is exact
ground truth and the test tolerance can be tight (a few EMU for rounding).

**Coverage gaps in the current fixture (the additional requirements).** Every
group is authored **1:1** (`chOff == off`, `chExt == ext`) and **no child carries
its own `@rot`/`@flipH`/`@flipV`**, so three interactions are unverified against
real PowerPoint and must be added (as new groups on the existing
grouped+ungrouped-twin slides, or a revised deck):

- **Scale + rotation composed** — a rotated group whose `chOff/chExt` differ from
  `off/ext` (scale ≠ 1). Pins that rot/flip is applied about the group center
  *after* the existing offset+scale mapping, not instead of it. (The scale path
  alone is covered by `mixed.pptx`; its interaction with rotation is not.)
- **Child-own rot/flip + group rot/flip** — a child with its own `@rot` (e.g.
  20°) and a child with its own `@flipH` placed inside a rotated/flipped group.
  Pins the composition rule (child rot **+** group rot; child flip **XOR** group
  flip) beyond the zero case the current children exercise.
- **Nested rotated group** — a rotated group nested inside a scaled (and/or
  rotated) group. Pins composition order through the multi-level ancestor walk
  (`mixed.pptx` covers nested *scaled* groups only).

**Shape of the additions.** Keep the grouped(slide 1) + ungrouped-twin(slide 2)
pattern and the `<name>` / `<name>-ungrouped` pairing convention so the test
compares against PowerPoint's own baked `off/ext/rot/flipH/flipV` at ~0 EMU. Give
each new group/child a stable, descriptive name (e.g. `scale-rot`,
`childrot-in-rot`, `nested-rot-in-scale`).

**Verification:** author in desktop PowerPoint, duplicate the slide and Ungroup
all (twice for nested groups), save; confirm slide 2 has zero `p:grpSp` and a
named flattened child per grouped child; re-run the 0-EMU closed-form check before
wiring the test.

## Not needed (already covered or out of scope)
- **Charts** — `mixed.pptx` carries a real chart and `chart.test.js` reads it; no
  minimal-chart gap is recorded in the backlog.
- **Generation-only constructs** (`bodypr-vert`/`dn-text-direction-serialization`,
  `normautofit-shrink`/`fit:'shrink'`, gradient *line* stroke
  `dn-gradient-line-stroke`, connector presets) — these are pinned by
  serialization schema fixtures in `test/schema.test.js` that assert
  "option X emits attribute Y". Per `dn-doc-render-caveats` they are deliberately
  XML-contract assertions, not read-model fixtures, so no PowerPoint `.pptx` is
  needed.

## Suggested order

1. `theme-colors.pptx` — unblocks the most circular/hand-oracle evidence
   (color transforms + theme resolution + line width in one deck).
2. `gradient-fill.pptx` — removes the gradient write→read round-trip.
3. `preset-geometry.pptx` — removes the adjust-handle write→read round-trip.
4. `multi-theme.pptx` — strengthens `importSlide` preserve/restyle once a
   distinctive second theme is available.
5. `group-transform.pptx` extensions — add the scale+rotation, child-own-rot,
   and nested-rotated-group cases; lower priority since the pure rot/flip cases
   are already pinned to 0 EMU.

For each: author with the `powerpoint-fixture-authoring` skill, run
`scripts/verify-powerpoint-fixture.ps1`, confirm the target OOXML construct is
present, then update `test/read/fixtures/README.md` (provenance table, purpose,
SHA-256, desktop-PowerPoint check date) and switch the corresponding read test
from its round-trip/oracle path to reading the new fixture.
