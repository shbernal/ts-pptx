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
| `multi-theme.pptx` | yes | yes (slide 1) | `fix(read): resolve p:style fillRef/lnRef` (59ea7bcf) | ✅ done (fillRef/lnRef); restyle-literals still deferred |
| `multi-theme.pptx` slide 2 (placeholder-inherited) | yes (2026-06-19) | ⏳ pending fix | ⏳ `Run.resolvedColor` fallback — Linux | 🟡 fixture done, fix + test remain (Remaining Work A) |
| `group-transform.pptx` (extended) | yes | yes | `fix(read): compose group rotations in absolute frames` (6871b337) | ✅ done |
| `rotation-flip.pptx` (per-shape rot/flip) | yes (2026-06-19) | yes | n/a (read accessor already shipped) | ✅ done (Remaining Work C) |

Fixtures were added in `4b1293c3 addition of new fixtures`; all five carry a
non-default **Ion** theme / PowerPoint-authored constructs and passed a Windows
desktop PowerPoint open-clean check on 2026-06-18 (README provenance + SHA-256
recorded). `multi-theme.pptx` then gained a placeholder-inherited slide 2 on
2026-06-19 (same Windows desktop PowerPoint COM path; slide 1 left byte-identical).

The original problem this set out to fix — read-model accessors validated only by
**round-tripping through PptxGenJS's own writer** (circular evidence) or
**hand-typed oracle tables** — is resolved for color transforms, theme
resolution, gradient stops, preset-geometry adjusts, style-matrix fill/line, and
per-shape rotation/flip (`rotation-flip.pptx`, 2026-06-19 — Remaining Work C). No
read-model accessor is now validated only through PptxGenJS's own writer.

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

### A. Placeholder-inherited run colour in the read getter — ✅ fixture landed; fix still needed (Linux)

`dn-readmodel-style-followups` had two halves: (1) walk `p:style` fillRef/lnRef,
and (2) resolve a **placeholder-inherited run colour** (a run whose `a:rPr`
carries no own color, inheriting from the layout/master placeholder). Half (1)
shipped via 59ea7bcf + `multi-theme.pptx`. Half (2) is **not** in the read getter:
`placeholderInheritedColor` exists in `src/read/oxml/theme.ts` only on the
`importSlide` flatten path, and `Run.resolvedColor` does not consult it.

- **Fixture — ✅ done (2026-06-19, Windows desktop PowerPoint COM).** Added as
  **slide 2 of `multi-theme.pptx`** (the recommended option — reuses the Ion theme
  + master). Slide 1's XML is byte-identical to before, so existing
  importSlide/style-matrix consumers (all read slide index 0) are unaffected;
  `pnpm run test:read` stays green (218 passed / 27 skipped). Two stably named
  placeholders on slide 2:
  - **`inherited-title`** (`<p:ph type="title"/>`): run is `<a:r><a:rPr
    lang="en-US"/><a:t>Inherited title color</a:t></a:r>` — **no own colour**.
    Both the layout (`slideLayout18.xml`) and master title placeholders carry an
    empty `<a:lstStyle/>`, so the colour resolves through the master `p:txStyles`
    `titleStyle` lvl1 → `<a:schemeClr val="tx2"/>` → clrMap `tx2="lt2"` → theme
    `lt2 = EBEBEB`. **Oracle: `resolvedColor.hex === "EBEBEB"`.**
  - **`explicit-body`** (`<p:ph type="body" idx="1"/>`): run carries an explicit
    `<a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>` — the negative control.
    **Oracle: `resolvedColor.hex === "FF00FF"` (explicit wins over the inherited
    `tx1 → lt1 → FFFFFF`).**
  - Fixture SHA-256 updated in `test/read/fixtures/README.md` to
    `737a28fa9832a1d009dc4588a868f856ec58c333843ba58f8eee3915a38cc659`.
- **Fix + test — ⏳ remaining (run on the Linux + tooling machine).** Extend
  `Run.resolvedColor` (`src/read/api/text.ts` / `theme-context.ts`) to fall back to
  the placeholder-inherited colour when the run, its paragraph, and the slide-level
  list style define none — reusing the `placeholderInheritedColor` chain logic
  (layout placeholder `lstStyle` → master placeholder `lstStyle` → master
  `p:txStyles` category style, resolved through the clrMap + theme). Then wire a
  `style-accessors.test.js` case reading `multi-theme.pptx` slide 2: assert
  `inherited-title` resolves to `EBEBEB` and `explicit-body` stays `FF00FF`.
- **Backlog:** `dn-readmodel-style-followups` is `partially-implemented` (updated
  2026-06-19) — fillRef/lnRef shipped (59ea7bcf), placeholder-inherited run colour
  is the remaining open leg (`next_action:
  layer-placeholder-inherited-run-colour`, stopgap retained). The fixture base is
  now in place (`multi-theme.pptx` slide 2); only the getter fix + test remain.

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

### C. Per-shape rotation/flip — ✅ done (2026-06-19)

`style-accessors.test.js` "Per-shape rotation / flip" reads `Shape.rotation`/
`flipH`/`flipV` mostly from synthetic `spWithXfrm` XML using **real PowerPoint
angle values** (e.g. `rot="2259366"` → 37.6561°, the unsigned `19216344` →
320.2724°). The one remaining write→read round-trip (`addShape({ rotate: 45,
flipH: true })` → reopen) is now **de-circularized**: a dedicated
**`rotation-flip.pptx`** fixture (desktop PowerPoint COM, 2026-06-19) carries two
ungrouped, stably-named rectangles — `rotated-45` (`<a:xfrm rot="2700000">` =
45°, no flip) and `flipped-h` (`<a:xfrm flipH="1">`, no rotation). The test now
reads those shapes instead of round-tripping through the writer; `pnpm run
test:read` stays **245** (218 passed / 27 skipped) on Windows.

A standalone fixture was used rather than mutating `theme-colors.pptx` /
`preset-geometry.pptx`: re-saving an existing committed fixture through PowerPoint
COM re-emits the whole package and could perturb the avLst/theme constructs those
fixtures already pin, so a fresh minimal deck is the conservative choice.

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

## Placeholder & notes serialization fixtures — oracles authored ✅; fixes remain (Linux)

Three `target-candidate` backlog items were **blocked on a PowerPoint-authored
oracle**: each is a write-side placeholder/notes/table-placeholder behaviour where
the fix is "emit the XML PowerPoint authors," and we must not guess that XML. They
follow the *authoring-oracle + serialization schema fixture* pattern (same as the
generation items): author a minimal desktop-PowerPoint deck with the
`powerpoint-fixture-authoring` skill, extract the relevant part, then pin the
writer output with a `test/schema.test.js` fixture.

**Oracle status (2026-06-19): all three oracle decks authored on Windows desktop
PowerPoint COM and opened clean (no repair).** Provenance + SHA-256 recorded in
`test/read/fixtures/README.md` under "Authoring oracles". They are inspection
oracles, **not** `test:read` fixtures. The remaining work for each (the writer fix
+ a `test/schema.test.js` assertion comparing the emitted XML against the oracle)
is the Linux + tooling job; the authentic target XML is now captured below so the
fix can be implemented and pinned without guessing.

| Item | Oracle deck | Construct captured | Remaining (Linux) |
|---|---|---|---|
| E | `layout-placeholder-bodypr.pptx` | layout placeholder `<a:bodyPr>` insets+anchor | thread bodyPr through placeholder emit + schema fixture |
| F | `table-placeholder.pptx` | table graphicFrame `<p:ph idx="1"/>` | add table `placeholder` option + schema fixture |
| G | `notes-slide-image.pptx` | notes `<p:ph type="sldImg">` geometry | align emitted notes `sldImg` + schema fixture |

### E. Master/layout placeholder margin + valign (`upstream-pr-1247`, `upstream-issue-1208`) — ✅ oracle authored; fix remains (Linux)

`genXmlBodyProperties` applies body properties (margin, `anchor`/valign, custom
bullets) only to ordinary text objects, not to placeholder objects on masters /
layouts, so a placeholder authored with a margin or vertical anchor may not carry
it the way PowerPoint expects when a user inserts a new slide from that layout
(`#1208`). `#1247` is the broader master-placeholder formatting item (margin,
valign, custom bullet).

- **Oracle — ✅ done (`layout-placeholder-bodypr.pptx`, 2026-06-19).** The "Title
  and Content" layout (`slideLayout2.xml`) was edited via COM (`TextFrame2`
  VerticalAnchor + Margin*) so its placeholders carry explicit `<a:bodyPr>`:
  - **`oracle-title-ph`** (`<p:ph type="title"/>`):
    `<a:bodyPr lIns="228600" tIns="114300" rIns="228600" bIns="114300" anchor="b"/>`
    (18pt/9pt insets, bottom-anchored).
  - **`oracle-body-ph`** (`<p:ph idx="1"/>`):
    `<a:bodyPr lIns="304800" tIns="190500" rIns="152400" bIns="76200" anchor="ctr"/>`
    (asymmetric 24/15/12/6pt insets, middle-anchored — each inset independently
    pinned).
  - Slide 1, inserted from the layout, carries an empty `<a:bodyPr/>` on each
    placeholder (the inheritance path the fix must honour).
- **Fix + fixture — ⏳ remaining (Linux).** Thread body properties through the
  placeholder emit path in `genXmlBodyProperties`; add a `test/schema.test.js`
  fixture asserting the emitted layout placeholder `<a:bodyPr anchor=… lIns=…>`
  matches the authored shape above. Blocks `upstream-pr-1247` and
  `upstream-issue-1208`.

### F. Table placeholder properties (`upstream-pr-1151`) — ✅ oracle authored; fix remains (Linux)

Table objects don't expose placeholder behaviour the way image/text objects do.
Before adding a `placeholder` option to table props we need the exact
`<p:graphicFrame>` + `<p:nvGraphicFramePr>/<p:nvPr><p:ph …>` shape PowerPoint
authors for a table that lives in a layout placeholder.

- **Oracle — ✅ done (`table-placeholder.pptx`, 2026-06-19).** One slide from the
  "Title and Content" layout whose content placeholder hosts a 2×3 table. Authored
  via COM by calling `Shapes.AddTable` over the empty content placeholder's exact
  geometry — PowerPoint binds the table into the placeholder, emitting the
  graphicFrame wiring:
  ```xml
  <p:graphicFrame>
    <p:nvGraphicFramePr>
      <p:cNvPr id="4" name="placeholder-table">…</p:cNvPr>
      <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
      <p:nvPr><p:ph idx="1"/>…</p:nvPr>
    </p:nvGraphicFramePr>
    <p:xfrm>…</p:xfrm>
    <a:graphic><a:graphicData uri="…/table"><a:tbl>…</a:tbl></a:graphicData></a:graphic>
  </p:graphicFrame>
  ```
  Key point: the placeholder binding is `<p:ph idx="1"/>` (content-placeholder
  idx, **no `type`**) on the graphicFrame's `nvPr`, independent of the contained
  graphic. (`mixed.pptx` slide 2 corroborates the same wiring for a non-table
  content graphic.)
- **Fix + fixture — ⏳ remaining (Linux).** Add the table `placeholder` option and a
  `test/schema.test.js` fixture asserting the emitted `p:ph` on the table
  graphicFrame matches the wiring above. Blocks `upstream-pr-1151`.

### G. Notes print-layout slide image placeholder (`upstream-issue-446`) — ✅ oracle authored; fix remains (Linux)

The notes slide / notesMaster `sldImg` placeholder is currently hard-coded and
unverified against what PowerPoint's notes print layout expects, so generated decks
reportedly don't show the slide image in notes print view.

- **Oracle — ✅ done (`notes-slide-image.pptx`, 2026-06-19).** One slide with
  speaker notes (notes text set via COM to materialize the notes parts):
  - **`notesMaster1.xml`** carries the geometry-bearing placeholder:
    `<p:ph type="sldImg" idx="2"/>` with `<a:xfrm><a:off x="685800" y="1143000"/>
    <a:ext cx="5486400" cy="3086100"/></a:xfrm>`, `<a:prstGeom prst="rect">`,
    `<a:noFill/>`, and `<a:ln w="12700"><a:solidFill><a:prstClr val="black"/>` (1pt
    black border). `spLocks` = `noGrp/noRot/noChangeAspect`.
  - **`notesSlide1.xml`** carries a bare `<p:ph type="sldImg"/>` with empty
    `<p:spPr/>` — it inherits the master geometry (this is the inheritance the fix
    must reproduce).
- **Fix + fixture — ⏳ remaining (Linux).** Align the emitted notes `sldImg`
  placeholder (notesMaster geometry + bare notesSlide placeholder) to the authored
  XML above; add a `test/schema.test.js` fixture. Blocks `upstream-issue-446`.

## Feature-serialization fixtures — authoring oracles needed (target-candidate, unscheduled)

Two further `target-candidate` write features emit OOXML we must not guess and so
follow the same *authoring-oracle + serialization schema fixture* pattern as E–G,
but are not yet scheduled (their backlog `next_action` is research/design, not an
oracle-then-fix step). They are recorded here so the oracle requirement is captured
before any implementation begins. **Do not implement either until its oracle
exists.** As with E–G these are inspection oracles, not `test:read` fixtures —
record provenance/SHA-256 in `test/read/fixtures/README.md` and pin the writer
output in `test/schema.test.js`.

### H. Bar-chart custom data labels + per-point colours (`upstream-pr-727`)

The chart writer already has per-series overrides (`seriesOverride.dataLabelColor`,
`gen-charts.ts`) and per-point `c:dPt` / per-point custom `c:dLbl` text, but the
full PR-727 surface (arbitrary per-point label text **and** per-point fill colour
together, kept consistent with the embedded workbook cache) is not landed, and the
exact `c:dLbl`/`c:dPt` shape plus the `numCache`/`strCache` it must agree with is
the thing not to guess. This is a **write-side** chart oracle — distinct from the
read-side "Charts" note below, which only says `mixed.pptx` covers chart *reads*.

- **Oracle:** a PowerPoint deck with a bar chart whose individual bars were
  manually recoloured and given custom per-point data-label text, captured to pin
  the per-point `<c:dPt>` (idx + `spPr` fill) and `<c:dLbl>` (idx + rich `c:tx`)
  ordering inside `CT_BarSer`, plus the embedded workbook (`xl/embeddings`) cache
  values the labels are derived from.
- **Fix + fixture (after oracle):** extend the bar-series API for per-point label
  text/colour, keep the workbook cache consistent, and add a `test/schema.test.js`
  fixture asserting the emitted `c:dPt`/`c:dLbl` against the authored shape. Blocks
  `upstream-pr-727`.

### I. Native math equation (OMML) text run (`upstream-issue-1456`)

PptxGenJS has math *symbol shapes* (`mathPlus`/`mathEqual`/… preset geometries) but
no editable-equation model: no `m:oMathPara`/`m:oMath` run inside a text body, no
`a14:m` alternate-content wrapper, and no LaTeX/MathML→OMML path. The OMML PowerPoint
authors for an inserted equation (and the `mc:AlternateContent` / `a14:m` wrapper it
sits in within `<a:p>`) is exactly what must not be guessed.

- **Oracle:** a PowerPoint deck with an inserted equation in a text box, captured to
  pin the `<a:p>` → `<a14:m>`/`<m:oMathPara><m:oMath>` structure, the `m:` namespace
  declarations, and the run-property handling PowerPoint emits.
- **Fix + fixture (after oracle):** introduce an OMML text-run model (raw OMML first,
  MathML/LaTeX conversion later), then add a `test/schema.test.js` fixture pinning
  the emitted math run against the authored XML. Blocks `upstream-issue-1456`. The
  oracle is also the natural first deliverable of the entry's
  `research-omml-in-presentation-text-runs` step.

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
