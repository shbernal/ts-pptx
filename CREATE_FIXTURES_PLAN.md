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

All write-side **serialization oracles** are now authored too (2026-06-19, Windows
desktop PowerPoint COM): the placeholder/notes set E/F/G
(`layout-placeholder-bodypr`, `table-placeholder`, `notes-slide-image`) and the
feature set H/I (`bar-chart-data-labels`, `math-omml`). Every fixture this plan
calls for on the Windows side now exists; the remaining items (read getter fix A,
and the writer fixes E–I) are Linux + tooling work with the authentic target XML
already captured.

| Fixture | Added | Tests wired | Fix landed | State |
|---|---|---|---|---|
| `theme-colors.pptx` | yes | yes | n/a (accessor already shipped) | ✅ done |
| `gradient-fill.pptx` | yes | yes | n/a | ✅ done |
| `preset-geometry.pptx` | yes | yes | n/a | ✅ done |
| `multi-theme.pptx` | yes | yes (slide 1) | `fix(read): resolve p:style fillRef/lnRef` (59ea7bcf) | ✅ done (fillRef/lnRef); restyle-literals still deferred |
| `multi-theme.pptx` slide 2 (placeholder-inherited) | yes (2026-06-19) | yes | `Run.resolvedColor` placeholder-inherited fallback (2026-06-19) | ✅ done (Remaining Work A) |
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

### A. Placeholder-inherited run colour in the read getter — ✅ done (2026-06-19)

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
- **Fix + test — ✅ done (2026-06-19).** `Run.resolvedColor` now falls back, when
  the run sets no own colour, to the colour it inherits — paragraph `a:defRPr`, then
  the slide text body `a:lstStyle`, then the placeholder's layout → master
  `a:lstStyle` → master `p:txStyles` chain (`resolveInheritedRunColor` in
  `theme-context.ts`, built on the new exported `placeholderInheritedFill` /
  `lstStyleLevelFill` helpers the flatten path already used). `resolveSlideColorContext`
  now carries the layout/master roots, and `AutoShape.textFrame` threads the shape's
  `p:ph` identity through `TextFrame → Paragraph → Run` (a memoized per-paragraph
  thunk does the lookup lazily, only for colourless placeholder runs).
  `style-accessors.test.js` "Placeholder-inherited run colour" reads
  `multi-theme.pptx` slide 2: `inherited-title` → `EBEBEB`, `explicit-body` →
  `FF00FF`. `pnpm run test:read` is **247** (was 245).
- **Backlog:** `dn-readmodel-style-followups` is now `implemented` (2026-06-19) —
  both legs (fillRef/lnRef via 59ea7bcf and placeholder-inherited run colour) shipped;
  `next_action: none`, the downstream stopgap dropped.

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

## Placeholder & notes serialization fixtures — ✅ done (oracles authored + fixes landed 2026-06-19)

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
| E | `layout-placeholder-bodypr.pptx` | layout placeholder `<a:bodyPr>` insets+anchor | ✅ done (2026-06-19) — bodyPr threaded through placeholder emit + schema fixture |
| F | `table-placeholder.pptx` | table graphicFrame `<p:ph idx="1"/>` | ✅ done (2026-06-19) — table `placeholder` option + schema fixture |
| G | `notes-slide-image.pptx` | notes `<p:ph type="sldImg">` geometry | ✅ done (2026-06-19) — output already byte-identical to oracle; locked by schema fixture |

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
- **Fix + fixture — ✅ done (2026-06-19).** `genXmlBodyProperties`
  (`src/gen-xml.ts`) now emits the configured `<a:bodyPr>` for placeholder objects,
  not only text objects (the `_type === text` guard was widened to include
  `SLIDE_OBJECT_TYPES.placeholder`), so a master/layout placeholder authored with
  `margin` (→ `lIns/tIns/rIns/bIns`) and/or `valign` (→ `anchor`) carries them.
  Schema fixture "master/layout placeholder carries bodyPr insets + anchor"
  (`test/schema.test.js`) pins the title (bottom-anchored 18/9pt) and body
  (center-anchored 24/15/12/6pt) placeholders against the oracle, validator-clean.
  `upstream-pr-1247` and `upstream-issue-1208` are `implemented`.

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
- **Fix + fixture — ✅ done (2026-06-19).** `TableProps` gained a
  `placeholder?: string` option; the table `<p:graphicFrame>` emits the bound
  layout placeholder's `<p:ph>` on `<p:nvPr>` (before `<p:extLst>`) via
  `genXmlPlaceholder`, and `addTableDefinition` inherits the placeholder geometry
  for any omitted x/y/w/h. Schema fixture "table bound to a layout placeholder
  emits p:ph on the graphicFrame" (`test/schema.test.js`) pins it against the
  oracle, validator-clean. `upstream-pr-1151` is `implemented`.

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
- **Fix + fixture — ✅ done (2026-06-19).** Verification showed the writer output is
  **already byte-identical** to the oracle: `makeXmlNotesMaster` emits exactly
  `<p:ph type="sldImg" idx="2"/>` with `off 685800,1143000 / ext 5486400,3086100`,
  `prstGeom rect`, `noFill`, and the 1pt black `<a:ln>`, and `makeXmlNotesSlide`
  emits the bare `<p:ph type="sldImg"/>` with empty `<p:spPr/>`. No code change was
  needed — the previously-unverified hard-coding is correct. Locked by the schema
  fixture "notes sldImg placeholder geometry (notesMaster) + bare placeholder
  (notesSlide)" (`test/schema.test.js`). `upstream-issue-446` is `implemented`.

## Feature-serialization fixtures — ✅ done (oracles authored + fixes landed 2026-06-19; OMML raw-leg, LaTeX/raster future)

Two further `target-candidate` write features emit OOXML we must not guess and so
follow the same *authoring-oracle + serialization schema fixture* pattern as E–G.
Both oracles are now **authored** (2026-06-19, Windows desktop PowerPoint COM;
provenance + SHA-256 in `test/read/fixtures/README.md`), so the
"do not implement until the oracle exists" precondition is satisfied. As with E–G
these are inspection oracles, not `test:read` fixtures — the remaining work for
each (the writer fix + a `test/schema.test.js` assertion comparing the emitted XML
against the oracle) is the Linux + tooling job; the authentic target XML is now
captured so the fix can be implemented and pinned without guessing.

| Item | Oracle deck | Construct captured | Remaining (Linux) |
|---|---|---|---|
| H | `bar-chart-data-labels.pptx` | per-point `c:dPt` fills + custom `c:dLbl` text + workbook cache | ✅ done (2026-06-19) — already supported via `pointStyles`/`customLabels`; locked by schema fixture |
| I | `math-omml.pptx` | `<a14:m>`/`<m:oMathPara>`/`<m:oMath>` equation run + `mc:Fallback` | ✅ raw-OMML done (2026-06-19) — `math` text prop + a14 envelope + schema fixture; LaTeX/MathML + raster fallback still future |

### H. Bar-chart custom data labels + per-point colours (`upstream-pr-727`) — ✅ oracle authored; fix remains (Linux)

The chart writer already has per-series overrides (`seriesOverride.dataLabelColor`,
`gen-charts.ts`) and per-point `c:dPt` / per-point custom `c:dLbl` text, but the
full PR-727 surface (arbitrary per-point label text **and** per-point fill colour
together, kept consistent with the embedded workbook cache) is not landed, and the
exact `c:dLbl`/`c:dPt` shape plus the `numCache`/`strCache` it must agree with is
the thing not to guess. This is a **write-side** chart oracle — distinct from the
read-side "Charts" note below, which only says `mixed.pptx` covers chart *reads*.

- **Oracle — ✅ done (`bar-chart-data-labels.pptx`, 2026-06-19).** One slide with a
  clustered **column** chart (`bar-chart-727`, `<c:barDir val="col"/>` → CT_BarSer),
  one series of four points whose bars were individually recoloured and given
  custom per-point label text via COM. `ppt/charts/chart1.xml` carries:
  - four `<c:dPt>` (idx 0–3), each `<c:spPr>` a solid fill —
    `FF0000`/`00B050`/`0070C0`/`FFC000`;
  - four `<c:dLbl>` (idx 0–3), each a rich `<c:tx>` overriding the numeric value
    with `Low`/`Mid`/`High`/`Peak`;
  - the embedded-workbook cache the labels must agree with: strCache categories
    `Q1`–`Q4` (`Sheet1!$A$2:$A$5`), numCache values `10`/`25`/`18`/`30`
    (`Sheet1!$B$2:$B$5`), workbook in `ppt/embeddings/`.
- **Fix + fixture — ✅ done (2026-06-19).** Verification showed the full surface is
  already implemented and public: a series' `pointStyles[].fill` emits the per-point
  `<c:dPt><c:spPr>` solid fills (`makeSeriesDataPointsXml`) and `customLabels[]`
  emits the per-point rich `<c:dLbl>` text (`makeCustomDLblXml`), combining in correct
  CT_BarSer order while the value cache keeps the real numbers. No code change was
  needed. Schema fixture "bar chart per-point colours + custom data labels
  (consistent with value cache)" (`test/schema.test.js`) pins all four `c:dPt` fills,
  the four `c:dLbl` texts, and the intact numCache against the oracle, validator-clean.
  `upstream-pr-727` is `implemented`.

### I. Native math equation (OMML) text run (`upstream-issue-1456`) — ✅ oracle authored; fix remains (Linux)

PptxGenJS has math *symbol shapes* (`mathPlus`/`mathEqual`/… preset geometries) but
no editable-equation model: no `m:oMathPara`/`m:oMath` run inside a text body, no
`a14:m` alternate-content wrapper, and no LaTeX/MathML→OMML path. The OMML PowerPoint
authors for an inserted equation (and the `mc:AlternateContent` / `a14:m` wrapper it
sits in within `<a:p>`) is exactly what must not be guessed.

- **Oracle — ✅ done (`math-omml.pptx`, 2026-06-19).** One slide with a text box
  (`equation-box`) holding the equation 𝑥²+1=𝑦. `ppt/slides/slide1.xml` carries
  `<mc:AlternateContent>` → `<mc:Choice Requires="a14">` → `<a14:m>` →
  `<m:oMathPara>` → `<m:oMath>` with an `<m:sSup>` superscript, `Cambria Math` run
  properties, and the variables as Unicode mathematical-italic letters
  (`U+1D465`/`U+1D466`); plus the `<mc:Fallback>` raster `blipFill`
  (`ppt/media/image1.png`, `rId2`) PowerPoint emits for back-compat. PowerPoint has
  **no COM equation-insert API**, so the equation was built in Word (`OMaths.Add` +
  `BuildUp`), copied, and pasted into the PowerPoint text box; **PowerPoint
  re-serialised it on `SaveAs`, so the package XML is genuine PowerPoint output.**
- **Fix + fixture — ✅ raw-OMML leg done (2026-06-19).** A text item gained a
  `math?: string` raw-OMML property (`TextProps`). `genXmlTextBody` isolates it as a
  display-math paragraph and emits `<a14:m><m:oMathPara><m:oMath>{OMML}</m:oMath>…`
  via `genXmlMathParagraph` (accepting inner OMML, a full `<m:oMath>`, or a full
  `<m:oMathPara>`; both `a14` and `m` namespaces declared on `<a14:m>`). The
  equation-bearing shape is wrapped in `<mc:AlternateContent><mc:Choice
  Requires="a14">` exactly as the oracle does, so it validates clean. Schema fixture
  "native math equation (OMML) text run" (`test/schema.test.js`) pins the envelope +
  OMML structure. `upstream-issue-1456` is `partially-implemented`: the
  **remaining** legs are a LaTeX/MathML→OMML conversion path and the `mc:Fallback`
  raster (the raw-OMML entry point is the deliberate first deliverable).

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
