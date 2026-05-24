# PptxGenJS — bug fix queue

This is the source of truth for the autofix loop. Each `## B<n>:` section is
one bug. Process top to bottom. **Commit each fix directly to `master`.**
Do **not** create per-bug branches. Do **not** open PRs.

Project conventions:

- `README.md` — install / build commands and library overview.
- `CHANGELOG.md` — [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
  format. Append a `### Fixed` entry under an `## [Unreleased]` section
  when closing a fix.
- `TESTING.md` — manual cross-platform release validation. Out of scope
  per bug; run before cutting a release, not per fix.

For each bug:

1. Read `README.md` and `package.json` at the repo root to confirm the
   project's install/build/test commands. The npm scripts are
   `npm run build` (rollup) and `npm test` (custom runner in `test/run.js`).
2. Reproduce against current `master` using the snippet under "Reproduction."
   If you cannot reproduce, mark the bug `cannot-reproduce` in
   `.autoloop/progress.md` and move on — do not fabricate a fix.
3. Locate the emit site. The TS source lives in `src/`. Most XML emission is
   in `src/gen-xml.ts`; chart emission in `src/gen-charts.ts`; object intake /
   mutation in `src/gen-objects.ts`. Validate via grep, not memory.
4. Implement the **minimal** fix. No refactoring beyond what the bug requires.
5. Add a regression test under `test/` using the existing mocha pattern. The
   test must fail before the fix and pass after.
6. `npm install && npm run build && npm test`. All existing tests must stay
   green.
7. Update `CHANGELOG.md`: add a bullet under `### Fixed` in the
   `## [Unreleased]` section (create the section at the top of the file if
   it does not exist) referencing the bug ID and any upstream issue
   number(s).
8. **Commit the fix directly to `master`** with a Conventional Commits subject:
   `fix(<area>): <summary>`. Body should reference the bug ID and any
   upstream issue number(s).
9. Mark the bug as `closed` in `.autoloop/progress.md` and append a one-line
   entry to `.autoloop/fix-log.md` (`Bxx: <commit sha> <one-line summary>`).
10. **Continue to the next unprocessed bug.** Do not emit `task.complete`
    until every bug in this file is marked `closed` or `cannot-reproduce` in
    `.autoloop/progress.md`.

If a bug already has an open upstream PR that fixes it correctly, study the
PR for guidance, but write the fix as a fresh commit on `master` of this
fork (no PR review, no cherry-pick — we are not pushing to upstream).

Priority tiers:

- **P1 (file-corruption / "needs repair"):** B1, B2, B3, B6, B10, B11, B12, B13, B14, B15
- **P2 (silent data loss / wrong output):** B4, B5, B7, B8, B18, B19
- **P3 (cosmetic / non-breaking):** B9, B16, B17

Process P1 first, then P2, then P3.

Termination contract:

- Emit `task.complete` **only** when every `## B<n>:` heading in this file
  has a corresponding `closed` or `cannot-reproduce` entry in
  `.autoloop/progress.md`.
- Until then, after each `bug.closed`, hand back to the diagnoser to pick
  the next unprocessed bug.

---

## B1: Duplicate `<a:pPr>` when paragraph has mixed-formatting runs

**Priority:** P1  
**Related upstream:** gitbrent/PptxGenJS#1322

### Symptom

When a paragraph is generated from an `addText` call whose `text` array
mixes formatting (e.g. one bold run plus one regular run, or runs with
different `align` values), the emitted slide XML contains two `<a:pPr>`
elements inside the same `<a:p>`. Valid OOXML allows only one, at the
start of the paragraph. PowerPoint reports "needs repair" and drops
formatting on the affected paragraphs.

### Reproduction

```javascript
var pptxgen = require('pptxgenjs');
var pres = new pptxgen();
var s = pres.addSlide();
s.addText([
  { text: 'Bold ', options: { bold: true } },
  { text: 'and regular' }
], { x: 1, y: 1, w: 6, h: 1, fontSize: 24 });
pres.writeFile({ fileName: 'b1-repro.pptx' });
// Verify: unzip b1-repro.pptx, count <a:pPr> in ppt/slides/slide1.xml
// per <a:p>. Bug: paragraph has 2; fix: 1.
```

### Root-cause hypothesis

`src/gen-xml.ts` paragraph-emission helper emits `<a:pPr>` both at
paragraph open and again when a run with different paragraph properties
is appended. Should emit once at paragraph start, merging properties
from runs into it.

### Acceptance criteria

Every `<a:p>` in the output contains at most one `<a:pPr>` regardless
of how the runs in the paragraph mix bold/italic/align/etc.

### References

- Upstream issue (related, may be the same root cause): https://github.com/gitbrent/PptxGenJS/issues/1322

---

## B2: Phantom slideMaster overrides in `[Content_Types].xml`

**Priority:** P1  
**Related upstream:** gitbrent/PptxGenJS#1444, #1449 (Defect 3)

### Symptom

`[Content_Types].xml` contains a `<Override>` entry for
`/ppt/slideMasters/slideMasterN.xml` for every slide in the deck, but
only `slideMaster1.xml` actually exists. PowerPoint repair removes
those phantom Override entries.

A 55-slide deck with 3 defined masters produces 55 slideMaster
overrides — see #1449 for a full reproducer.

### Reproduction

```javascript
var pptxgen = require('pptxgenjs');
var pres = new pptxgen();
for (var i = 0; i < 5; i++) {
  var s = pres.addSlide();
  s.addText('Slide ' + (i + 1), { x: 1, y: 1, w: 6, h: 1 });
}
pres.writeFile({ fileName: 'b2-repro.pptx' });
// Verify: unzip, look at [Content_Types].xml. Bug: 5 slideMaster
// Override entries; fix: 1 (matching ppt/slideMasters/*.xml).
```

### Root-cause hypothesis

The Content_Types generation loop iterates over slides instead of over
the actual slide-master collection. Should iterate over the masters.

### Acceptance criteria

Every `<Override PartName="/ppt/slideMasters/...">` entry resolves to
an existing `slideMasterN.xml` part inside the archive, regardless of
slide count.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1444
- Upstream meta: https://github.com/gitbrent/PptxGenJS/issues/1449 (Defect 3)

---

## B3: Dangling `<Relationship>` entries in `*.rels`

**Priority:** P1

### Symptom

`*.rels` files contain `<Relationship>` entries whose `Target` points
to a part not present in the archive.

### Reproduction

Reproducer is configuration-dependent (depends on which features
trigger the orphaned relationship). Start by building a minimal deck
that exercises slide layouts, theme overrides, and notes, then dump
each `_rels/*.rels` and check that every `Target` resolves. If you
cannot find a deterministic repro on current master, mark
`cannot-reproduce` and move on — but check #1449 first for hints.

### Root-cause hypothesis

Same general pattern as B2 — relationships emitted speculatively
without checking whether the target part is being written.

### Acceptance criteria

Every `<Relationship Target="...">` in every `.rels` file resolves to
an existing part in the archive.

### References

- Upstream meta: https://github.com/gitbrent/PptxGenJS/issues/1449

---

## B4: `writeFile()` mutates option objects in place (pt → EMU)

**Priority:** P2  
**Related upstream:** gitbrent/PptxGenJS#1366, #1293

### Symptom

Calling `writeFile()` mutates internal option objects (notably
`shadow`) by converting `pt` units to `EMU` in place. A second
`writeFile()` call re-converts the already-converted values, producing
absurd numbers (e.g. `blurRad="967740000"` instead of `76200`).
PowerPoint silently strips elements with these values, dropping
content.

### Reproduction

```javascript
var pptxgen = require('pptxgenjs');
var pres = new pptxgen();
var s = pres.addSlide();
s.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 4, h: 1,
  fill: { color: 'FF0000' },
  shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
});
pres.writeFile({ fileName: 'b4-a.pptx' });
pres.writeFile({ fileName: 'b4-b.pptx' });
// Compare. Bug: b4-b has corrupted shadow values; fix: byte-equal.
```

### Root-cause hypothesis

pt→EMU conversion in `src/gen-objects.ts` (search for arithmetic on
`shadow.blur`, `shadow.offset`, `line.width`, etc.) mutates the user's
option object. Conversion must operate on a clone.

### Acceptance criteria

Building a `pres`, calling `writeFile()`, then calling `writeFile()`
again on the same `pres` produces two files whose post-write XML is
identical for every option-object field. Add a direct test for this
exact double-write scenario.

### References

- Related upstream: https://github.com/gitbrent/PptxGenJS/issues/1366, https://github.com/gitbrent/PptxGenJS/issues/1293

---

## B5: Reusing one options object across calls produces wrong values

**Priority:** P2 (likely subsumed by B4)

### Symptom

Same root cause as B4 (in-place mutation), observed at the user level:

```javascript
var shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 };
slide.addShape(pres.shapes.RECTANGLE, { shadow: shadow, /* ... */ });
slide.addShape(pres.shapes.RECTANGLE, { shadow: shadow, /* ... */ });
// Second shape gets a shadow built from already-converted values.
```

### Reproduction

The snippet above. Inspect `ppt/slides/slide1.xml` to confirm the
second shape's shadow uses corrupted EMU values.

### Root-cause hypothesis

At every `addX` entry point, mutable nested option objects (`shadow`,
`line`, `fill`, `glow`, etc.) are stored on the internal slide-data
structure by reference. Subsequent emit/conversion mutates them.
Implementing B4 correctly should fix this; verify and close together
or split into a separate clone-on-ingest PR if that is cleaner.

### Acceptance criteria

The snippet above produces two byte-equal shape XML blocks (modulo
position).

---

## B6: Combo charts emit too few axis defs vs IDs referenced

**Priority:** P1  
**Related upstream:** gitbrent/PptxGenJS#1355, #1448

### Symptom

Combo chart created with `secondaryValAxis: true` and
`secondaryCatAxis: true` flags but **without** `catAxes` and `valAxes`
arrays produces XML with 2 axis definitions but references 5+ axis
IDs. PowerPoint detects the mismatch and triggers "needs repair."

### Reproduction

See #1355 / #1448 for full reproducers. Minimal:

```javascript
// Create a combo chart using the secondary*Axis flags only.
// Inspect generated chart XML: count <c:catAx>/<c:valAx> defs vs
// distinct axId references in <c:plotArea>. They should match.
```

### Root-cause hypothesis

Combo-chart axis emission in `src/gen-charts.ts`. Two options:
(a) auto-synthesize the missing axis defs when only the secondary
flags are set, (b) throw at API time if a combo chart is requested
without the axis arrays. Option (a) is more user-friendly. Whichever
you choose, never produce a corrupt file.

### Acceptance criteria

Every `axId` referenced in `<c:plotArea>` resolves to a defined
`<c:catAx>` or `<c:valAx>`. If the user omits required axis arrays,
either auto-synthesize or throw a clear error.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1355, https://github.com/gitbrent/PptxGenJS/issues/1448

---

## B7: Hex color with `#` prefix produces a corrupt file

**Priority:** P2

### Symptom

`color: '#FF0000'` (instead of `'FF0000'`) silently produces a file
PowerPoint cannot open.

### Reproduction

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 2, h: 1, fill: { color: '#FF0000' }
});
```

### Root-cause hypothesis

Color validation/normalization helper does not strip a leading `#`.

### Acceptance criteria

`'#FF0000'` either works (preferred — strip `#` and emit `FF0000`) or
throws a clear, actionable error at API time. Never silently corrupts.

---

## B8: 8-character hex (RGB+alpha) silently corrupts the file

**Priority:** P2

### Symptom

`color: '00000020'` (intended as black with 12.5% opacity) corrupts
the output.

### Reproduction

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 2, h: 1,
  shadow: { type: 'outer', blur: 6, offset: 2, color: '00000020' }
});
```

### Root-cause hypothesis

Color validation accepts arbitrary hex length and emits as-is.

### Acceptance criteria

8-character hex either parses the alpha out and emits an `<a:alpha>`
sibling, or throws a clear error. Never silently corrupts.

---

## B9: Unicode bullet glyphs duplicate with `bullet: true`

**Priority:** P3

### Symptom

`addText('• item', { bullet: true })` renders two bullets — the literal
unicode glyph and the generated bullet.

### Reproduction

```javascript
slide.addText('• item', { x: 1, y: 1, w: 4, h: 0.5, bullet: true });
```

### Root-cause hypothesis

Text emission does not strip leading bullet glyphs when `bullet: true`.

### Acceptance criteria

`addText('• item', { bullet: true })` renders one bullet, not two.

---

## B10: `addShape("oval", ...)` writes invalid OOXML preset string

**Priority:** P1  
**Source:** gitbrent/PptxGenJS#1449 (Defect 1)

### Symptom

`slide.addShape("oval", ...)` writes `<a:prstGeom prst="oval">`. The
valid OOXML preset is `"ellipse"`. PowerPoint cannot parse `prst="oval"`
and removes the shape during repair.

`pres.shapes.OVAL` returns `"ellipse"`, so the enum-constant API is
correct — but the bare string `"oval"` is accepted and written
verbatim.

### Reproduction

```javascript
var pptxgen = require('pptxgenjs');
var pres = new pptxgen();
var s = pres.addSlide();
s.addShape("oval", {
  x: 1, y: 1, w: 0.4, h: 0.4, fill: { color: '00B0B9' }
});
pres.writeFile({ fileName: 'b10-repro.pptx' });
// Verify: unzip, grep 'prstGeom prst' ppt/slides/slide1.xml.
// Bug: prst="oval"; fix: prst="ellipse" or throw.
```

### Root-cause hypothesis

`addShape()` accepts an arbitrary string for the preset and writes it
verbatim into the XML. Should validate against the known presets (the
internal shapes enum) and either map common aliases (`oval` →
`ellipse`) or throw.

### Acceptance criteria

`addShape("oval", ...)` either produces `prst="ellipse"` (preferred —
maps the alias) or throws a clear error at API time. Never produces an
invalid `prst` string.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1449 (Defect 1)

---

## B11: `addShape("roundedRectangle", ...)` writes invalid OOXML preset string

**Priority:** P1  
**Source:** gitbrent/PptxGenJS#1449 (Defect 2)

### Symptom

Same class of bug as B10. The string `"roundedRectangle"` is written
verbatim; valid OOXML is `"roundRect"`.

### Reproduction

```javascript
slide.addShape("roundedRectangle", {
  x: 1, y: 1, w: 2, h: 1, rectRadius: 0.1
});
// Verify: prst="roundedRectangle" in slide XML; fix: prst="roundRect"
// or throw.
```

### Root-cause hypothesis

Same as B10 — no shape-string validation/aliasing in `addShape()`.

### Acceptance criteria

Same shape (heh) as B10. The fix for B10 should fix this too if
implemented as a general alias map. If it does, close B11 as fixed-by-B10
and reference both bugs in the PR.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1449 (Defect 2)

---

## B12: Solid color background missing `<a:effectLst/>` triggers repair dialog

**Priority:** P1  
**Source:** gitbrent/PptxGenJS#1442

### Symptom

Setting a solid background color emits `<p:bgPr>` without an
`<a:effectLst/>` child element. PowerPoint repairs the file by adding
the missing element.

### Reproduction

See #1442 for the exact reproducer.

### Root-cause hypothesis

Background emission in `src/gen-xml.ts` (or wherever `<p:bgPr>` is
generated) does not include the empty `<a:effectLst/>` that OOXML
treats as required.

### Acceptance criteria

`<p:bgPr>` always contains `<a:effectLst/>`. PowerPoint opens the file
without a repair prompt.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1442

---

## B13: Shapes without text emit `<p:sp>` missing `<p:txBody>`, triggering repair

**Priority:** P1  
**Source:** gitbrent/PptxGenJS#1441

### Symptom

A shape with no text content emits a `<p:sp>` element missing
`<p:txBody>`. PowerPoint flags this as malformed and repairs.

### Reproduction

See #1441 for the exact reproducer. Minimal: `slide.addShape(pres.shapes.RECTANGLE, { x:1, y:1, w:2, h:1 });`
with no `text` property.

### Root-cause hypothesis

`<p:sp>` emission in `src/gen-xml.ts` skips `<p:txBody>` when text is
empty. OOXML requires `<p:txBody>` to be present, even if empty.

### Acceptance criteria

Every `<p:sp>` element contains a `<p:txBody>`, regardless of whether
the shape has text content.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1441

---

## B14: Table cell margin produces `NaN` XML attributes for non-numeric values

**Priority:** P1  
**Source:** gitbrent/PptxGenJS#1440

### Symptom

When a table cell's `margin` option is not a number or array of
numbers, the emitted XML contains `NaN` attribute values, producing a
corrupt file.

### Reproduction

See #1440. Minimal: pass a string, object, or undefined to the
`margin` cell option and inspect the resulting `<a:tcPr>` margin
attributes.

### Root-cause hypothesis

Margin normalization in table-cell emission does not validate inputs
and lets `NaN` propagate into XML.

### Acceptance criteria

Invalid `margin` values either coerce to a sensible default (with a
warning) or throw a clear error. No `NaN` ever appears in the emitted
XML.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1440

---

## B15: Malformed notesMaster placeholder shapes get stripped by PowerPoint repair

**Priority:** P1  
**Source:** gitbrent/PptxGenJS#1443, #1449 (Defect 5)

### Symptom

`notesMaster1.xml` contains 6 placeholder shapes (Header, Date, Slide
Image, Notes Body, Footer, Slide Number) that PowerPoint considers
malformed; repair strips all 6, leaving an empty shape tree.
Additionally, the notesMaster's `_rels` file references `theme1.xml`
(presentation theme) — PowerPoint expects a separate `theme2.xml`.

### Reproduction

Build any presentation; inspect `ppt/notesMasters/notesMaster1.xml`
and its rels. Open in PowerPoint and confirm the repair dialog.

### Root-cause hypothesis

The notesMaster shape-tree generator emits placeholder shapes with the
wrong attributes / type values for OOXML. Theme reference uses the
wrong slot.

### Acceptance criteria

Either: (a) generate a conformant notesMaster with its own `theme2.xml`
and valid placeholder shapes that survive PowerPoint repair, or (b)
add an option to disable notes generation entirely (default off would
be a behavior change — discuss with maintainer in PR).

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1443, https://github.com/gitbrent/PptxGenJS/issues/1449 (Defect 5)

---

## B16: Unused Default extension types in `[Content_Types].xml`

**Priority:** P3 (file bloat, non-conformant — does not trigger repair on its own)  
**Source:** gitbrent/PptxGenJS#1449 (Defect 4)

### Symptom

`[Content_Types].xml` registers Default extension types for formats
not used in the presentation (e.g. `jpeg`, `gif`, `m4v`, `xlsx`,
`vml`) even when only PNG images are embedded.

### Reproduction

```javascript
var pres = new pptxgen();
pres.addSlide().addImage({ data: PNG_BASE64, x:1, y:1, w:2, h:2 });
pres.writeFile({ fileName: 'b16-repro.pptx' });
// Inspect [Content_Types].xml — should only have png, xml, rels
// extensions; bug: lots of unused entries.
```

### Root-cause hypothesis

Content-Types emitter writes a hard-coded list of Default extensions
rather than a dynamic list based on actual media in the package.

### Acceptance criteria

Only Default extensions for media types actually present in the
package are emitted.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1449 (Defect 4)

---

## B17: Empty `charts/` and `embeddings/` directories created always

**Priority:** P3  
**Source:** gitbrent/PptxGenJS#1449 (Defect 6)

### Symptom

Every generated `.pptx` contains empty `ppt/charts/_rels/` and
`ppt/embeddings/` directories, even when the presentation has no
charts or embedded objects.

### Reproduction

Generate any minimal deck. Unzip. Check for empty directories.

### Root-cause hypothesis

Scaffolding directories are unconditionally created in the zip-write
phase.

### Acceptance criteria

Empty `charts/` and `embeddings/` directories are not present in the
output unless the presentation actually contains charts / embedded
objects.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1449 (Defect 6)

---

## B18: Placeholder objects created as TEXT instead of PLACEHOLDER

**Priority:** P2  
**Source:** gitbrent/PptxGenJS#1453

### Symptom

When `addText` is given a placeholder name, the resulting shape is
emitted as a `<p:sp>` text body rather than as a `<p:ph>` placeholder
referencing the master, so PowerPoint shows "Click to add text" in
the placeholder slot.

### Reproduction

See #1453 for the full reproducer.

### Root-cause hypothesis

Placeholder dispatch in `src/gen-xml.ts` does not distinguish
placeholder-targeting `addText` calls from regular text-shape calls.

### Acceptance criteria

A placeholder-targeting `addText` produces a `<p:ph type="..." idx="...">`
on the shape rather than free text, and the master placeholder is no
longer visible at "Click to add text" state.

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1453

---

## B19: `bullet: { type: "bullet" }` produces no bullet

**Priority:** P2  
**Source:** gitbrent/PptxGenJS#1432

### Symptom

Passing `{ type: "bullet" }` to the `bullet` option silently disables
bullets, where the user expected the default bullet rendering.

### Reproduction

```javascript
slide.addText('hello', { x:1, y:1, w:4, h:0.5, bullet: { type: 'bullet' } });
// Expected: bulleted line. Actual: no bullet.
```

### Root-cause hypothesis

Bullet-config parsing checks for specific `type` values but does not
treat `"bullet"` as the default. Either accept it as the default
(`<a:buChar>` with default char) or document and error.

### Acceptance criteria

`bullet: { type: "bullet" }` produces a default bullet (matching
`bullet: true`).

### References

- Upstream: https://github.com/gitbrent/PptxGenJS/issues/1432

---

## Out of scope for this loop

Listed for future work, not for autofix to pick up:

- **#1396 Charts not showing in Apple Numbers** — has a dedicated
  `issue-1396` upstream branch and is Keynote/Numbers-rendering specific;
  larger scope than a focused autofix.
- **Slide transitions, multi-column text, shape grouping** — feature
  gaps not currently tracked as defects.
- **#1349 RTL/LTR mixed text rendering, #1262 hyperlink WPS, #1306
  graphs in Mac Keynote, etc.** — third-party renderer quirks, not
  bugs in PptxGenJS XML emission.
- Any issue older than 2024 with no recent activity — likely
  resolved in v3.13/v4 era; let triage flag separately if still
  reproducible.
