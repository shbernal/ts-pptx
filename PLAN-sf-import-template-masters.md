# Plan: `dn-import-template-masters` — `Presentation.fromTemplate()`

## Context

Backlog item `dn-import-template-masters` (status `deferred`, p3, `docs/backlog.yml:3102`)
asks for a way to **reuse a real PowerPoint template's slide masters, layouts, and
theme** when authoring a new deck, instead of rebuilding them in code with
`defineSlideMaster()`. Its `next_action` was to "scope a read-and-adopt path over the
OPC model once Phase 2 stabilizes" — Phase 2 (`dn-append-onto-existing-deck`) is now
implemented, so the precondition is met.

Exploration showed the read-centric machinery is **already in place**; this is an
ergonomics + edge-case feature, not new core infrastructure:

- `Presentation.appendSlides(source, { layout })` (`src/read/api/presentation.ts:619`)
  authors generator slides and splices them into a loaded deck bound to an existing
  layout, keeping masters/layouts/theme **byte-identical**.
- `Presentation.layouts()` (`presentation.ts:572`) enumerates the deck's layout
  gallery as `LayoutHandle[]` — the names a caller binds to.
- `removeSlide(index)` (`presentation.ts:389`) prunes only slide-private parts and
  never shared chrome; its own docstring states *"removing every slide leaves a valid
  master/layout-only package (a template shell)."*
- `OpcPackage.save()` re-emits byte-stably; `ContentTypes.ensureRegistered()`
  (`src/read/opc/content-types.ts:50`) can flip a part's content-type override.

So "use this corporate template" reduces to: **load → strip slides to a shell →
`appendSlides` → `save`**. The gap is (1) an ergonomic entry point so callers don't
have to know the loop-`removeSlide` recipe, and (2) accepting `.potx` template
packages and saving them as an editable `.pptx`.

**Decisions (confirmed with user):** primary scenario = *author a fresh deck on a
template*; accept `.potx` and normalize the saved output to an editable `.pptx`;
API shape = a static `Presentation.fromTemplate(input)` factory (mirrors
`Presentation.load()`).

## Intended outcome

```ts
import { Presentation } from 'pptxgenjs/read'
import PptxGenJS from 'pptxgenjs'

const deck = await Presentation.fromTemplate(templateBytes)  // .pptx or .potx
deck.layouts().map(l => l.name)                              // discover template layouts

const pptx = new PptxGenJS()
pptx.defineLayout?.(/* match template size */); pptx.layout = 'LAYOUT_WIDE'
const slide = pptx.addSlide(); slide.addText('Hello', { x: 1, y: 1, w: 6, h: 1 })

await deck.appendSlides(pptx, { layout: 'Title and Content' })
const out = await deck.save()  // editable .pptx using the template's masters/layouts/theme
```

## Implementation

### 1. Fixtures (do this first — `.potx` is a hard precondition)

Per `CLAUDE.md` fixture-gated rules, author the test oracles before writing code
that targets them.

- **`test/read/fixtures/template.potx` — DONE (authored 2026-06-24).**
  A genuine PowerPoint-authored `.potx` template package, authored on Windows desktop
  PowerPoint COM via the `powerpoint-fixture-authoring` skill (`SaveAs(..., 26)` =
  `ppSaveAsOpenXMLTemplate`; note `27` is the macro-enabled `.potm` variant and writes
  the wrong content-type). It:
  - declares its main part `/ppt/presentation.xml` with content-type
    `application/vnd.openxmlformats-officedocument.presentationml.template.main+xml`
    (this is what makes it a real `.potx`, and is exactly what the content-type flip
    must rewrite);
  - carries one slide master, a theme, and the 11 standard **named** layouts
    (incl. "Title and Content" and "Blank");
  - carries **zero sample slides**, so the shell-strip step is a proven no-op on real
    template input.
  Slide size is 16:9 widescreen (960×540 pt). Opens clean via COM (no repair prompt);
  SHA-256 `dd96acd1f395cb961f2222047e03263df4cbe1bdacce3735bcd934783fad0556`. Provenance,
  hash, and purpose are recorded in `test/read/fixtures/README.md`. This fixture gates
  the `.potx` → editable `.pptx` content-type flip assertion (test #4) and the "opens
  as an editable deck" check; both can now be implemented. The rest of the feature is
  not blocked on it.

- **`.pptx` template path — no new fixture needed.** Verified existing fixtures cover
  it: `test/read/fixtures/multi-theme.pptx` (2 sample slides + 18 named layouts incl.
  "Title and Content"/"Blank") and `theme-colors.pptx` (1 sample slide) exercise
  shell-stripping, chrome byte-stability, and author-on-template end-to-end.

### 2. `Presentation.fromTemplate(input, options?)` — `src/read/api/presentation.ts`

Add a static factory near `load()` (`presentation.ts:294`) and `fromPackage()` (`:299`).

```
static async fromTemplate(input: OpcInput, options?: FromTemplateOptions): Promise<Presentation>
```

Behavior:
1. `const pres = new Presentation(await OpcPackage.load(input))`.
2. **Normalize a `.potx` main part to an editable `.pptx`** (default on): if the
   presentation part's content-type is the template main type
   `application/vnd.openxmlformats-officedocument.presentationml.template.main+xml`,
   call `pres.opc.contentTypes.ensureRegistered(pres.presentationPart.partName,
   PRESENTATION_MAIN_CT)` to flip it to
   `…presentationml.presentation.main+xml` (the constant already exists in the
   `PRESERVED_*` set at `presentation.ts:51` — promote it to a named const). The
   officeDocument relationship that `presentationPart` resolves through is
   content-type-independent, so a `.potx` already loads; only the override needs
   flipping. Skippable via `options.keepTemplateContentType`.
3. **Strip sample slides to a shell:** `while (this.slides.length > 0)
   this.removeSlide(0)`. Reuses `removeSlide`'s pruning + byte-stability guarantee;
   a `.potx` that already has zero slides is a no-op. (Most real templates carry
   sample slides we don't want.)
4. Return `pres`.

`FromTemplateOptions` (new exported interface): `{ keepTemplateContentType?: boolean }`.
Keep it minimal; do not add a `keepSlides` toggle — that case is just `load()`.

JSDoc must state: masters/layouts/theme stay byte-identical; the caller authors
slides with a generator sized to match the template (`appendSlides` enforces equal
slide size) and binds them to a layout name from `layouts()`.

### 3. Re-export — `src/read.ts`

Export the new `FromTemplateOptions` type alongside the existing
`Presentation` / `LayoutHandle` / `AppendSlidesOptions` re-exports (`src/read.ts:26-28`
area). `Presentation` itself is already exported.

### 4. No generator-side changes

The generator (`defineSlideMaster`, `makeXmlMaster/Layout/Theme`) is **not** touched.
The read-centric path keeps the template's authored parts verbatim, which is higher
fidelity than round-tripping master/layout/theme XML back through the generator's
lossy `SlideLayoutInternal` model. This matches the architecture chosen for
`dn-append-onto-existing-deck`.

## Critical files

- `src/read/api/presentation.ts` — new `fromTemplate()` factory + `FromTemplateOptions`;
  promote the `presentation.main+xml` literal to a named const, add a
  `template.main+xml` const for the check.
- `src/read.ts` — re-export `FromTemplateOptions`.
- `src/read/opc/content-types.ts` — **reused as-is** (`ensureRegistered`).
- `test/read/fixtures/template.potx` — **authored 2026-06-24** (see step 1); gates the
  `.potx` assertions only.
- `test/read/template-masters.test.js` — new test (below). Reuses existing
  `multi-theme.pptx` / `theme-colors.pptx` for the `.pptx` path.
- `CHANGELOG.md` — record the new public read API.
- `docs/backlog.yml` — flip `dn-import-template-masters` to `implemented` (below).

## Testing & verification

New `test/read/template-masters.test.js`, following the harness in
`test/read/append-onto-existing.test.js` (`partBodies()` byte-diff helper,
`validateBuf`, fixture loader). Use an existing multi-master fixture such as
`test/read/fixtures/multi-theme.pptx` or `theme-colors.pptx`.

1. **Shell strip:** `fromTemplate(bytes)` → `slides` is empty; `layouts().length`
   unchanged vs a plain `load()` of the same bytes (chrome preserved).
2. **Chrome byte-stability:** diff `partBodies()` before/after — every
   `slideMasters/*`, `slideLayouts/*`, `theme/*` part is byte-identical; only
   `presentation.xml`, its `.rels`, and `[Content_Types].xml` change.
3. **End-to-end author-on-template:** `fromTemplate` → `appendSlides(pptx, { layout:
   <a real template layout name> })` → `save()` → reopen via `Presentation.load`;
   assert slide count, layout binding points at the existing layout (no new chrome),
   and (when the validator is installed) the output is schema-valid.
4. **`.potx` content-type flip — gated on `template.potx` (step 1).** Load
   `template.potx` via `fromTemplate`; assert the presentation part now resolves to
   `…presentation.main+xml`, that `keepTemplateContentType: true` leaves it as
   `…template.main+xml`, that masters/layouts/theme are byte-identical, and (validator
   installed) the saved output is schema-valid. This is the test the `.potx` fixture
   exists for; it cannot be written against existing fixtures.

**Sequencing:** `template.potx` is now in place (authored 2026-06-24), so all four
tests can land together — tests #1–#3 on the existing `.pptx` fixtures and test #4 on
`template.potx`. The remaining open item is the "opens as editable deck in PowerPoint"
desktop verification of the *saved* (content-type-flipped) output, which is performed
after the code lands by reopening the `fromTemplate → save` result in PowerPoint.

Commands:
- `pnpm run build && pnpm run typecheck`
- `pnpm run test:unit` and the new read test (`pnpm run test:schema` for validator).
- `pnpm run backlog:validate` after editing `docs/backlog.yml`.

## Backlog update (on implementation)

Per `docs/backlog-workflow.md`: set `dn-import-template-masters` `status:
implemented`, `last_reviewed` to today, `evidence.kinds` to include
`local-source-inspection` + `regression-test` (+ `validator-result` if asserted),
`evidence.local_files` to `src/read/api/presentation.ts`, `src/read.ts`,
`test/read/template-masters.test.js`, `next_action: none`, and write the
"where it landed" summary into `current_project_notes`. Add a `CHANGELOG.md` entry.
If the `.potx` open-in-PowerPoint assertion stays open, leave a one-line note (or a
small fixture-gated companion item) for the genuine `.potx` fixture.
