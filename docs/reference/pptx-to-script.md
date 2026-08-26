---
doc-schema-version: 1
title: "PPTX To Script"
summary: "Turn an existing .pptx into runnable TypeScript that rebuilds it through the public write API, with a declared, machine-checked list of what it drops."
read_when:
  - Converting an existing deck into editable ts-pptx source
  - Choosing between the template-anchored and standalone output tiers
  - Interpreting a conversion's fidelity notes
  - Changing the deck IR, either printer, or the round-trip check
doc_type: "guide"
---

# Turning a deck back into source (`ts-pptx/script`)

The `ts-pptx/script` subpath reads an existing `.pptx` through `ts-pptx/read`
and emits **TypeScript source** that rebuilds an equivalent deck through this
library's public write API. The deck stops being an opaque binary and becomes
something you can diff, parameterize, and regenerate.

It is **lossy by construction and by agreement**. This library does not cover
every OOXML construct, and (measured, not assumed) the *read* side is the
tighter of the two constraints. So the deliverable is never "a perfect copy".
It is a faithful script plus an honest account of what it dropped, and that
account is data rather than log output: see [Fidelity notes](#fidelity-notes).

```ts
import { copyFile, readFile, writeFile } from 'node:fs/promises'
import { Presentation } from '@shbernal/ts-pptx/read'
import { readModelToIr, printScript } from '@shbernal/ts-pptx/script'

const deck = await Presentation.load(await readFile('source.pptx'))
const { code, assets, notes } = printScript(readModelToIr(deck))

await writeFile('out/deck.ts', code)
for (const [name, bytes] of assets) await writeFile(`out/assets/${name}`, bytes)

// Template-anchored output expects the source deck beside the script, unmodified.
await copyFile('source.pptx', 'out/template.pptx')

console.log(`${notes.length} declared loss(es)`)
```

Then `node out/deck.ts` writes `out/output.pptx`. `code` is a runnable ESM
TypeScript module (it uses top-level `await`), and Node runs it directly by type
stripping: no build step and no extra dependency. Every path inside the script
resolves against the script's own location rather than the working directory.

`printStandaloneScript` is the same call with no `copyFile` step.

## The two tiers

The same IR is printed by two printers. They differ in exactly one thing: where
the deck's **chrome** (its masters, layouts and theme) comes from.

| | `printScript` (template-anchored | `printStandaloneScript`) standalone |
|---|---|---|
| Chrome | the original, byte for byte | re-authored from what the read model exposes |
| Ships | the script + the source deck as a binary asset | the script alone |
| Entry | `Presentation.fromTemplate` + `appendSlides` | `new TsPptx()` + `defineSlideMaster` |
| Editable | slide bodies | everything |
| Charts, `chartEx`, table styles, `docProps` | carried across untouched | transcribed, or lost |

Pick the template-anchored tier when the output has to *look* like the source
and you can ship the source deck alongside the script. Pick standalone when the
script must stand on its own, and accept that the deck wears a different suit.

### Template-anchored output

The template **is the source deck, unmodified**. No strip step is needed:
`Presentation.fromTemplate` already removes every slide while leaving masters,
layouts, theme and document properties byte-identical, so only slide content is
ever regenerated.

```ts
const deck = await Presentation.fromTemplate(here('./template.pptx'))

function generator(): TsPptx {
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'source', width: 13.333333, height: 7.5 })
	pptx.layout = 'source'
	return pptx
}

const gen1 = generator()
const slide1 = gen1.addSlide()
slide1.addText([{ text: 'Q3 Results', options: { bold: true } }], { x: '2068830emu', /* … */ })

await deck.appendSlides(gen1, { layout: 'Titelfolie' })
await writeFile(here('./output.pptx'), await deck.save())
```

Slides are emitted in source order and every operation appends, so `p:sldIdLst`
comes out right with no position arithmetic. Contiguous slides sharing a layout
share one generator, because `appendSlides` binds one layout per call.

Binding is by layout **name** where that is unambiguous, since a name survives
being re-pointed at a different template; a deck whose layouts repeat a name
falls back to gallery position, because `appendSlides` throws on an ambiguous
name rather than choosing.

### Standalone output

```ts
const pptx = new TsPptx()
pptx.defineLayout({ name: 'source', width: 13.333333, height: 7.5 })
pptx.layout = 'source'
pptx.theme = { headFontFace: 'Calibri Light', bodyFontFace: 'Calibri', colorScheme: { /* 12 slots */ } }
pptx.author = 'Thomas Singer'
pptx.defineSlideMaster({ background: { color: 'FFFFFF' }, title: 'Titelfolie' })

const slide1 = pptx.addSlide({ masterTitle: 'Titelfolie' })
slide1.addText(/* … */)

await pptx.writeFile({ fileName: here('./output.pptx') })
```

One `defineSlideMaster` per source layout, carrying **a title and a background
and nothing else**. That thinness is a write-path constraint, not a shortcut:
`addPlaceholdersToSlideLayouts` seeds every slide with each layout placeholder
the slide did not populate, as an empty text shape. This converter authors every
source shape as concrete absolute-positioned content and binds none of them to a
placeholder, so re-declaring a layout's placeholders would add a ghost shape to
every slide for each one: measured at five to eight per slide on the `mixed`
fixture, which is a deck of ordinary complexity. The title
still earns its place: it is the key `addSlide({ masterTitle })` matches on, so
the output keeps a layout gallery a reader recognises.

## Why there are two tiers: the chrome cliff

Slide *bodies* map near one-to-one from the read model onto write-API option
objects. Slide *chrome* does not, and four constructs are the reason. Each is
free in the template-anchored tier and unreachable in the standalone one, and
**no amount of printer work moves them**: two are unreachable from *both*
directions at once:

- **`a:fmtScheme`**: the three fill, three line and three effect style lists a
  shape's `<p:style>` indexes into. Nothing on the read path exposes it and
  nothing on the write path sets it; the write path emits a hardcoded Office
  one. So a shape whose outline comes from the theme line matrix keeps its
  colour but not its width, dash or effect.
- **`p:txStyles`**: the master's per-level default size, face, colour, indent
  and bullet for each of the nine list levels. No read accessor, although
  `SlideMasterProps.textStyles` could author them if they could be seen.
- **A master's own decoration**: the shapes a *master* carries. Structural
  rather than a reading gap: `defineSlideMaster` creates a layout under the one
  shared master, so a master's shape tree has no write-side counterpart at all.
- **`p:clrMap`**: readable, with no write-API setter.

Multi-master decks add a fifth: a generated deck has a single shared master, so
a source deck with several has no structural counterpart and collapses.

A **layout's** decoration used to head that list and no longer does.
`SlideLayout.shapes` decodes it and the converter transcribes each shape into
that layout's `defineSlideMaster({ objects })` array: through the same mapper
the slides go through, so a band or a wordmark on a layout is decided by the
code that decides one on a slide. What is left is narrow and named per shape
under the `layout.` prefix: a table has no variant in that union, and a group is
flattened into its children because the union has no `group` either.

This asymmetry is why the template-anchored tier shipped first. It rides
primitives that already existed and were already tested, and it makes those five
losses disappear rather than manage them.

## The pipeline

```
read model  →  DeckIr  →  printed text
```

Never read model → text directly. `readModelToIr(presentation)` owns every
semantic decision; a printer only knows how a value is spelled. That is what
makes the mapping testable without a printer and keeps "how a number is
formatted" from changing what a deck means.

`DeckIr` is `{ slideSize, props, chrome, slides, assets, fidelity }` and is
serializable: no DOM, no read-model object references. A slide is
`{ number, source, layout, hidden, name?, background?, notesText?, transition?, calls }`,
where each call is `{ method, args }` and `args` are literal write-API option
objects. Media cannot be a literal, so it is an `AssetRef` (`{ $asset }`)
resolved against `assets`; the *printer*, not the IR, decides between a file
beside the script and an inline `data:` URI.

`DeckIr.chrome` is read by the standalone printer only: for the
template-anchored one the source deck *is* the chrome.

### Geometry is EMU-exact

Every `Coord`-typed option accepts a raw `"<n>emu"` string, so geometry is
printed verbatim and reaches `a:off`/`a:ext` unrounded. Four options are typed
in inches instead and are printed at **six decimal places**: `colW`, `rowH`,
`margin`, and `defineLayout`'s `width`/`height`.

Six is a proven minimum, not a preference. Inches round-trip EMU exactly at full
double precision; the loss appears only if the printed decimal is truncated. A
six-decimal round shifts a value by at most 0.4572 EMU, inside the half-EMU
bound that makes `Math.round` return the original; at five decimals the bound is
4.572 EMU and the round trip fails for most values. Rounding below six to
suppress cosmetic `0.5000000001` noise would be a real geometry loss.

`defineLayout` is the strictest of the four: `appendSlides` compares the two
decks' EMU sizes for *equality*, so imprecision there throws rather than drifts.

### Transitions are filtered against a closed vocabulary

A slide's show transition is transcribed in both tiers, as a property assignment
rather than a call: `slide.transition = { type: 'push', speed: 'slow',
durationMs: 1250, variant: { dir: 'd' } }`. Speed bucket, exact `p14:dur`
duration, `advClick`/`advTm` advance behaviour and the type-specific variant
attributes all carry across; each is omitted from the emitted literal when the
source left it at its OOXML default.

The one judgement is the **type**. `TransitionInfo.type` is an open string,
because the read model also decodes PowerPoint's modern effects (Morph, Vortex,
Ripple, …) and distinguishes them by *namespace*; the write API's
`TransitionType` is a closed union of the 21 base ECMA-376 names. So the
converter admits only `p`-namespaced names it can spell and files a
`slide.transition` note for the rest: the alternative is a printed script that
does not compile, on exactly the decks a converter is most likely to meet.
PowerPoint's own probed effect table (captured in
`test/read/fixtures/slide-transition.oracle.json`) lists 21 base effects and 21
modern ones, and the base 21 match the write union exactly.

Transition **sounds** map in both OOXML forms: the stop-previous `p:endSnd`, and
an embedded start sound whose WAV is resolved through the slide's own `r:embed`
and carried as an asset like any image. The second survives the standalone tier
only: see [Read the printer's notes](#read-the-printers-notes-not-the-irs).

### Theme colours are not uniformly flattened

The write path accepts 10 scheme tokens (`tx1 tx2 bg1 bg2 accent1`–`accent6`)
where `ST_SchemeColorVal` has 17. In the template-anchored tier the destination
theme is the source theme by construction, so those 10 pass straight through as
`schemeClr`: equivalent *and* still theme-responsive. The other 7 (`dk1 lt1 dk2
lt2 hlink folHlink phClr`) are silently repainted `000000` by the write path, so
the converter resolves them to hex against the theme and files a note. The
standalone tier flattens accordingly.

## Fidelity notes

Every construct that cannot survive becomes a `FidelityNote` on the IR rather
than a warning to a log. A warning is prose that nothing consumes, and a missing
one looks exactly like a clean run. Notes as data invert that:

- an **undeclared** loss fails the round-trip check, because nothing excluded it;
- a **declared** loss that actually survives is a stale note, and the check can
  say so.

```ts
interface FidelityNote {
	slideNumber: number | null // 1-based source slide, null for a deck-level loss
	shapeName: string | null // p:cNvPr/@name, so the note points at a call in the script
	construct: string // stable dotted key — 'line.width', 'text.tabStops'
	disposition: 'dropped' | 'flattened' | 'approximated'
	cause: 'unread' | 'unwritable' | 'unsupported'
	detail: string // why, in a sentence
}
```

`construct` is an identifier rather than a sentence because the round-trip check
matches it mechanically against a field path. `cause` is what makes a note
**actionable**: `unread` and `unwritable` are gaps in a specific subsystem and
could be closed, while `unsupported` is a property of OOXML or of the chosen tier
and will not yield to more converter work.

### Read the printer's notes, not the IR's

`PrintedScript.notes` is **not** `DeckIr.fidelity`. A tier both suppresses and
adds:

- *Suppressed:* all twelve document properties ride in the template, so the IR's
  `deck.docProps` note does not describe a template-anchored output. A caveat
  that does not apply teaches the reader to skim the ones that do.
- *Added:* a slide's own name (`p:cSld@name`) reads fine and would survive a byte
  copy, but has no public write-API setter, so it dies in both tiers and nowhere
  else in the library.
- *Added, template-anchored only:* a transition's **embedded start sound**
  (`slide.transitionSound`). The standalone tier writes a real package and keeps
  it; the append path this tier rides never runs the pass that registers a
  transition's audio part, so the sound is dropped: silently, and without a
  dangling reference, which is the safe half of the failure and still a loss. The
  stop-previous form (`p:endSnd`) needs no part and survives in both tiers.

The applicable set is reproduced as a comment block at the top of the emitted
script, so the artifact carries its own caveats, and it is the set a round-trip
check must exclude from its diff.

## What actually gets lost

Measured across the 46-fixture corpus by `pnpm run script:census`, which is what
keeps the numbers below honest: a closed reader gap or a new fixture moves them
without failing anything. The count is how many fixtures raise the note at least
once, not how many notes fired; the corpus is construct-targeted, so this
measures **coverage, not frequency**: it says what a converter meets, not what a
real deck is mostly made of.

Both tiers, in corpus order:

| construct | fixtures | cause | what it costs |
|---|---|---|---|
| `text.color.inherited` | 28/46 | unsupported | an uncoloured run would be painted black, so the inherited colour is resolved and baked in |
| `shape.placeholder` | 9/46 | unsupported | placeholder *identity* degrades; 6 of 16 `ST_PlaceholderType` values are expressible and `idx` has no setter |
| `shape.frameInherited` | 8/46 | unsupported | geometry inherited from a layout is reproduced exactly, then frozen: it stops tracking layout edits |
| `text.color.default` | 8/46 | unread | nothing resolves what this run inherits, so the write path paints it black: the one case where the output colour is not merely frozen but possibly *wrong* |
| `line.width` | 7/46 | unread | an outline from the theme line matrix (`p:style/a:lnRef`) keeps its colour and loses its width and dash |
| `slide.animation` | 7/46 | unread | build animation has no structural reader |
| `slide.carried` | 3/46 | unwritable | template-anchored only: the slide holds a graphic frame with no write-API emitter, so it is copied from the source rather than transcribed |
| `slide.carriedChrome` | 3/46 | unsupported | template-anchored only: the cost of that copy, one duplicate layout-gallery entry per carried slide |
| `media.audioVideo` | 2/46 | unread | only the poster frame is readable, so embedded A/V becomes a still image |
| `text.equation` | 2/46 | unread | the whole `m:` namespace is absent from the read path, so OMML math is invisible |

Plus, at 1–2 fixtures each: `chart.workbook`, `diagram.all`,
`graphicFrame.unknown`, `group.childSpace`, `group.transform`, `image.recolor`,
`shape.empty`, `connector.binding`, `fill.gradient.path`, `fill.schemeToken`,
`group.child`, `image.svg`, `line.arrowSize`, `shape.custGeom.guides`,
`slide.layout`, `table.cell.fill.picture.geometry`, `table.rowAuto`,
`text.bullet.schemeToken`, `text.field`, `text.paraSpaceZero`.

**A slide holding a graphic frame the write API cannot author is copied, not
transcribed.** That is the rule, and `slide.carried` is how the converter says it
applied. Three frame payloads qualify today (an extended chart, a SmartArt
diagram, and a frame the reader does not decode at all), but the list is not the
point: any frame that produces no call has the same consequence, because a script
that silently omits it would claim to describe a slide it does not. The per-shape
note beside `slide.carried` says which construct forced the copy.

The standalone tier has no source package to copy from, so it transcribes such a
slide and genuinely loses the frame. That is why the per-shape notes stay: they
are what that tier reports, and `slide.carried` is suppressed there.

**`diagram.all` and `graphicFrame.unknown` are different losses, and used to be
one note.** A SmartArt frame has a full reader, and its text can now be edited in
place through it (`DiagramPoint.text`, see
[the read reference](./pptx-read.md#diagram-smartart)). What a converted script
loses is the *authoring* leg: no write API builds a diagram from scratch, which
makes it `unwritable` alongside `chartEx.all` rather than `unread`.
`graphicFrame.unknown` keeps its original meaning and its original cause, and
names only the frames that really are undecoded: the corpus raises it on
`model3d.pptx` alone.

**An inherited bullet is not a loss.** It used to be the largest one here, at
34/44: the top of this table, and 305 of the standalone tier's notes on its own.
A paragraph with no bullet child of its own inherits whatever the layout's or
master's list style says, and the write API had no way to state that: omitting
`bullet` emitted an explicit `<a:buNone/>` plus `marL="0" indent="0"`, which
*overrides* the list style rather than deferring to it. That is a different fact
even where the inherited style has no bullet, because a later edit to the master
then stops arriving, and a visible change where it did have one, along with an
inherited hanging indent flattened to zero in the same stroke. `bullet: 'inherit'`
is the spelling for the third state, emitting neither a bullet child nor
`a:buNone` nor the margins, so `Paragraph.bulletDetail` returning `null` (no
bullet child) maps onto it and `{ kind: 'none' }` (a stated `a:buNone`) still maps
onto `false`. The distinction always survived the read leg; it died on the write
leg, which made this a missing option rather than an unreadable construct. Neither
state notes now, and `layout.text.bullet.inherited` closes with it.

**A paragraph's own margins are not a loss either, and they were the other half
of the same element.** `text.indent` read 5/44 and was the largest note left on
`a:pPr` once the bullet one closed. `a:pPr/@marL` and `@indent` had no write
option at all, so whichever `bullet` state a paragraph mapped onto decided them:
a drawn bullet re-hung the first line by the writer's own 27pt default no matter
what the source said, and `bullet: false` flattened both to zero. `paraMarginLeft`
and `paraIndent` state them now, in points, with `'inherit'` for the paragraph
that states neither: which a bulleted paragraph needs, since omitting the option
is what writes the default. That is the third state again, one attribute over
from `bullet: 'inherit'`, and the reader did not move here either:
`Paragraph.marginLeftPt` and `Paragraph.indentPt` already separated a stated
margin from an absent one. The note was empty in the round trip's exclusion table
(neither IR carried the field, so the check compared two models both missing it)
which is why closing it is what makes the margins *verified* rather than merely
declared.

**A styled cell's own fill is not a loss.** It used to be: the note read at 7
fixtures, because `resolvedFill` answers "what colour is this cell" by folding
the cell's own fill together with the colour it merely inherits from the style's
header and banding rules, and writing that back would turn every banded cell into
an explicitly filled one. `TableCell.hasOwnFill` separates the two: a cell whose
`a:tcPr` carries an `EG_FillProperties` child emits that fill, and a cell with
none is left to the style GUID, which reproduces the banding exactly rather than
approximately. Neither case records a note, and the bare `table.cell.fill` key is
retired rather than merely unfired: its `.gradient`, `.gradient.path`,
`.picture` and `.picture.geometry` children are separate constructs and stay.

**Picture fills carry; their geometry does not.** An image-filled *surface* (a
shape's `p:spPr/a:blipFill` or a cell's `a:tcPr/a:blipFill`) is re-embedded
through the same asset resolver an `addImage` uses, so the bytes and the blip's
`a:alphaModFix` opacity survive. What does not is everything around them: the
write path emits every picture fill as a plain stretched blip (`dpi="0"
rotWithShape="1"`, `<a:srcRect/><a:stretch><a:fillRect/></a:stretch>`), so a
tiled fill comes back stretched and a cropped or inset one comes back whole.
That is `fill.picture.geometry` / `table.cell.fill.picture.geometry`,
`approximated` and `unwritable`, and it is recorded only when the source
actually uses one of them, one fixture does, the PowerPoint-authored tiled
cell in `table-cell-image-fill.pptx`.

**An outline's `@cap` carries; its `@algn` does not.** `a:ln/@cap` is mapped
onto `ShapeLineProps.cap` (`flat`/`sq`/`rnd` → `flat`/`square`/`round`) and
records no note, because both legs exist: the write API authors the attribute and
`AutoShape.lineCap` reads it back. It is not cosmetic (on a thick dashed rule the
cap extends every dash by the stroke width and decides whether each draws as a
rectangle or a lozenge), so before the mapping existed a deck this library wrote
could not survive its own converter, and nothing said so. `@algn` is the case
where only one leg exists: readable through `AutoShape.lineAlign`, with no write
option for it, so `line.align` is `dropped`/`unwritable`. It is recorded only for
`algn="in"`, the inset stroke that sits half its width further in; `ctr` is what
an omitted `@algn` already renders as, so noting it would fire on most
PowerPoint-authored shapes while describing no loss. No corpus fixture states
`in`, so the note reads 0/46.

**A baked autofit carries its scale, and a bare one is a different state.** A
`normAutofit` frame maps onto `fit`, but not onto a single spelling: one that
bakes `a:normAutofit/@fontScale` or `@lnSpcReduction` emits the object form
`fit: { type: 'shrink', fontScale, lnSpcReduction }`, and one with neither
attribute emits `fit: 'shrink'`, which is what writes a bare `<a:normAutofit/>`.
Collapsing the two would not be a rounding, ECMA-376 §21.1.2.1.3 defaults each
attribute to 100%/0% only when it is *omitted*, and PowerPoint recomputes an
unbaked scale on edit while drawing a baked one exactly as written, so a deck
baked at `fontScale="40000"` would come back painting its text two and a half
times too large until someone clicked into the frame. Neither case notes.
`text.autofit.fontScale` / `text.autofit.lnSpcReduction` are the one arm that
does: the write path rejects a percentage outside 0–100 and drops the attribute
with a warning, so a malformed source falls back to bare `'shrink'` with the loss
declared instead of passing through a number that would vanish silently. No
corpus fixture is malformed, so both read 0/46.

**The explicit off for a text decoration is a state, not silence.** `u="none"`,
`strike="noStrike"` and `cap="none"` carry into the IR as
`underline: { style: 'none' }`, `strike: 'noStrike'` and `caps: 'none'`; only an
*absent* attribute maps to an absent option. Each is a member of its own
enumeration (ECMA-376 §20.1.10.81, §20.1.10.78, `ST_TextCapsType`) and would be
redundant with omission if omission were the only way to be off. It is not,
because run properties resolve down the `a:lstStyle` → placeholder → layout →
master chain: a run that would take `u="sng"` from its list style and states
`u="none"` is not underlined, and the same run with the attribute dropped is. The
loss was invisible on a deck with no inherited decoration and a wrong answer on
one that has any, and undeclared either way, since `canonicalDeckIr` did not
carry the field, so `diffDeckIr` compared two models that were both missing it.
Neither state notes. Two PowerPoint-authored fixtures state these tokens:
`mixed.pptx` and `table.pptx` carry 132 runs stating `u="none"` and
`strike="noStrike"`, and 100 of those also state `cap="none"`.

`fill.picture` / `table.cell.fill.picture` are what remain for a fill that
cannot carry its bytes at all, and neither fires on the corpus: a blip embedding
no part (an external or linked image), a part missing from the package, or an
SVG, which `addImage` accepts but a *fill* does not, so emitting one would
produce a script that runs, warns, and paints nothing. Those surfaces come out
unfilled, as they did before, with the note saying which case it was.

**Standalone only**: the chrome cliff, quantified. Five notes fire on *every*
fixture, which is the honest headline of that tier:

| construct | fixtures | what it costs |
|---|---|---|
| `theme.fmtScheme` | 46/46 | the output carries Office's format scheme |
| `master.txStyles` | 46/46 | placeholder text falls back to built-in defaults |
| `master.placeholders` | 46/46 | layout placeholder definitions are not reproduced |
| `deck.docProps` | 46/46 | 5 of 12 document properties have setters |
| `master.default` | 46/46 | every presentation carries an unremovable blank `DEFAULT` layout |
| `master.background` | 45/46 | a `p:bgRef` theme reference is baked to the colour it resolves to |
| `master.decoration` | 6/46 | the shapes a *master* carries: `defineSlideMaster` creates a layout, so there is nowhere to put them |
| `master.name` | 5/46 | a layout name containing a tab or line break collapses |
| `master.colorMap` | 4/46 | `p:clrMap` has no setter |
| `master.multiple` | 1/46 | multi-master decks collapse to one |
| `master.nameCollision` | 1/46 | layout titles are deduplicated, since a title doubles as a lookup key |

A **`layout.` prefix** marks the rest: a loss in re-authoring a *layout's* own
decoration, which the standalone tier rebuilds into that layout's
`defineSlideMaster({ objects })`. The vocabulary after the prefix is the slide
one, because the shape mapper is shared: `layout.line.width` is `line.width`
seen from the chrome. The prefix is not cosmetic: without it a themed outline on
a layout would be reported by the template-anchored tier, which rebuilds no
layout, and the round trip would let one excuse the same difference on a *slide*.

| construct | fixtures | what it costs |
|---|---|---|
| `layout.text.color.inherited` | 4/46 | an inherited run colour on a decorative text box, resolved and baked in |
| `layout.group` | 2/46 | a group on a layout becomes loose objects: they land unmoved, but stop being one selectable object |
| `layout.fill.schemeToken` | 1/46 | a token outside the ten the write path maps is baked to hex |
| `layout.shape.custGeom.guides` | 1/46 | a freeform's guides and adjust handles, as on a slide |
| `layout.decoration` | 0/46 | a table on a layout: no `SlideMasterObject` variant at all |

The two remaining rolled-up chrome notes are `master.decoration` and
`master.placeholders`, one each, naming the counts. A twelve-layout deck
emitting one note per layout would put twelve near-identical paragraphs at the
top of the script and bury the per-shape notes underneath that a reader can act
on. Per deck the tier adds 4 to 13 notes, not fifty (across the corpus: 705
notes against the template-anchored tier's 411).

### The read path is the binding constraint

Worth stating plainly, because it is the opposite of what it looks like from the
write side. Constructs are lost purely because **nothing reads them**, while the
write API can already express them today: tab stops, preset text warp, and
custGeom guides, adjust handles and connection sites. `pnpm run read:census`
measures that surface directly.

The bullet half of that list is now closed. `Paragraph.bulletDetail` reads
`a:buAutoNum/@startAt` and a bullet's own `a:buFont` / `a:buSzPct` / `a:buClr`,
so `text.bullet.numberStartAt` and `text.bullet.style` are gone rather than
merely unmapped: the converter emits `numberStartAt`, `fontFace`, `size` and
`color` instead of noting their absence. What remains is `text.bullet.sizePt`,
an absolute `a:buSzPts` the write API has no unit for, and
`text.bullet.picture`, where the bytes of an `a:buBlip` are readable but the
paragraph mapper carries no asset resolver to re-embed them with. Reading the
colour also opened one write-side gap of its own: `text.bullet.schemeToken`
(1/46), an `a:buClr/a:schemeClr` outside the ten tokens the write path maps,
which is baked to a literal hex and stops tracking the theme.

`text.bullet.inherited` is the counter-example to this section's thesis, and it
was the biggest note on the corpus, so it is worth being precise about why. It
was filed `unread`, and that was true of the wrong thing: what nothing reads is
the inherited *value*, since `a:lvl1pPr` list styles are still undecoded. But
reproducing an inherited bullet never required reading it: it required not
overwriting it, and `Paragraph.bulletDetail` already separated "no bullet child"
(`null`) from "a stated `a:buNone`" (`{ kind: 'none' }`). The whole loss was a
missing *write* spelling, and `bullet: 'inherit'` closed it without the reader
moving at all. A note's `cause` records the gap its author could see; it is a
hypothesis about where the fix lives, not a finding.

`text.indent` was filed the same way and closed the same way, one attribute over:
which is the part worth generalizing. It read `unwritable` rather than
`unread`, so its hypothesis was right about *which* leg, but the loss it
described stopped at "hanging indents flatten to the level default" when the
element also carried a margin that no bullet state could leave alone. Both notes
were on the same `a:pPr`, both were closed by naming a third state on the write
side, and neither needed a line of reader work. Where one note on an element
turns out to be a missing write spelling, its neighbours are worth re-reading
before they are believed.

Layout decoration is the largest case of the same pattern, and it closed
completely. It was `unread`, on the strength of the read model documenting a
template's non-placeholder content as outside its scope. `SlideLayout.shapes`
decodes it now, and the guess about the *write* side turned out to be wrong: the
`objects` union covers a plain rect, line, image, chart and text box, but it
also covers any preset via `{ shape: { type } }` (which includes `custGeom`)
and its `ShapeProps` carries fill, line, shadow, rotation and adjust handles. So
what looked like a write-side ceiling was mostly reachable, and re-tagging the
existing slide mapper's output was the whole of the work. What genuinely does
not fit is a table (no variant) and a group (no variant, so it is flattened into
children that land unmoved). A connector has no variant either and is re-authored
as a `line` preset, which paints the identical stroke and, unlike the slide-side
`addConnector`, keeps its rotation.

A **master's** decoration stayed lost, and moved from `unread` to `unwritable`
rather than closing: `defineSlideMaster` creates a layout, so a master's shape
tree has no counterpart to receive them at all. Closing a reader gap does not
always close the loss; sometimes it only tells you which half was actually
holding it, and sometimes, as here, the half everyone assumed was holding it
was not.

## Verifying a conversion

```sh
pnpm run script:roundtrip                          # template-anchored, whole corpus
pnpm run script:roundtrip -- --tier a              # standalone
pnpm run script:roundtrip -- --fixture mixed.pptx --verbose
pnpm run script:roundtrip -- --dir ~/decks         # your own decks, any path
pnpm run script:roundtrip -- --json

pnpm run script:census                             # counts, not differences
pnpm run script:census -- --names 3 --dir ~/decks
```

The harness runs `source → IR₁ → script → execute it → output → IR₂`, then
diffs the two IRs using the printer's fidelity notes as the exclusion list. A
difference no note predicted is a defect.

The comparison is a **projection diff, not byte identity**. The output package
can never be byte-identical (regenerated shape ids, fresh rel ids) so
comparing packages would report a total mismatch for every deck and measure
nothing. `canonicalDeckIr` removes what is noise rather than loss: a value
spelled out that means what its absence means (`bold: false`, `wrap: true`, the
default `a:bodyPr` insets), and asset identity by content digest instead of
generated filename. Every rule cites the OOXML default that makes it an
equivalence, because a rule that merely shortens the report hides a defect
permanently.

`--tier a` deliberately does **not** lay a template down, so a standalone script
that still needed one fails here rather than passing on the very file it is
meant to replace.

`script:census` answers the question the round trip cannot: *how much* is lost
and by which construct. To the harness a note that excuses a difference and a
note that never fires look identical, so the counts above can drift silently
while every gate stays green. The census prints both tiers without running the
scripts, so it is fast, and `--dir` accepts a corpus of real decks to trade the
coverage reading for a frequency-weighted one.

### What a clean run does not prove

Stated here because "0 undeclared differences across 44 decks" reads like proof
of correctness and is not:

- **It detects asymmetry only.** Both IRs come from the same reader through the
  same mapper, so a construct the read path cannot see is missing from both and
  compares equal. A converter that simply never emits `flipH` yields an output
  that also lacks it. `pnpm run read:census` is what measures that surface.
- **The converter need not be injective.** Two source constructs mapping onto one
  call compare equal.
- Mutation testing puts a number on that. Of twelve deliberately planted
  converter defects, the template-anchored round trip catches six and the
  standalone one seven; every survivor is a *symmetric* defect, and they are
  covered instead by direct IR expectations written from `src/types/*.ts` rather
  than from the converter.

Pair it with `pnpm run read:census` and the IR unit tests. All three are in
`pnpm run verify`.

## Options

Both printers share `assets` (`'file'`, the default, written beside the script;
or `'inline'`, one self-contained file at roughly 4/3 the byte size),
`assetDir`, `outputPath`, and `packageName` (defaults to this package's own
published name; override it to point a generated script at a local build or a
fork). `printScript` adds `templatePath`, which is where the emitted script
expects the source deck.

Every path in an emitted script resolves against the script's own location
rather than the working directory.

## Deliberately not built

- **Carrying an embedded transition sound onto a template.** The gap is in the
  append path rather than in this converter: `extractSlides` does not surface a
  transition's audio part, so `appendSlides` has nothing to reserve or wire, and
  the emitter finds no relationship id and writes no `p:sndAc`. Closing it means
  an `ExtractedSlide` descriptor and a rel-wiring branch alongside the ones for
  charts and A/V: at which point `slide.transitionSound` stops being a
  template-anchored note.
- **`p15`/`p159` transitions cannot be exercised end to end.** They are dropped
  correctly, by the same namespace check as the 19 `p14` effects, but neither
  prefix is in the read DOM's registry, so no test can author one to prove it.
- **What the reader still does not decode** remains the binding constraint: the
  converter cannot print what it cannot see:
  `lnRef` width/dash, `a:fillToRect`, `effectRef`, custGeom
  `gdLst`/`ahLst`/`cxnLst`, `a:tabLst`, `a:prstTxWarp`, plus `p:txStyles`. Every
  one of them raises the ceiling both tiers build against. The bullet entries are
  gone from this list because `bulletDetail` closed them, and master/layout
  decoration because `SlideMaster.shapes` / `SlideLayout.shapes` did: a layout's
  is now re-authored, and a master's is blocked on the write side instead.
- **A table on a layout, and `defineSlideMaster({ objects })` in general.** The
  union has no `table` variant and no `group` variant. A group is flattened into
  children that land unmoved, so it costs an editing affordance rather than a
  pixel; a table is dropped outright. Both are write-side gaps now, not reading
  ones. Neither fires on the corpus, and the write API cannot author either onto
  a layout, so the table arm is proved against a `p:graphicFrame` relocated into
  a layout part.
- **A frequency-weighted corpus.** `test/read/fixtures/` is construct-targeted
  (one feature per deck), so every count on this page is a coverage argument and
  not a frequency one. Point `--dir` at real decks to get the other kind.

## See also

- [PPTX Read / Round-Trip](pptx-read.md): the read model this builds on.
- [Architecture](../architecture.md): where `src/script/` sits and why.
- [Testing Guide](../testing.md): the verification commands.
