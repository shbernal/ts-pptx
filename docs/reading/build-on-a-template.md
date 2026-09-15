---
doc-schema-version: 1
title: "Build on a template"
summary: "Open a PowerPoint template with Presentation.fromTemplate, list its layouts, and add slides generated with TsPptx bound to a named layout, keeping the template's masters, layouts and theme byte for byte. Works for .potx and .pptx files and for decks that already have slides."
read_when:
  - Generating slides on a corporate template instead of rebuilding its masters in code
  - Turning a .potx template into an editable .pptx
  - Binding generated slides to a layout by name
  - Adding generated slides to a deck that already has slides
doc_type: "guide"
---

# Build on a template

`Presentation.fromTemplate(input)` opens a template with its sample slides removed. `appendSlides` then adds slides you generate with `TsPptx`, each bound to one of the template's layouts:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { TsPptx } from 'pptx-ts'
import { Presentation } from 'pptx-ts/read'

const deck = await Presentation.fromTemplate(await readFile('brand.potx'))
console.log(deck.layouts().map((layout) => layout.name))

const pptx = new TsPptx()
pptx.layout = 'LAYOUT_WIDE'
pptx.addSlide().addText('Hello', { x: 1, y: 1, w: 6, h: 1 })

await deck.appendSlides(pptx, { layout: 'Title and Content' })
await writeFile('brand-deck.pptx', await deck.save())
```

The saved deck keeps the template's masters, layouts and theme as the file stores them. The save rewrites `presentation.xml`, its relationships and `[Content_Types].xml`, and adds the parts of the new slides. When the new slides have speaker notes and the template has no notes master, it also adds one.

## Open a template

- `fromTemplate` accepts the same input as [`Presentation.load`](read-and-edit.md#load-and-save), from a `.pptx` or a `.potx` file.
- It removes every slide, as `removeSlide` does. The parts only those slides used go with them, such as their notes, media and charts. Masters, layouts and theme stay byte-identical.
- On a template with no slides, the removal step does nothing.
- A `.potx` declares its main part with the template content type. `fromTemplate` changes it to the presentation content type, so the saved file opens as an editable deck and not as a new copy of a template.
- `{ keepTemplateContentType: true }` keeps the template content type. A `.pptx` has the presentation type already, so the option changes nothing there.

## Find a layout

`deck.layouts()` returns one [`LayoutHandle`](../reference/api/read/interfaces/LayoutHandle.md) per layout, in master order, then in layout order within each master. It reads the deck and changes nothing.

| Field | Value |
| --- | --- |
| `name` | the layout's name (`p:cSld/@name`), or `''` when it has none |
| `partName` | the layout's part, such as `/ppt/slideLayouts/slideLayout2.xml` |
| `masterPartName` | the part of the master the layout belongs to |
| `masterIndex` | the master's zero-based position in `p:sldMasterIdLst` |
| `layoutIndex` | the layout's zero-based position within its master |

`appendSlides` takes the name or the handle. When two masters each have a layout with the same name, the name is ambiguous, so pass the handle:

```ts
const layout = deck.layouts().find((handle) => handle.name === 'Two Content' && handle.masterIndex === 1)
if (layout) await deck.appendSlides(pptx, { layout, at: 0 })
```

[Masters, layouts and themes](../reference/read-object-model.md#masters-layouts-and-themes) covers the model behind `deck.masters()`, with each master's colour map, theme and placeholders.

## Add generated slides

`deck.appendSlides(pptx, options)` builds the slides of a `TsPptx` and adds them to the deck. It returns a promise of the new slides. [`AppendSlidesOptions`](../reference/api/read/interfaces/AppendSlidesOptions.md) has three options:

| Option | Default | Effect |
| --- | --- | --- |
| `layout` | required | A layout name or a `LayoutHandle` from this deck. Every new slide binds to it. |
| `at` | append | Zero-based position of the first new slide. The others follow it. An out-of-range `at` appends. |
| `onMediaError` | `'throw'` | `'placeholder'` puts a placeholder in place of an `addImage` source that cannot be read. `'throw'` throws `MediaError` `media/load-failed`. |

- Each new slide's layout relationship points at the existing layout part. The call creates no master, layout or theme.
- Text, tables, images, charts with their workbooks, chartEx charts with their style and colour parts, embedded and linked audio and video, and speaker notes come across.
- A hyperlink to another slide of the same `TsPptx` points at the matching new slide. A link to a slide the call does not add throws `import/unresolved-slide-link`, before the deck changes.
- Fonts embedded with `pptx.embedFont` merge into the deck. [Carry fonts when importing slides](../embedded-fonts.md#carry-fonts-when-importing-slides) has the merge rules.
- Speaker notes bind to the deck's notes master. The generator's notes master is added only when the deck has none, as [One notes master per deck](read-and-edit.md#one-notes-master-per-deck) sets out.

The layout decides which theme and colour map a slide resolves against, and which layout PowerPoint shows the slide as based on. It does not place content. Generated slides keep the positions you gave them and take no geometry from the layout's placeholders. A scheme colour on a generated slide resolves against the template's theme.

### Match the slide size

The generator's slide size must equal the deck's, or `appendSlides` throws `import/slide-size-mismatch`. When the template uses a standard size, set `pptx.layout` to its name ([Layout units](../reference/layout-units.md) lists them). Otherwise read the size from the deck and define a layout from it:

```ts
const size = deck.slideSize
if (!size) throw new Error('the template declares no slide size')

const pptx = new TsPptx()
pptx.defineLayout({ name: 'TEMPLATE', width: size.widthIn, height: size.heightIn })
pptx.layout = 'TEMPLATE'
```

## Add slides to a deck that already has slides

`appendSlides` works on any loaded deck. Open it with `Presentation.load` to keep its slides:

```ts
const deck = await Presentation.load(await readFile('deck.pptx'))

const pptx = new TsPptx()
pptx.layout = 'LAYOUT_16x9'
pptx.addSlide().addText('Appendix', { x: 1, y: 1, w: 6, h: 1 })

const added = await deck.appendSlides(pptx, { layout: 'Blank' })
console.log(added.map((slide) => slide.index))
```

## Reuse slides the template file already has

To copy a slide that the template file itself holds, open the file twice. Use `fromTemplate` for the deck and `load` for the source. `importSlide` then binds the copy to the deck's own layout, master and theme instead of adding second copies of them. [Reuse chrome the deck already has](copy-between-decks.md#reuse-chrome-the-deck-already-has) has the rule.

## Compared with defineSlideMaster

`pptx.defineSlideMaster()` builds a master in code from the options it takes. `fromTemplate` keeps the template's master, layout and theme parts as the file stores them, including anything those options cannot express.

The template-anchored output of [`pptx-ts/script`](../reference/pptx-to-script.md#two-outputs) prints scripts that follow this path. They open the source deck with `fromTemplate` and add the slides with `appendSlides`.

## Invalid input

| Condition | Result | Code |
| --- | --- | --- |
| the input is not an OPC package | throws `PackageReadError` | `package/not-an-opc-package` |
| `layout` names no layout of the deck | throws `InvalidOptionError` | `layout/not-found` |
| `layout` names a layout that several masters have | throws `InvalidOptionError` | `layout/ambiguous-name` |
| `layout` is a `LayoutHandle` from another deck | throws `InvalidOptionError` | `layout/foreign-handle` |
| the generator's slide size differs from the deck's | throws `InvalidOptionError` | `import/slide-size-mismatch` |
| the deck declares no slide size | throws `InvalidOptionError` | `import/slide-size-unknown` |
| a generated slide links to a slide the call does not add | throws `InvalidOptionError` | `import/unresolved-slide-link` |
| an `addImage` source cannot be read, with `onMediaError: 'throw'` | throws `MediaError` | `media/load-failed` |

## Limits

- `appendSlides` does not rescale. The generator and the deck must have the same slide size.
- Generated slides do not take placeholder geometry from the layout they bind to.
- One call binds every slide to one layout. Call `appendSlides` once per layout.
- A generated slide can link only to slides added by the same call.
- `fromTemplate` removes every slide. To keep some, open the file with `Presentation.load` and call `removeSlide(index)` on the others.

## See also

- [Read and edit a deck](read-and-edit.md)
- [Copy slides between decks](copy-between-decks.md)
- [Embedded fonts](../embedded-fonts.md)
- [Layout units](../reference/layout-units.md)
- [Deck to script](../reference/pptx-to-script.md)
- [Errors and warnings](../errors-and-warnings.md)
- API reference: [`Presentation.fromTemplate`](../reference/api/read/classes/Presentation.md#fromtemplate), [`Presentation.layouts`](../reference/api/read/classes/Presentation.md#layouts), [`Presentation.appendSlides`](../reference/api/read/classes/Presentation.md#appendslides), [`AppendSlidesOptions`](../reference/api/read/interfaces/AppendSlidesOptions.md), [`FromTemplateOptions`](../reference/api/read/interfaces/FromTemplateOptions.md), [`LayoutHandle`](../reference/api/read/interfaces/LayoutHandle.md)
