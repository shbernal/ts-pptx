---
doc-schema-version: 1
title: "Embedded Fonts"
summary: "How PptxGenJS embeds whole font faces — author-side pptx.embedFont() and import-carry importSlide({ embedFonts: true }) — the shared OOXML model behind both, and the PowerPoint-authored oracle."
read_when:
  - Changing embedded-font emit or merge (src/embedded-fonts.ts and its callers)
  - Touching pptx.embedFont() or importSlide({ embedFonts })
  - Adjusting presentation.xml font wiring (embeddedFontLst, font rels, fntdata Default)
  - Regenerating or interpreting the embedded-fonts fixtures
doc_type: "decision"
---

# Embedded Fonts

## Status

**Shipped (2026-06-25).** A deck can carry whole font faces so it renders with
them on machines that lack the font, mirroring PowerPoint's *Save → Embed fonts
in the file*. Two independent entry points share one OOXML model:

- **Author-side** — `await pptx.embedFont({ path | data, typeface, style })`
  embeds a face when generating a deck from scratch.
  Source: `src/pptxgen.ts` (`embedFont`, `_embeddedFonts`), write-side wiring in
  `src/gen/pres/presentation.ts`, field on the internal model in `src/core-interfaces.ts`.
- **Import-carry** — `importSlide(source, i, { embedFonts: true })` brings a
  source deck's presentation-level embedded fonts across when lifting a slide.
  Source: `src/read/api/presentation.ts` (`#carryEmbeddedFonts`).

Both converge on the shared model and serializer in `src/embedded-fonts.ts`.
See `CHANGELOG.md` for the import-carry limits. Tests:
`test/regression/embed-font.test.js` (author-side), `test/read/embedded-fonts.test.js`
(import-carry, incl. schema validity), and an author-side case in
`test/schema-cases.js` (validator-checked against the oracle).

## OOXML target (ECMA-376 transitional)

Embedded fonts are three coordinated pieces:

1. **Binary font parts** — one per face — at `/ppt/fonts/fontN.fntdata`. The bytes
   are the **raw** TTF/OTF file: PresentationML does *not* obfuscate embedded
   fonts (unlike WordprocessingML's `.odttf`). `[Content_Types].xml` carries a
   single `Default Extension="fntdata" ContentType="application/x-fontdata"`
   covering every font part.
2. **Relationships** from `presentation.xml` (in `ppt/_rels/presentation.xml.rels`),
   one per face, type `…/relationships/font`, target `fonts/fontN.fntdata`.
3. **`p:embeddedFontLst`** inside `presentation.xml`, at **index 7** of the
   `CT_Presentation` child sequence — after `notesSz`, before `defaultTextStyle`.

```xml
<p:embeddedFontLst>
  <p:embeddedFont>
    <p:font typeface="My Font"/>   <!-- CT_TextFont: typeface required (+ optional panose/pitchFamily/charset) -->
    <p:regular    r:id="rIdN"/>    <!-- CT_EmbeddedFontDataId, 0..1 -->
    <p:bold       r:id="rIdM"/>    <!-- 0..1 -->
    <p:italic     r:id="rIdO"/>    <!-- 0..1 -->
    <p:boldItalic r:id="rIdP"/>    <!-- 0..1 -->
  </p:embeddedFont>
</p:embeddedFontLst>
```

`CT_EmbeddedFontListEntry` child order is `font`, `regular`, `bold`, `italic`,
`boldItalic`. Only `font` is required; each of the four face slots is `0..1` and
carries the single `r:id` to its font part. The constant `EMBEDDED_FONT_SLOTS`
fixes this order so the read- and write-side emitters agree.

We embed **whole** faces, never a glyph subset, so on `p:presentation` we set
`embedTrueTypeFonts="1" saveSubsetFonts="0"` whenever any face has bytes. When no
font is embedded, output is unchanged: the historical inert `saveSubsetFonts="1"`
stays and `embedTrueTypeFonts` is absent.

## Shared model (`src/embedded-fonts.ts`)

One representation both features build/consume:

```ts
interface EmbeddedFontFace { slot: EmbeddedFontSlot; bytes?: Uint8Array }
interface EmbeddedFont { typeface: string; panose?; pitchFamily?; charset?; faces: EmbeddedFontFace[] }
```

The module owns only OOXML-shape knowledge; rId allocation and part placement stay
with each caller (their packaging models differ). It exposes:

- `flattenEmbeddedFaces(fonts, firstRId)` → ordered `FlatEmbeddedFace[]` assigning
  sequential 1-based part indices and rIds; faces without bytes are skipped. Used
  by the write path so the part writer, rels writer, and list emitter all agree.
- `serializeEmbeddedFontLst(fonts, rIdForFace)` → the `<p:embeddedFontLst>` string
  (empty when no face has an allocated rId), assuming the enclosing doc declares
  the `p:`/`r:` prefixes.
- The constants `FONT_DATA_EXTENSION`, `FONT_DATA_CONTENT_TYPE`, `FONT_REL_TYPE`,
  `EMBEDDED_FONT_SLOTS`.

## Author-side: `pptx.embedFont`

```ts
await pptx.embedFont({ path: '/fonts/Silkscreen-Regular.ttf', typeface: 'Silkscreen' })
await pptx.embedFont({ path: '/fonts/Silkscreen-Bold.ttf', typeface: 'Silkscreen', style: 'bold' })
slide.addText('hi', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen' })
```

- `path` is a path/URL loaded via the runtime adapter (`loadFontData`); `data` is
  in-memory bytes — `Uint8Array`, `ArrayBuffer`, or a base64 string (with or
  without a data-URL prefix).
- `style` is one of `'regular'` (default) | `'bold'` | `'italic'` | `'boldItalic'`.
  Repeated calls with the same `typeface` and different `style` accumulate into one
  `p:embeddedFont` entry; a repeat of the same `typeface`+`style` replaces the
  prior bytes (last call wins).
- **The declared `typeface` MUST match the family name your runs/`fontFace` use**
  or PowerPoint won't bind the embedded face.
- Validates input: missing `typeface`, missing byte source, and an invalid `style`
  all throw.

At write time (in `src/gen/pres/presentation.ts`) the accumulated `_embeddedFonts` drive:
`makeXmlContTypes` (the `fntdata` Default), `makeXmlPresentationRels` (one `font`
rel per face, rIds allocated after the slide/master rels), `makeXmlPresentation`
(the `p:embeddedFontLst` between `notesSz` and `defaultTextStyle`, plus the
`embedTrueTypeFonts`/`saveSubsetFonts` flags), and the `zip.add` of each
`/ppt/fonts/fontN.fntdata` part (STORE-compressed — fonts are already compact).

## Import-carry: `importSlide({ embedFonts: true })`

Opt-in (default off, so existing behaviour is unchanged). When set,
`#carryEmbeddedFonts` runs a **separate** traversal of the source
`presentation.xml` (not part of the slide-part copy chain, which still skips
`p:embeddedFontLst`):

1. Parse the source `p:embeddedFontLst`; for each face resolve its `r:id` against
   the source `presentation.xml.rels`.
2. `ensureDefault('fntdata', …)` **before** copying, so `#copyPart`'s `addPart`
   resolves the content type via the Default (one Default, no per-part Override).
3. Copy each referenced font part via `#copyPart` — the per-source registry
   dedupes, so faces shared across repeated imports copy exactly once — and add a
   fresh `font` relationship in the target `presentation.xml.rels`.
4. Merge into the target `p:embeddedFontLst` (created at index 7 if absent),
   cloning the source `p:font` identity (`typeface` + optional
   `panose`/`pitchFamily`/`charset`) and inserting each face slot in schema order.
5. De-dupe by `typeface` + face slot: a face this deck already embeds is reused,
   not duplicated — so importing the same slide twice carries each face once.

## Oracle & fixtures

Per the fork's fixture-gated-work rule, the emitted/merged XML is validated against
**PowerPoint-authored** output, not synthetic XML:

- `test/read/fixtures/embedded-fonts.pptx` — a real PowerPoint deck that embeds
  Silkscreen regular + bold (whole characters), one slide using the face. Its
  verbatim `embeddedFontLst`, font rels, and part list are captured in
  `test/read/fixtures/embedded-fonts.oracle.json`, the comparison oracle for both
  the import-carry merge test and the author-side emit test.
- `test/read/fixtures/fonts/Silkscreen-Regular.ttf` / `Silkscreen-Bold.ttf` —
  raw redistributable (SIL OFL) faces fed to the author-side API in tests.

`pnpm run test:schema` (validator installed via `./tools/ooxml-validator/install.sh`)
confirms the validator accepts the `fntdata` Default and the `embeddedFontLst`
placement.

## Standing caveats

- **Font licensing is the caller's responsibility.** TTF `OS/2.fsType` carries
  embedding-permission bits; v1 does not enforce them — it embeds whatever bytes
  the caller hands over.
- **No subsetting.** We embed whole faces, so `saveSubsetFonts="0"`. Subsetting
  (and the matching `saveSubsetFonts="1"`) is not implemented.
- **No auto-detection of `typeface`/style** from the font's `name`/`OS/2` tables —
  the caller declares them (no font-table parser dependency). Revisit if needed.
- Import-carry is wired for `importSlide`; `importShape`/`importSlideMasters` do
  not carry fonts (no consumer needs it yet).
