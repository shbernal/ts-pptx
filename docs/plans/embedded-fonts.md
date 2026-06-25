# Plan: Embedded Font Support

Status: **draft / not started**
Branch: `feat/embedded-fonts`
Related backlog: `dn-importslide-v1-limits` (gap #1 → Feature A), new entry to be filed for Feature B.

## Goal

Add embedded-font support to PptxGenJS in two independent, separately-shippable
features:

- **Feature A — import-carry.** When `importSlide` (and friends) lift a slide
  from a source deck that already embeds fonts, bring the font parts across and
  merge `p:embeddedFontLst`. This is the actual backlog gap. Read/import only;
  fully testable in the Node suite.
- **Feature B — author-side embedding.** A public API to embed an arbitrary
  `.ttf`/`.otf` file when generating a deck from scratch. Larger surface; new
  binary parts, content types, presentation rels, and `p:embeddedFontLst`
  emission.

The two share the same OOXML target (the embedded-font part layout below), so
Feature A's read/merge helpers and Feature B's writer should converge on one
internal representation of "an embedded font set".

---

## 0. OOXML reference (verified against ECMA-376 transitional)

Embedded fonts in PresentationML consist of three coordinated pieces:

1. **Binary font parts** — one per face — at `/ppt/fonts/fontN.fntdata`.
   - Content type: `application/x-fontdata`.
   - Bytes are the **raw** TTF/OTF font file. PresentationML does *not* obfuscate
     embedded fonts (unlike WordprocessingML's `.odttf`).
   - `[Content_Types].xml` needs a `Default Extension="fntdata"
     ContentType="application/x-fontdata"` entry (one Default covers all font
     parts).

2. **Relationships** from `presentation.xml` — in
   `ppt/_rels/presentation.xml.rels` — one per face:
   - Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/font`
   - Target: `fonts/fontN.fntdata`

3. **`p:embeddedFontLst`** inside `presentation.xml`. Schema position matters:
   `CT_Presentation` child sequence puts `embeddedFontLst` at **index 7** —
   *after* `notesSz`, *before* `defaultTextStyle`. The current writer emits
   `notesSz` then jumps to `defaultTextStyle`, so the list slots cleanly between
   them.

   ```xml
   <p:embeddedFontLst>
     <p:embeddedFont>
       <p:font typeface="My Font"/>     <!-- CT_TextFont: required; typeface (+ optional panose/pitchFamily/charset) -->
       <p:regular   r:id="rIdN"/>       <!-- CT_EmbeddedFontDataId, 0..1 -->
       <p:bold      r:id="rIdM"/>       <!-- 0..1 -->
       <p:italic    r:id="rIdO"/>       <!-- 0..1 -->
       <p:boldItalic r:id="rIdP"/>      <!-- 0..1 -->
     </p:embeddedFont>
     <!-- ...more embeddedFont entries... -->
   </p:embeddedFontLst>
   ```

   - `CT_EmbeddedFontListEntry` child order: `font`, `regular`, `bold`,
     `italic`, `boldItalic`. Only `font` is required; the four face slots are
     each 0..1, and each carries the single `r:id` to its font part.
   - `p:presentation@saveSubsetFonts="1"` is **already emitted** today
     (`src/gen-xml.ts:2923`) but is currently inert. It declares "only the used
     glyph subset is embedded". We will NOT subset (we embed whole faces); set it
     to `"0"` when we actually embed, OR keep `"1"` only if we ever implement
     subsetting. Decide in Feature B; until then it is harmless because nothing
     is embedded.

---

## 1. Fixtures (do this first — gate for both features)

Per the fork's fixture-gated-work rule, the embedded-font XML we emit/merge must
be validated against *PowerPoint-authored* output, not synthetic XML.

### 1.1 Author the source fixture

Produce a real `.pptx` in Microsoft PowerPoint that embeds a font:

- File → Options → Save → **Embed fonts in the file** → *Embed all characters*
  (so faces are whole, not subset — matches our no-subset emission).
- Use a small, redistributable font to keep the fixture light and license-clean
  (e.g. an SIL OFL face). Embed at least a regular + bold pair so the
  multi-face `embeddedFont` path is exercised.
- One slide whose text actually uses the embedded face.

Store under `test/fixtures/` (mirror the existing fixture layout). Record:

- the part paths (`ppt/fonts/font1.fntdata`, …),
- the `presentation.xml.rels` font relationships,
- the verbatim `p:embeddedFontLst` block,

as the **oracle** for both the Feature A merge test and the Feature B emit test.

### 1.2 Extract a reusable font blob for Feature B

From the same fixture (or the original font file), keep one raw `.ttf`/`.otf`
under `test/fixtures/fonts/` to feed Feature B's authoring API in tests.

### 1.3 Validator availability

`pnpm run test:schema` needs `./tools/ooxml-validator/install.sh`. Confirm the
validator accepts `application/x-fontdata` Defaults and the `embeddedFontLst`
placement before relying on it in CI.

**Exit criterion for §1:** fixtures committed; the expected `embeddedFontLst` +
rels + part list written into the schema test as the comparison oracle.

---

## 2. Shared internal model

Define one representation both features use (location TBD — likely a small new
module `src/embedded-fonts.ts` or a section of `gen-xml.ts`):

```ts
interface EmbeddedFontFace {
  slot: 'regular' | 'bold' | 'italic' | 'boldItalic'
  bytes: Uint8Array          // raw font file
}
interface EmbeddedFont {
  typeface: string           // p:font/@typeface
  panose?: string
  pitchFamily?: number
  charset?: number
  faces: EmbeddedFontFace[]  // 1..4
}
```

- Feature A *reads* source `presentation.xml` into `EmbeddedFont[]`.
- Feature B *builds* `EmbeddedFont[]` from user calls.
- A single serializer turns `EmbeddedFont[]` + an rId allocator into the
  `p:embeddedFontLst` block + the rels + the part list.

---

## 3. Feature A — import-carry (`importSlide`)

Read/import only. Reuses the existing `#copyPart` registry machinery in
`src/read/api/presentation.ts`.

### 3.1 API

Add an opt-in flag so default behavior is unchanged:

```ts
importSlide(source, index, { embedFonts?: boolean })   // default false
```

(Apply the same flag to `importSlideMasters`/`importShape` if/when they need it;
start with `importSlide`.)

### 3.2 Mechanism

1. Stop skipping the font list on the *source* side: today
   `p:embeddedFontLst` is in the skip-lists at `presentation.ts:77` and `:1417`
   — those guard the slide/master/theme copy templates and should stay; the font
   carry is a *separate* traversal of the source `presentation.xml`, not part of
   the slide-part copy chain.
2. When `embedFonts` is set:
   - Read the source `OpcPackage` `presentation.xml` → parse `p:embeddedFontLst`
     and resolve each face `r:id` against `ppt/_rels/presentation.xml.rels`.
   - For each referenced font part, `#copyPart(source.opc, fontPartName)` — the
     existing registry already dedupes and reserves a fresh part name, so fonts
     shared across repeated imports copy exactly once.
   - Merge entries into the *target* `presentation.xml`'s `p:embeddedFontLst`
     (create it at the correct schema position if absent), rewriting each face
     `r:id` to a freshly-allocated relationship in the **target**
     `presentation.xml.rels`.
   - De-dupe by `typeface` + slot: if the target already embeds that face
     (e.g. from a prior import), reuse it rather than adding a duplicate
     `embeddedFont`/part.
3. Add the `fntdata` Default to the target `[Content_Types].xml` if not present.

### 3.3 Tests

- Unit (Node): import a slide from §1.1 fixture with `embedFonts: true`; assert
  the target package gains the font part(s), the rels, and a merged
  `p:embeddedFontLst` matching the oracle.
- Idempotency: importing twice yields one copy of each face.
- Default off: without the flag, no font parts/list appear (current behavior).
- Schema: round-trip the resulting package through `test:schema`.

### 3.4 Backlog

On completion, update `dn-importslide-v1-limits`: mark gap #1 implemented, set
`last_reviewed`, update `current_project_notes`/`evidence.local_files`. Run
`pnpm run backlog:validate`.

---

## 4. Feature B — author-side embedding

New public authoring API + write-side parts. Larger; ship after A.

### 4.1 API (proposed)

```ts
pptx.embedFont({
  path?: string                 // node fs path to .ttf/.otf
  data?: ArrayBuffer | Uint8Array | base64 string
  typeface: string              // family name as referenced by runs
  style?: 'regular' | 'bold' | 'italic' | 'boldItalic'   // default 'regular'
})
```

- Multiple calls with the same `typeface` and different `style` accumulate faces
  into one `embeddedFont` entry.
- Open question: derive `typeface`/style from the font's `name`/`OS/2` tables
  automatically vs. require the caller to declare them. **Decision: require the
  caller to declare** for v1 (no font-table parser dependency); revisit
  auto-detection later. Document that the declared `typeface` MUST match the
  family name used in `fontFace`/run typefaces or PowerPoint won't bind it.

### 4.2 Write-side wiring (all in the `write()` assembly, `src/pptxgen.ts:~670`)

1. Collect embedded fonts on the presentation (new field on the internal
   presentation model).
2. Emit each face as `ppt/fonts/fontN.fntdata` via `zip.add(path, bytes)`
   (same mechanism as media; STORE compression is fine since fonts are already
   compact binary — match the media path's compression choice).
3. `makeXmlContTypes` (`src/gen-xml.ts:2036`): add
   `<Default Extension="fntdata" ContentType="application/x-fontdata"/>` when any
   font is embedded.
4. `makeXmlPresentationRels` (`src/gen-xml.ts:2255`): append one `font`
   relationship per face with a stable rId allocation that does not collide with
   the slide/master rIds already computed there.
5. `makeXmlPresentation` (`src/gen-xml.ts:2919`): emit `p:embeddedFontLst`
   between `notesSz` (line ~2943) and `defaultTextStyle` (line ~2946), using the
   rIds allocated in step 4. Set `saveSubsetFonts` appropriately (see §0).

### 4.3 Licensing note

TTF `OS/2.fsType` carries embedding permission bits. v1: do not enforce — embed
what the caller hands us and document that the caller is responsible for font
licensing. (Optionally warn if `fsType` indicates no-embedding; defer.)

### 4.4 Tests

- Unit: `embedFont` a fixture face; write; reopen the zip; assert the part,
  Content-Types Default, font rel, and `embeddedFontLst` entry all present and
  matching the oracle.
- Multi-face: regular + bold under one `typeface` → one `embeddedFont`, two
  faces.
- Schema: `test:schema` fixture covering the emitted `embeddedFontLst`.

### 4.5 Backlog & changelog

- File a new `downstream-need` (or fork-feature) backlog entry for B before
  starting (per CLAUDE.md, candidate work is recorded in `docs/backlog.yml`).
- Record the new public API in `CHANGELOG.md` on implementation.

---

## 5. Sequencing

1. **§1 fixtures** (blocks everything).
2. **§2 shared model** (small).
3. **Feature A** (§3) — ship, update backlog.
4. **Feature B** (§4) — file backlog entry, ship, changelog.

A and B are independent after §1–2; A is the smaller, higher-value piece and
goes first.

## 6. Open questions

- `saveSubsetFonts`: keep `"1"` (implies subsetting we don't do) or switch to
  `"0"` when embedding whole faces? Lean `"0"` for honesty; confirm PowerPoint
  doesn't object.
- Auto-detect `typeface`/style from font tables (Feature B) — deferred to v2.
- Should Feature A also carry fonts in `importShape`/`importSlideMasters`, or is
  `importSlide` the only consumer that needs it? Start with `importSlide`.
- License-bit enforcement (`fsType`): warn vs. ignore. Lean ignore + document.
