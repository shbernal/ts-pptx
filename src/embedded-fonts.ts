/**
 * Shared model + serializer for PresentationML embedded fonts, used by both the
 * write side (author-side `pptx.embedFont`, `src/gen-xml.ts`) and the read side
 * (`importSlide({ embedFonts: true })`, `src/read/api/presentation.ts`).
 *
 * Embedded fonts are three coordinated pieces (ECMA-376 transitional):
 *  1. binary font parts `/ppt/fonts/fontN.fntdata` — the **raw** TTF/OTF bytes
 *     (PresentationML does not obfuscate embedded fonts), content type
 *     `application/x-fontdata` via a single `fntdata` Default extension;
 *  2. one `font` relationship per face from `presentation.xml`;
 *  3. a `p:embeddedFontLst` in `presentation.xml`, at index 7 of the
 *     `CT_Presentation` child sequence (after `smartTags`, before `custShowLst`
 *     … `defaultTextStyle`).
 *
 * This module owns only the OOXML-shape knowledge; rId allocation and part
 * placement stay with each caller (their packaging models differ).
 */

/** Extension for the binary font parts (`/ppt/fonts/fontN.fntdata`). */
export const FONT_DATA_EXTENSION = 'fntdata'
/** Content type for a `.fntdata` part — one `Default` covers every font part. */
export const FONT_DATA_CONTENT_TYPE = 'application/x-fontdata'
/** Relationship type from `presentation.xml` to a font part (one per face). */
export const FONT_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'

/**
 * The four face slots of `CT_EmbeddedFontListEntry`, in schema child order
 * (`font` first, then these). Iterate in this order everywhere so the read- and
 * write-side emitters agree on element order.
 */
export const EMBEDDED_FONT_SLOTS = ['regular', 'bold', 'italic', 'boldItalic'] as const
export type EmbeddedFontSlot = (typeof EMBEDDED_FONT_SLOTS)[number]

/** One embeddable face: a slot plus the raw font bytes (absent in pure read-side merges). */
export interface EmbeddedFontFace {
	slot: EmbeddedFontSlot
	bytes?: Uint8Array
}

/** A typeface family with 1..4 embedded faces — one `p:embeddedFont` entry. */
export interface EmbeddedFont {
	/** `p:font/@typeface` — MUST match the family name used in run/`fontFace` typefaces or PowerPoint won't bind it. */
	typeface: string
	panose?: string
	pitchFamily?: number
	charset?: number
	faces: EmbeddedFontFace[]
}

/**
 * One face flattened to its package coordinates. `partIndex` is 1-based and maps
 * to `/ppt/fonts/font${partIndex}.fntdata`; `rId` is the numeric relationship id
 * (use `rId${rId}`). Produced once and shared by the part writer, the rels
 * writer, and the `embeddedFontLst` emitter so all three agree.
 */
export interface FlatEmbeddedFace {
	fontIndex: number
	slot: EmbeddedFontSlot
	bytes: Uint8Array
	partIndex: number
	rId: number
}

/**
 * Flatten `fonts` into one ordered face list, assigning sequential part indices
 * (1-based) and relationship ids (starting at `firstRId`). Faces are emitted in
 * `fonts` order, each font's faces in {@link EMBEDDED_FONT_SLOTS} order. Faces
 * without `bytes` are skipped (nothing to write).
 */
export function flattenEmbeddedFaces(fonts: EmbeddedFont[], firstRId: number): FlatEmbeddedFace[] {
	const flat: FlatEmbeddedFace[] = []
	let partIndex = 0
	let rId = firstRId
	fonts.forEach((font, fontIndex) => {
		for (const slot of EMBEDDED_FONT_SLOTS) {
			const face = font.faces.find(f => f.slot === slot)
			if (!face || !face.bytes) continue
			flat.push({ fontIndex, slot, bytes: face.bytes, partIndex: ++partIndex, rId: rId++ })
		}
	})
	return flat
}

/** Minimal XML attribute escaper (kept local so this module stays dependency-free). */
function escapeAttr(value: string): string {
	return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/**
 * Serialize a `<p:embeddedFontLst>` from `fonts` and a `(fontIndex, slot) → rId`
 * lookup (typically backed by {@link flattenEmbeddedFaces}). Returns `''` when no
 * font has any face with an allocated rId, so callers can emit nothing. Assumes
 * the enclosing document declares the `p:` and `r:` prefixes (presentation.xml
 * does), so no namespace declarations are emitted here.
 */
export function serializeEmbeddedFontLst(fonts: EmbeddedFont[], rIdForFace: (fontIndex: number, slot: EmbeddedFontSlot) => number | undefined): string {
	const entries: string[] = []
	fonts.forEach((font, fontIndex) => {
		const faceXml: string[] = []
		for (const slot of EMBEDDED_FONT_SLOTS) {
			const rId = rIdForFace(fontIndex, slot)
			if (rId === undefined) continue
			faceXml.push(`<p:${slot} r:id="rId${rId}"/>`)
		}
		if (faceXml.length === 0) return
		let fontAttrs = `typeface="${escapeAttr(font.typeface)}"`
		if (font.panose !== undefined) fontAttrs += ` panose="${escapeAttr(font.panose)}"`
		if (font.pitchFamily !== undefined) fontAttrs += ` pitchFamily="${font.pitchFamily}"`
		if (font.charset !== undefined) fontAttrs += ` charset="${font.charset}"`
		entries.push(`<p:embeddedFont><p:font ${fontAttrs}/>${faceXml.join('')}</p:embeddedFont>`)
	})
	if (entries.length === 0) return ''
	return `<p:embeddedFontLst>${entries.join('')}</p:embeddedFontLst>`
}
