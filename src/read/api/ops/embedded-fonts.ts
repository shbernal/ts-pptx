/**
 * Merging embedded font faces into a deck's `p:embeddedFontLst`.
 *
 * Fonts arrive two ways — copied out of another `Presentation`, or handed over as raw bytes by a
 * generator during an append — and the two differ only in how the binary part comes into being.
 * That difference is confined to a per-face thunk, so both paths share one merge core: entries
 * join by typeface, faces de-dupe by slot, and the thunk runs only for a face actually being
 * added, so a face the deck already embeds never mints an orphan part.
 */

import {
	attr,
	createElement,
	firstChild,
	getElements,
	getOrAddChild,
	insertInOrder,
	setAttr,
	type Element,
} from '../../oxml/dom.js'
import {
	EMBEDDED_FONT_SLOTS,
	FONT_DATA_CONTENT_TYPE,
	FONT_DATA_EXTENSION,
	FONT_REL_TYPE,
	type EmbeddedFont,
	type EmbeddedFontSlot,
} from '../../../embedded-fonts.js'
import { relativePartName } from '../../opc/partnames.js'
import { copyPart, type ImportContext } from '../part-copy.js'
import type { Presentation } from '../presentation.js'
import { PackageReadError } from '../../../errors.js'

/**
 * `p:embeddedFontLst`'s document-order successors in `CT_Presentation` (index 7,
 * after `smartTags`): everything that may legally follow it, so a created list
 * lands in the right slot when the deck has none yet.
 */
const PRESENTATION_EMBEDDED_FONT_LST_SUCCESSORS = [
	'p:custShowLst',
	'p:photoAlbum',
	'p:custDataLst',
	'p:kinsoku',
	'p:defaultTextStyle',
	'p:modifyVerifier',
	'p:extLst',
]

/**
 * A face slot's document-order successors in `CT_EmbeddedFontListEntry`
 * (`font`, `regular`, `bold`, `italic`, `boldItalic`), so a newly-inserted face
 * keeps the schema's child order regardless of which slots already exist.
 */
const EMBEDDED_FONT_FACE_SUCCESSORS: Record<EmbeddedFontSlot, string[]> = {
	regular: ['p:bold', 'p:italic', 'p:boldItalic'],
	bold: ['p:italic', 'p:boldItalic'],
	italic: ['p:boldItalic'],
	boldItalic: [],
}

/**
 * One typeface's faces normalized for the embedded-font merge core (`#mergeEmbeddedFontEntries`):
 * the `p:font` identity attributes plus, per face slot, a thunk that creates the
 * binary font part on demand and returns its partname. The thunk runs only for a
 * face actually being added (after the typeface+slot de-dupe), so no orphan part is
 * created for a face the deck already embeds. Lets the import-side (copy a part out
 * of a source package) and append-side (write raw generator bytes) callers share one
 * merge core while differing only in how the binary part is produced.
 */
interface IncomingEmbeddedFont {
	typeface: string
	/** `p:font` identity attrs other than `typeface` (panose/pitchFamily/charset), in document order. */
	identity: Array<{ name: string; value: string }>
	faces: Array<{ slot: EmbeddedFontSlot; createPart: () => string }>
}

/**
 * Copy `source`'s embedded fonts into this deck and merge them into our
 * `p:embeddedFontLst`. Font binaries come across via {@link copyPart} (so the
 * per-source registry dedupes faces shared across repeated imports); entries are
 * merged by `typeface` + face slot, so a face this deck already embeds is reused
 * rather than duplicated. No-op when the source embeds no fonts. See
 * {@link ImportSlideOptions.embedFonts}.
 */
export function carryEmbeddedFonts(dest: Presentation, source: Presentation, ctx: ImportContext): void {
	const sourceRoot = source.presentationPart.dom.documentElement
	const sourceLst = sourceRoot && firstChild(sourceRoot, 'p:embeddedFontLst')
	const sourceEntries = sourceLst ? getElements(sourceLst, 'p:embeddedFont') : []
	if (sourceEntries.length === 0) return

	const sourcePresRels = source.opc.relationshipsFor(source.presentationPart.partName)
	const incoming: IncomingEmbeddedFont[] = []
	for (const srcEntry of sourceEntries) {
		const srcFont = firstChild(srcEntry, 'p:font')
		const typeface = srcFont ? attr(srcFont, 'typeface') : null
		if (!srcFont || !typeface) continue

		// Copy the source p:font identity attributes (panose/pitchFamily/charset).
		const identity: IncomingEmbeddedFont['identity'] = []
		for (const name of ['panose', 'pitchFamily', 'charset']) {
			const value = attr(srcFont, name)
			if (value !== null) identity.push({ name, value })
		}

		const faces: IncomingEmbeddedFont['faces'] = []
		for (const slot of EMBEDDED_FONT_SLOTS) {
			const srcFace = firstChild(srcEntry, `p:${slot}`)
			const srcRid = srcFace && attr(srcFace, 'r:id')
			if (!srcFace || !srcRid) continue
			// Binary comes across via copyPart, so the per-source registry dedupes faces
			// shared across repeated imports; the thunk runs only when the face is added.
			faces.push({ slot, createPart: () => copyPart(ctx, sourcePresRels.resolveTarget(srcRid)) })
		}
		incoming.push({ typeface, identity, faces })
	}
	mergeEmbeddedFontEntries(dest, incoming)
}

/**
 * Carry a generator's presentation-level embedded fonts ({@link ExtractedSlides.embeddedFonts},
 * from `pptx.embedFont`) into this deck during {@link Presentation.appendSlides}. Each face's raw bytes are
 * written as a fresh `/ppt/fonts/fontN.fntdata` part; merge/de-dupe by typeface + slot is shared
 * with {@link carryEmbeddedFonts} via {@link mergeEmbeddedFontEntries}, so appending the same
 * generator twice (or onto a deck that already embeds the face) carries each face once.
 */
export function carryGeneratedEmbeddedFonts(dest: Presentation, fonts: EmbeddedFont[]): void {
	const incoming: IncomingEmbeddedFont[] = []
	for (const font of fonts) {
		if (!font.typeface) continue
		const identity: IncomingEmbeddedFont['identity'] = []
		if (font.panose !== undefined) identity.push({ name: 'panose', value: font.panose })
		if (font.pitchFamily !== undefined) identity.push({ name: 'pitchFamily', value: String(font.pitchFamily) })
		if (font.charset !== undefined) identity.push({ name: 'charset', value: String(font.charset) })

		const faces: IncomingEmbeddedFont['faces'] = []
		for (const slot of EMBEDDED_FONT_SLOTS) {
			const face = font.faces.find((f) => f.slot === slot)
			if (!face?.bytes) continue
			const bytes = face.bytes
			faces.push({
				slot,
				createPart: () => {
					const partName = dest.opc.reservePartNameLike('/ppt/fonts/font1.fntdata')
					dest.opc.addPart(partName, FONT_DATA_CONTENT_TYPE, bytes)
					return partName
				},
			})
		}
		if (faces.length > 0) incoming.push({ typeface: font.typeface, identity, faces })
	}
	mergeEmbeddedFontEntries(dest, incoming)
}

/**
 * Merge normalized {@link IncomingEmbeddedFont} entries into this deck's
 * `p:embeddedFontLst` — the shared core of {@link carryEmbeddedFonts} (import-side)
 * and {@link carryGeneratedEmbeddedFonts} (append-side). Entries merge by `typeface`,
 * faces de-dupe by slot (a face this deck already embeds is left as is). For each newly
 * added face the `fntdata` Default is ensured, the binary part is created via the face's
 * `createPart` thunk, a `font` rel is added to presentation.xml, and the `p:<slot>` element
 * is inserted in schema child order. The list is created at CT_Presentation index 7 when
 * the deck has none yet. No-op for empty input.
 */
export function mergeEmbeddedFontEntries(dest: Presentation, entries: IncomingEmbeddedFont[]): void {
	if (entries.length === 0) return

	const presPart = dest.presentationPart
	const presRoot = presPart.dom.documentElement
	if (!presRoot)
		throw new PackageReadError(
			'package/part-has-no-root',
			'presentation.xml has no document element to carry embedded fonts into'
		)
	const presRels = dest.opc.relationshipsFor(presPart.partName)

	const targetLst = getOrAddChild(presRoot, 'p:embeddedFontLst', PRESENTATION_EMBEDDED_FONT_LST_SUCCESSORS)
	const targetByTypeface = new Map<string, Element>()
	for (const entry of getElements(targetLst, 'p:embeddedFont')) {
		const font = firstChild(entry, 'p:font')
		const typeface = font && attr(font, 'typeface')
		if (typeface) targetByTypeface.set(typeface, entry)
	}

	let copiedAny = false
	for (const incoming of entries) {
		// Find or create the target entry for this typeface, carrying its
		// p:font identity attributes (typeface + optional panose/pitchFamily/charset).
		let targetEntry = targetByTypeface.get(incoming.typeface)
		if (!targetEntry) {
			targetEntry = createElement(presPart.dom, 'p:embeddedFont')
			const targetFont = createElement(presPart.dom, 'p:font')
			setAttr(targetFont, 'typeface', incoming.typeface)
			for (const { name, value } of incoming.identity) setAttr(targetFont, name, value)
			targetEntry.appendChild(targetFont)
			targetLst.appendChild(targetEntry)
			targetByTypeface.set(incoming.typeface, targetEntry)
		}

		for (const face of incoming.faces) {
			if (firstChild(targetEntry, `p:${face.slot}`)) continue // de-dupe: face already present

			// Ensure the fntdata Default exists *before* creating the part, so addPart
			// resolves the content type via the Default (no per-part Override).
			dest.opc.contentTypes.ensureDefault(FONT_DATA_EXTENSION, FONT_DATA_CONTENT_TYPE)
			const newFontPart = face.createPart()
			const relId = presRels.add(FONT_REL_TYPE, relativePartName(presPart.partName, newFontPart)).id

			const targetFace = createElement(presPart.dom, `p:${face.slot}`)
			setAttr(targetFace, 'r:id', relId)
			insertInOrder(targetEntry, targetFace, EMBEDDED_FONT_FACE_SUCCESSORS[face.slot])
			copiedAny = true
		}
	}

	if (copiedAny) presPart.markDirty()
}
