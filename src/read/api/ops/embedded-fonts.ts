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
import { EMBEDDED_FONT_ENTRY_AFTER, PRESENTATION_AFTER_EMBEDDED_FONT_LST } from '../../../ooxml/sequence.js'
import { relativePartName } from '../../opc/partnames.js'
import { copyPart, type ImportContext } from './part-copy.js'
import type { Presentation } from '../presentation.js'
import { PackageReadError } from '../../../errors.js'
import { presentationRels, type DeckTarget } from './deck-target.js'

/** One `p:embeddedFont` entry as {@link readEmbeddedFontEntries} reads it. */
export interface EmbeddedFontEntry {
	typeface: string
	/** `p:font/@panose`, or `null` when the entry declares none. */
	panose: string | null
	/** `p:font` identity attrs other than `typeface` (panose/pitchFamily/charset), in document order. */
	identity: Array<{ name: string; value: string }>
	/**
	 * Every face slot that carries an `r:id`, in schema order. `partName` is the part that id
	 * resolves to, or `null` when it names no internal relationship of the presentation part.
	 */
	faces: Array<{ slot: EmbeddedFontSlot; relId: string; partName: string | null }>
}

/**
 * A deck's `p:embeddedFontLst`, read once.
 *
 * The getter, the import check and the carry each walked the list themselves and disagreed on a
 * face whose `r:id` names no relationship: the getter skipped it, the other two called
 * `resolveTarget` and threw an option error. Reading it here, with the dangling face reported as
 * `partName: null` rather than thrown, lets each consumer decide what that face means for it. An
 * entry with no `@typeface` and a face slot with no `r:id` are left out, as all three always did.
 */
export function readEmbeddedFontEntries(deck: DeckTarget): EmbeddedFontEntry[] {
	const root = deck.presentationPart.dom.documentElement
	const lst = root && firstChild(root, 'p:embeddedFontLst')
	if (!lst) return []
	const rels = presentationRels(deck)
	const entries: EmbeddedFontEntry[] = []
	for (const entry of getElements(lst, 'p:embeddedFont')) {
		const font = firstChild(entry, 'p:font')
		const typeface = font && attr(font, 'typeface')
		if (!font || !typeface) continue
		const identity: EmbeddedFontEntry['identity'] = []
		for (const name of ['panose', 'pitchFamily', 'charset']) {
			const value = attr(font, name)
			if (value !== null) identity.push({ name, value })
		}
		const faces: EmbeddedFontEntry['faces'] = []
		for (const slot of EMBEDDED_FONT_SLOTS) {
			const face = firstChild(entry, `p:${slot}`)
			const relId = face && attr(face, 'r:id')
			if (!relId) continue
			const rel = rels.get(relId)
			faces.push({ slot, relId, partName: rel && rel.targetMode !== 'External' ? rels.resolveTarget(relId) : null })
		}
		entries.push({ typeface, panose: attr(font, 'panose'), identity, faces })
	}
	return entries
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
 * Dry-run {@link carryEmbeddedFonts} against one source, reading only the *source* package:
 * every face this deck would copy names a part that is actually there.
 *
 * Every import that carries fonts calls it before it changes anything, because the carry runs
 * last — after the page or the masters are already in the deck — and a face it cannot copy would
 * throw with parts added and no way back. A face whose `r:id` names no relationship and a face
 * whose relationship names no part are the same failure to a caller, the face's binary is not in
 * the package, so both are `package/part-missing`.
 *
 * It reads the list through {@link readEmbeddedFontEntries}, as the carry does, so the two cannot
 * skip different entries. The one throw the carry has past this point is
 * `package/part-has-no-root` on the *destination*.
 * @param source - the deck whose fonts would be carried
 * @param api - the public method asking, which opens the error message
 */
export function checkEmbeddedFontsCopyable(source: Presentation, api = 'importSlides'): void {
	for (const { typeface, faces } of readEmbeddedFontEntries(source)) {
		for (const { slot, relId, partName } of faces) {
			if (partName === null)
				throw new PackageReadError(
					'package/part-missing',
					`${api}: embedded font ${typeface} (${slot}) names relationship ${relId}, which ${source.presentationPart.partName} does not have`
				)
			if (!source.opc.part(partName))
				throw new PackageReadError(
					'package/part-missing',
					`${api}: source package has no part ${partName} (embedded font ${typeface}, ${slot})`
				)
		}
	}
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
	const incoming: IncomingEmbeddedFont[] = readEmbeddedFontEntries(source).map(({ typeface, identity, faces }) => ({
		typeface,
		identity,
		// Every caller runs `checkEmbeddedFontsCopyable` first, so a face with no part cannot reach
		// here. Binary comes across via copyPart, so the per-source registry dedupes faces shared
		// across repeated imports; the thunk runs only when the face is added.
		faces: faces.flatMap(({ slot, partName }) =>
			partName === null ? [] : [{ slot, createPart: () => copyPart(ctx, partName) }]
		),
	}))
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
function mergeEmbeddedFontEntries(dest: Presentation, entries: IncomingEmbeddedFont[]): void {
	if (entries.length === 0) return

	const presPart = dest.presentationPart
	const presRoot = presPart.dom.documentElement
	if (!presRoot)
		throw new PackageReadError(
			'package/part-has-no-root',
			'presentation.xml has no document element to carry embedded fonts into'
		)
	const presRels = presentationRels(dest)

	const targetLst = getOrAddChild(presRoot, 'p:embeddedFontLst', PRESENTATION_AFTER_EMBEDDED_FONT_LST)
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
			insertInOrder(targetEntry, targetFace, EMBEDDED_FONT_ENTRY_AFTER[`p:${face.slot}`] ?? [])
			copiedAny = true
		}
	}

	if (copiedAny) presPart.markDirty()
}
