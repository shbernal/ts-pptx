/**
 * What a page copy may share with the page it was copied from, and what it must own.
 *
 * Duplicating a page — `cloneSlide`, `importSlide` of the same page twice,
 * `importSlides` naming it twice — copies the page part and then has to decide,
 * for every part hanging off it, whether the copy points at the same part or gets
 * one of its own. Sharing is the cheap answer and is right for most of the graph:
 * a layout, a master, a theme, an image are deck-wide assets PowerPoint itself
 * stores once and points many shapes at.
 *
 * It is not right for the parts a page *owns*. A chart and its embedded workbook,
 * a SmartArt diagram's five parts, an OLE embedding, a notes slide: PowerPoint
 * treats each as belonging to exactly one page, and a package where two slides
 * resolve to one of them is one it refuses to open at all — the "the file or
 * directory is corrupted and unreadable" (`0x80070570`) dialog, with the whole
 * deck rejected rather than the duplicate page. The schema validator accepts such
 * a package: nothing in ECMA-376 says a chart part has one referrer, so this rule
 * is only visible against the application.
 *
 * Measured against desktop PowerPoint on the `mixed` fixture (a chart page and a
 * SmartArt page), duplicating a page that owns one of these:
 *   - sharing the chart part → refused;
 *   - a copy of the chart part sharing its embedded workbook and its
 *     `chartUserShapes` drawing → still refused;
 *   - the chart subtree copied whole → opens.
 * So ownership is transitive: a part a page owns owns its own subtree in turn,
 * down to the media blobs at the leaves.
 *
 * Hence the list below is of what may be **shared**, and everything else is
 * copied. The asymmetry is deliberate: a wrongly shared part is a deck nobody can
 * open, a wrongly copied one is some duplicated bytes. A new relationship type
 * that nobody has classified therefore lands on the safe side by default.
 */

import type { OpcPackage } from '../../opc/package.js'
import { relativePartName } from '../../opc/partnames.js'
import {
	AUDIO_REL,
	HYPERLINK_REL,
	IMAGE_REL,
	MS_MEDIA_REL,
	NOTES_MASTER_REL,
	OFFICE_REL,
	SLIDE_LAYOUT_REL,
	SLIDE_MASTER_REL,
	SLIDE_REL,
	TABLE_STYLES_REL,
	THEME_REL,
	VIDEO_REL,
} from '../../../ooxml/rel-types.js'

/** Relationship types whose target two copies of one page may point at together. */
const SHARED_BY_PAGE_COPIES: ReadonlySet<string> = new Set([
	// Deck furniture: reached through a page, owned by the deck.
	SLIDE_LAYOUT_REL,
	SLIDE_MASTER_REL,
	NOTES_MASTER_REL,
	THEME_REL,
	OFFICE_REL + 'themeOverride',
	OFFICE_REL + 'handoutMaster',
	OFFICE_REL + 'presProps',
	OFFICE_REL + 'viewProps',
	TABLE_STYLES_REL,
	// Package singletons that a page's own parts point back at — a comments part
	// names the deck's author list, which is one list for the whole package.
	OFFICE_REL + 'commentAuthors',
	'http://schemas.microsoft.com/office/2018/10/relationships/authors',
	// Media blobs. PowerPoint stores one copy and points every shape that shows it
	// at that copy, so sharing here is what the application does itself.
	IMAGE_REL,
	AUDIO_REL,
	VIDEO_REL,
	MS_MEDIA_REL,
	'http://schemas.microsoft.com/office/2017/06/relationships/model3d',
	OFFICE_REL + 'font',
	// Another page is its own page, never a part this one owns: a jump link points
	// at whatever copy of the target page the import decided on.
	SLIDE_REL,
	// External by nature; an internal one resolves to a page, handled above.
	HYPERLINK_REL,
])

/**
 * Whether a page copy may point at this relationship's target rather than taking
 * a copy of it. Everything not listed is owned by the page and copied — see the
 * module comment for why that is the safe default.
 */
export function isSharedByPageCopies(relType: string): boolean {
	return SHARED_BY_PAGE_COPIES.has(relType)
}

/**
 * Give a just-cloned page its own copy of every part the source page owned,
 * within one package. The clone's relationships start as a byte copy of the
 * source page's, so this walks them and repoints the owned ones at fresh parts;
 * shared targets are left alone.
 *
 * The subtree under each owned part is copied the same way, so a chart's workbook
 * and user-shapes drawing come along while the image inside that drawing stays
 * shared. A relationship *back* to the source page — a notes slide names the
 * slide it annotates — is repointed at the clone, which is what makes the copied
 * notes belong to it.
 *
 * A dangling relationship is left dangling rather than made to throw: cloning a
 * damaged deck is not this function's problem to discover.
 *
 * @param opc               the package both pages live in
 * @param sourcePartName    partname of the page that was cloned
 * @param clonePartName     partname of the clone, whose rels are still the source's
 */
export function duplicateOwnedTargets(opc: OpcPackage, sourcePartName: string, clonePartName: string): void {
	// Seeded with the page itself, so a back-reference to it lands on the clone.
	const copies = new Map<string, string>([[sourcePartName, clonePartName]])
	const rels = opc.relationshipsFor(clonePartName)
	// Snapshot: the loop repoints relationships, which rewrites the live set.
	for (const rel of Array.from(rels)) {
		if (rel.targetMode === 'External' || isSharedByPageCopies(rel.type)) continue
		const target = rels.resolveTarget(rel.id)
		const fresh = duplicateSubtree(opc, target, copies)
		if (fresh === target) continue
		rels.remove(rel.id)
		rels.addWithId(rel.id, rel.type, relativePartName(clonePartName, fresh))
	}
}

/**
 * Copy `partName` and, recursively, every part it owns, into fresh partnames.
 * `copies` dedupes within the one page copy, so a part two of the page's
 * relationships reach is copied once here even though the next page copy gets its
 * own. Returns the copy's partname, or `partName` unchanged when there is no such
 * part to copy.
 */
function duplicateSubtree(opc: OpcPackage, partName: string, copies: Map<string, string>): string {
	const already = copies.get(partName)
	if (already !== undefined) return already
	const part = opc.part(partName)
	if (!part) return partName

	const fresh = opc.reservePartNameLike(partName)
	opc.addPart(fresh, part.contentType, part.bytes)
	// Record before recursing, so a cycle (a notes slide naming its slide) terminates.
	copies.set(partName, fresh)

	const sourceRels = opc.relationshipsFor(partName)
	const freshRels = opc.relationshipsFor(fresh)
	for (const rel of sourceRels) {
		if (rel.targetMode === 'External') {
			freshRels.addWithId(rel.id, rel.type, rel.target, 'External')
			continue
		}
		const target = sourceRels.resolveTarget(rel.id)
		// A shared target is still routed through `copies`: that is how the notes
		// slide's `slide` relationship finds the clone instead of the original.
		const newTarget = isSharedByPageCopies(rel.type)
			? (copies.get(target) ?? target)
			: duplicateSubtree(opc, target, copies)
		freshRels.addWithId(rel.id, rel.type, relativePartName(fresh, newTarget))
	}
	return fresh
}
