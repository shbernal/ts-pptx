/**
 * Recognising a part the destination deck **already holds**, so an import binds to it
 * instead of copying it in again.
 *
 * The case this exists for is the deck templated from its own source. `fromTemplate`
 * keeps a package's masters, layouts and theme byte-identical and strips only its
 * slides, so a script that then imports a slide back out of that same file hands
 * `copyPart` a `slideLayout -> slideMaster -> theme` subgraph whose every part is
 * already in the destination under its own partname. The copy registry maps *source*
 * partnames to copies this import made and knows nothing about what the destination
 * arrived with, so each of those parts was copied again: the output grew a duplicate
 * layout and master per imported slide, visible in PowerPoint's layout picker and
 * enough to make a later `appendSlides({ layout: <name> })` ambiguous.
 *
 * The test is deliberately strict, because the cost of getting it wrong is a slide
 * bound to chrome that only resembles what it asked for. A part is reusable only when
 * the destination holds it at the *same partname*, with the same content type, the same
 * bytes (as they would be written, so an edited destination part disqualifies itself),
 * the same relationships, and the same rule applied recursively to everything those
 * relationships reach. Byte-equality all the way down means "reuse" and "copy" produce
 * the same rendering; anything less falls back to the copy, which is always correct and
 * merely wasteful.
 *
 * Two further conditions have nothing to do with identity and everything to do with the
 * destination being a real deck: a reused master must already be registered in
 * `presentation.xml`, and a reused layout must already sit in a registered master's
 * `p:sldLayoutIdLst`. An identical part that is an orphan in the destination is not
 * something to bind a slide to, and the copy path (which registers and links what it
 * copies) is the right answer there.
 *
 * Reuse mutates nothing, which is what lets it hold across repeated imports: the
 * destination's chrome stays byte-identical, so the second import recognises it exactly
 * as the first one did.
 */

import type { OpcPackage } from '../../opc/package.js'
import type { DeckTarget } from './deck-target.js'
import { SLIDE_CONTENT_TYPE, SLIDE_LAYOUT_CONTENT_TYPE, SLIDE_MASTER_CONTENT_TYPE } from '../../../ooxml/rel-types.js'
import { copyTraversalStep } from './copy-traversal.js'
import { layoutPartNamesOf, slideMasterPartNames } from './part-index.js'

/**
 * Whether the destination already holds `sourcePartName`'s subgraph, part for part and
 * byte for byte, in a state an imported slide can bind to.
 *
 * @param dest            the deck being imported into
 * @param source          the package being imported out of
 * @param sourcePartName  the part the copy is about to make
 * @return                `true` when the copy can be skipped and the existing part used
 */
export function destinationAlreadyHolds(dest: DeckTarget, source: OpcPackage, sourcePartName: string): boolean {
	const part = dest.opc.part(sourcePartName)
	// A page is the one part an import never shares, whatever its bytes say: two
	// `p:sldId` entries pointing at one part is a package PowerPoint refuses to open.
	if (!part || part.contentType === SLIDE_CONTENT_TYPE) return false
	if (!identicalSubgraph(dest.opc, source, sourcePartName, new Set())) return false
	return wiredIntoDeck(dest, sourcePartName)
}

/**
 * `partName` and everything it reaches are the same in both packages: same content type,
 * same body bytes, same relationships, recursively.
 *
 * The destination side is compared as it would be **written** (`serialize()`), not as it
 * was read, so a part some earlier edit marked dirty is treated as different even when
 * the edit was a no-op. Under-reusing costs a duplicated part; over-reusing binds a slide
 * to something else's chrome.
 *
 * The traversal skips exactly what `copyPart` skips, because it asks `copyTraversalStep` --
 * so it is about the subgraph the copy would actually have made rather than about a rule
 * transcribed from it.
 */
function identicalSubgraph(dest: OpcPackage, source: OpcPackage, partName: string, visited: Set<string>): boolean {
	if (visited.has(partName)) return true
	visited.add(partName)

	const sourcePart = source.part(partName)
	const destPart = dest.part(partName)
	if (!sourcePart || !destPart) return false
	if (sourcePart.contentType !== destPart.contentType) return false
	if (!bytesEqual(sourcePart.bytes, destPart.serialize())) return false

	const sourceRels = [...source.relationshipsFor(partName)]
	const destRels = [...dest.relationshipsFor(partName)]
	if (sourceRels.length !== destRels.length) return false

	const destById = new Map(destRels.map((rel) => [rel.id, rel]))
	for (const rel of sourceRels) {
		const counterpart = destById.get(rel.id)
		if (!counterpart) return false
		if (counterpart.type !== rel.type || counterpart.targetMode !== rel.targetMode) return false
		if (rel.targetMode === 'External') {
			if (counterpart.target !== rel.target) return false
			continue
		}
		const target = source.relationshipsFor(partName).resolveTarget(rel.id)
		if (dest.relationshipsFor(partName).resolveTarget(rel.id) !== target) return false
		if (copyTraversalStep(sourcePart, rel) !== 'recurse') continue
		if (!identicalSubgraph(dest, source, target, visited)) return false
	}
	return true
}

/**
 * Whether the destination's copy of this part is wired into the deck, for the two kinds
 * where "present in the package" is not the same as "usable".
 *
 * A master absent from `p:sldMasterIdLst` is inert: renderers ignore its background and
 * shape tree, so a slide bound through it renders blank. A layout absent from a
 * registered master's `p:sldLayoutIdLst` is not in the deck's gallery, so `layouts()`
 * cannot name it and `appendSlides` cannot bind to it. Every other kind (theme, media,
 * table styles) is reached by relationship alone and needs no registration.
 */
function wiredIntoDeck(dest: DeckTarget, partName: string): boolean {
	const contentType = dest.opc.part(partName)?.contentType
	if (contentType === SLIDE_MASTER_CONTENT_TYPE) return slideMasterPartNames(dest).includes(partName)
	if (contentType === SLIDE_LAYOUT_CONTENT_TYPE) {
		return slideMasterPartNames(dest).some((master) => layoutPartNamesOf(dest, master).includes(partName))
	}
	return true
}

/** Byte-for-byte equality of two part bodies. */
function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a === b) return true
	if (a.length !== b.length) return false
	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
	return true
}
