/**
 * Orphan pruning after a part stops being referenced.
 *
 * Removing a slide can strand the parts only it used — its notes slide, media, chart parts. This
 * walks that fringe: a part is dropped when nothing left in the package (or the package root)
 * still resolves an internal relationship to it, and the parts *it* referenced are then examined
 * the same way. Shared chrome (masters, layouts, themes, the notes master) is exempt, so a deck
 * that momentarily has no slides does not lose its design.
 */

import { relsPartNameFor } from '../../opc/partnames.js'
import type { Presentation } from '../presentation.js'
import { SHARED_PARTS } from '../../../ooxml/rel-types.js'

/**
 * Content types that are shared deck chrome, not owned by any one slide: the `deck` rows of
 * {@link SHARED_PARTS}, reached through the presentation → master → layout → theme graph or named
 * by every comment part. {@link Presentation.removeSlide} never prunes these as a removed slide's
 * orphan, even while momentarily unreferenced.
 */
const SHARED_CHROME_CONTENT_TYPES: ReadonlySet<string> = new Set(
	SHARED_PARTS.flatMap((kind) => (kind.scope === 'deck' && kind.contentType !== null ? [kind.contentType] : []))
)

/**
 * Remove `partName` if it is neither shared chrome nor still referenced by any
 * remaining part, then recurse into the parts it referenced. The pruning a
 * removed slide triggers (notes/media/charts the slide alone used).
 */
export function pruneIfOrphan(pres: Presentation, partName: string): void {
	const part = pres.opc.part(partName)
	if (!part || SHARED_CHROME_CONTENT_TYPES.has(part.contentType)) return
	if (isReferenced(pres, partName)) return
	const rels = pres.opc.relationshipsFor(partName)
	const childTargets = [...rels].filter((rel) => rel.targetMode !== 'External').map((rel) => rels.resolveTarget(rel.id))
	pres.opc.removePart(relsPartNameFor(partName))
	pres.opc.removePart(partName)
	for (const child of childTargets) pruneIfOrphan(pres, child)
}

/** Whether any remaining part (or the package root) resolves an internal relationship to `partName`. */
function isReferenced(pres: Presentation, partName: string): boolean {
	for (const owner of [...pres.opc.parts.keys(), '/']) {
		if (owner.endsWith('.rels')) continue
		const rels = pres.opc.relationshipsFor(owner)
		for (const rel of rels) {
			if (rel.targetMode === 'External') continue
			if (rels.resolveTarget(rel.id) === partName) return true
		}
	}
	return false
}
