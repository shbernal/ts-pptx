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

const SLIDE_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const SLIDE_LAYOUT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
const NOTES_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml'
const PRESENTATION_MAIN_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'

/**
 * Content types that are shared deck chrome: reachable through the
 * presentation → master → layout → theme graph, not owned by any one slide.
 * {@link Presentation.removeSlide} never prunes these as a removed slide's
 * orphan, even while momentarily unreferenced.
 */
const SHARED_CHROME_CONTENT_TYPES = new Set([
	SLIDE_MASTER_CONTENT_TYPE,
	SLIDE_LAYOUT_CONTENT_TYPE,
	THEME_CONTENT_TYPE,
	'application/vnd.openxmlformats-officedocument.themeOverride+xml',
	NOTES_MASTER_CONTENT_TYPE,
	'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
	PRESENTATION_MAIN_CONTENT_TYPE,
])

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
export function isReferenced(pres: Presentation, partName: string): boolean {
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
