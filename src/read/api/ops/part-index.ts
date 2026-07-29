/**
 * Walking a deck's master/layout spine by partname.
 *
 * The `p:sldMasterIdLst` / `p:sldLayoutIdLst` entries hold relationship ids, not partnames, so
 * every traversal has to resolve through the owning part's relationships. These are the
 * source-side reads that do it — used against a *foreign* presentation as often as against this
 * one, which is why they take the deck rather than living on it.
 */

import { attr, firstChild, getElements } from '../../oxml/dom.js'
import type { OpcPackage } from '../../opc/package.js'
import type { Presentation } from '../presentation.js'

/** Master partnames in `p:sldMasterIdLst` order. */
export function slideMasterPartNames(pres: Presentation): string[] {
	const root = pres.presentationPart.dom.documentElement
	const lst = root && firstChild(root, 'p:sldMasterIdLst')
	if (!lst) return []
	const rels = pres.opc.relationshipsFor(pres.presentationPart.partName)
	const out: string[] = []
	for (const entry of getElements(lst, 'p:sldMasterId')) {
		const relId = attr(entry, 'r:id')
		if (relId) out.push(rels.resolveTarget(relId))
	}
	return out
}

/** A master's layout partnames in `p:sldLayoutIdLst` order. */
export function layoutPartNamesOf(pres: Presentation, masterPartName: string): string[] {
	const root = pres.opc.part(masterPartName)?.dom.documentElement
	const lst = root && firstChild(root, 'p:sldLayoutIdLst')
	if (!lst) return []
	const rels = pres.opc.relationshipsFor(masterPartName)
	const out: string[] = []
	for (const entry of getElements(lst, 'p:sldLayoutId')) {
		const relId = attr(entry, 'r:id')
		if (relId) out.push(rels.resolveTarget(relId))
	}
	return out
}

/** Resolve the single relationship of `type` owned by `partName`, or `null`. */
export function resolveSingleRel(opc: OpcPackage, partName: string, type: string): string | null {
	const rels = opc.relationshipsFor(partName)
	const rel = rels.byType(type)[0]
	return rel ? rels.resolveTarget(rel.id) : null
}
