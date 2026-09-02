/**
 * Walking a deck's master/layout spine by partname.
 *
 * The `p:sldMasterIdLst` / `p:sldLayoutIdLst` entries hold relationship ids, not partnames, so
 * every traversal has to resolve through the owning part's relationships. These are the
 * source-side reads that do it — used against a *foreign* presentation as often as against this
 * one, which is why they take the deck rather than living on it.
 */

import { attr, firstChild, getElements, type Element } from '../../oxml/dom.js'
import type { DeckTarget } from './deck-target.js'

// Resolving a single-of-a-kind relationship is an OPC concern, not a spine-walking one; it is
// re-exported here because the spine walkers were its first callers.
export { resolveSingleRel } from '../../opc/partnames.js'

/**
 * Master partnames in `p:sldMasterIdLst` order -- the masters renderers treat as active.
 *
 * Takes the structural {@link DeckTarget} rather than a `Presentation`, because it reads only
 * `opc` and `presentationPart` and both callers exist: the import ops walk a source deck,
 * `part-reuse.ts` walks a destination that is not a `Presentation` at all and had its own copy
 * of this for that reason alone.
 */
export function slideMasterPartNames(deck: DeckTarget): string[] {
	const root = deck.presentationPart.dom.documentElement
	const lst = root && firstChild(root, 'p:sldMasterIdLst')
	if (!lst) return []
	const rels = deck.opc.relationshipsFor(deck.presentationPart.partName)
	return idListTargets(lst, 'p:sldMasterId', (relId) => rels.resolveTarget(relId))
}

/** A master's layout partnames in `p:sldLayoutIdLst` order -- its slice of the deck's gallery. */
export function layoutPartNamesOf(deck: DeckTarget, masterPartName: string): string[] {
	const root = deck.opc.part(masterPartName)?.dom.documentElement
	const lst = root && firstChild(root, 'p:sldLayoutIdLst')
	if (!lst) return []
	const rels = deck.opc.relationshipsFor(masterPartName)
	return idListTargets(lst, 'p:sldLayoutId', (relId) => rels.resolveTarget(relId))
}

/** Resolve every `@r:id` in an id list to its partname, skipping entries that carry none. */
function idListTargets(lst: Element, qname: string, resolve: (relId: string) => string): string[] {
	const out: string[] = []
	for (const entry of getElements(lst, qname)) {
		const relId = attr(entry, 'r:id')
		if (relId) out.push(resolve(relId))
	}
	return out
}
