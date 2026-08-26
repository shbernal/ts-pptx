/**
 * The structural destination-deck context shared by the modules that mutate a
 * package on `Presentation`'s behalf.
 *
 * Those modules take this rather than the class so they stay independent of their
 * caller; `Presentation` satisfies it structurally, so a call site passes `this`
 * and threads nothing. Pass the instance rather than spreading it into an object
 * literal -- `presentationPart` is a memoizing getter, and a spread resolves it
 * eagerly at the call site instead of on first use.
 */

import type { OpcPackage } from '../../opc/package.js'
import type { Part } from '../../opc/part.js'
import type { Relationships } from '../../opc/relationships.js'

/** The destination deck: just the package and its main part. `Presentation` satisfies it. */
export interface DeckTarget {
	readonly opc: OpcPackage
	readonly presentationPart: Part
}

/**
 * `ppt/_rels/presentation.xml.rels` -- the deck's own relationship part.
 *
 * Every part that masters, layouts, the notes master and embedded fonts hang off
 * is reached through here, so the two-step `dest.presentationPart` ->
 * `dest.opc.relationshipsFor(part.partName)` was written out at every one of
 * those wiring sites. Naming it once means a call site says which rels it wants
 * rather than how to find them.
 * @param {DeckTarget} dest - the destination deck
 * @returns {Relationships} the presentation part's relationships
 */
export function presentationRels(dest: DeckTarget): Relationships {
	return dest.opc.relationshipsFor(dest.presentationPart.partName)
}
