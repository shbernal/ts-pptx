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

/** The destination deck: just the package and its main part. `Presentation` satisfies it. */
export interface DeckTarget {
	readonly opc: OpcPackage
	readonly presentationPart: Part
}
