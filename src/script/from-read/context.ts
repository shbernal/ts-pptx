/**
 * ts-pptx: the two things every read-side mapper needs, in one parameter.
 *
 * A mapper records what it could not carry (`notes`) and registers the bytes it references
 * (`assets`), and the pair was threaded positionally through about twenty signatures — several
 * of which had already grown a third and fourth parameter around them. A pair that always
 * travels together and is never split is one argument.
 *
 * `AssetResolver` lives here rather than in `shape.ts` for the same reason: it is a contract
 * between the deck-level walk and every mapper, not part of the shape mapper, and putting it
 * beside the context keeps `context.ts` a leaf that the mappers can all import.
 */

import type { AssetRef } from '../ir.js'
import type { NoteScope } from '../fidelity.js'

/** Resolves an image/media part name to bytes the deck-level walk has registered. */
export interface AssetResolver {
	/** Register a part's bytes and hand back the reference that stands in for them. */
	assetFor(partName: string): AssetRef | null
	/**
	 * The part's content type, or `null` when it is not in the package.
	 *
	 * Separate from {@link assetFor} because it has to be answerable *without* registering
	 * anything: a caller that rejects a part on its type (a picture fill's SVG blip, which
	 * the write path refuses) would otherwise leave bytes in the asset list that no call
	 * references, and those bytes are emitted as a file next to the script.
	 */
	contentTypeOf(partName: string): string | null
}

/** What every `from-read/` mapper is handed alongside the thing it is mapping. */
export interface MapContext {
	/** Where a loss this mapper cannot avoid is recorded. */
	readonly notes: NoteScope
	/** How a referenced part's bytes become an asset the emitted script can carry. */
	readonly assets: AssetResolver
}

/**
 * The same context, with its notes bound to one shape.
 *
 * A note recorded through the result names that shape, which is what lets a reader of the
 * emitted script find the shape a loss belongs to. Callers used to re-derive this inline as
 * `notes.forShape(shape.name || null)`, sometimes twice in one function and once in a callee
 * that already had a scoped one.
 * @param ctx - the enclosing context
 * @param shape - the shape to scope to; an unnamed one scopes to `null`
 */
export function forShape(ctx: MapContext, shape: { name: string | null }): MapContext {
	return { notes: ctx.notes.forShape(shape.name || null), assets: ctx.assets }
}
