/**
 * ts-pptx: slide-size preconditions for the import surface
 *
 * Every import entry point starts by asking whether the two decks agree on canvas size, and
 * the five that asked wrote out the same comparison, the same `unknown`-printing formatter
 * and the same `InvalidOptionError` each time. That cost more than the duplication: the
 * copies disagreed on what an *unknown* size means, so a consumer branching on `err.code`
 * could not tell "these decks are different sizes" from "I could not read a size" anywhere
 * except `importSlide`, and only when a rescale had been asked for.
 *
 * The two conditions are separated here, and the codes now mean what their names say:
 * - `import/slide-size-unknown` — a deck does not declare `p:sldSz`, so there is no size to
 *   compare and no rescale to compute. Nothing the caller passes can make it comparable.
 * - `import/slide-size-mismatch` — both sizes are known and they differ. This one *is*
 *   answerable, which is why each caller supplies the escape hatch it offers.
 */

import { InvalidOptionError } from '../../../errors.js'
import type { SlideSize } from '../presentation-types.js'

/** A size as the error messages print it. */
function fmt(size: SlideSize | null): string {
	return size ? `${size.widthEmu}×${size.heightEmu} EMU` : 'unknown'
}

/**
 * Both decks' sizes, or a throw when either deck does not declare one.
 *
 * Returned rather than asserted because every caller that needs both sizes known needs the
 * values too — to compute a rescale — and TypeScript cannot narrow two parameters from one
 * assertion signature.
 *
 * @param target - this deck's slide size
 * @param incoming - the source deck's slide size
 * @param api - the method name, opening the message
 */
export function requireKnownSlideSizes(
	target: SlideSize | null,
	incoming: SlideSize | null,
	api: string
): [SlideSize, SlideSize] {
	if (!target || !incoming)
		throw new InvalidOptionError(
			'import/slide-size-unknown',
			`${api} requires both decks to declare a slide size (p:sldSz); target is ${fmt(target)}, source is ${fmt(incoming)}`
		)
	return [target, incoming]
}

/**
 * Assert the two decks share a canvas, throwing whichever condition actually applies.
 *
 * @param target - this deck's slide size
 * @param incoming - the source deck's slide size
 * @param api - the method name, opening the message
 * @param hint - the escape hatch this caller offers, parenthesised into the message; omit
 *   where there is none (`importSlides` and `appendSlides` have no rescale spelling)
 */
export function requireEqualSlideSize(
	target: SlideSize | null,
	incoming: SlideSize | null,
	api: string,
	hint = ''
): void {
	const [known, source] = requireKnownSlideSizes(target, incoming, api)
	if (known.widthEmu === source.widthEmu && known.heightEmu === source.heightEmu) return
	throw new InvalidOptionError(
		'import/slide-size-mismatch',
		`${api} requires equal slide sizes${hint ? ` (${hint})` : ''}; target is ${fmt(target)}, source is ${fmt(incoming)}`
	)
}

/** Whether both sizes are known and equal — the question a caller with a rescale asks first. */
export function slideSizesMatch(target: SlideSize | null, incoming: SlideSize | null): boolean {
	return !!target && !!incoming && target.widthEmu === incoming.widthEmu && target.heightEmu === incoming.heightEmu
}
