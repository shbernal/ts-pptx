/**
 * A background (`p:cSld/p:bg`) → {@link BackgroundIr}, for whichever tier read it.
 *
 * There were two of these and only one of them could carry a picture. The layout/master arm
 * handled `solid`, `image` and `themeRef`; the slide arm handled `solid` and `none` and
 * recorded everything else as *not expressible through the write API*. That claim was not
 * true for `image`: `SlideProps.background` takes the same `BackgroundProps` the layout arm
 * authors, and `BackgroundIr.data` is a declared, documented field that nothing could produce
 * for a slide. So a slide-scoped picture background was dropped with a note that misdescribed
 * why.
 *
 * One mapper, two tiers. What stays per tier is the *note*: `slide.background` and
 * `master.background` are separate constructs in the fidelity catalogue, they carry different
 * wording (a slide falls back to its layout's background; a layout has nothing to fall back
 * to), and an existing declared loss keeps its own key.
 *
 * `transparency` had no producer on either arm and now has one on both: the read model
 * resolves an `a:alpha` on the background colour into {@link ResolvedColor.alpha}, and
 * `BackgroundProps.transparency` is the write option that emits it back.
 */

import type { SlideBackground } from '../../read/api/slide-background.js'
import type { AssetResolver } from './context.js'
import type { BackgroundIr } from '../ir.js'
import type { NoteScope } from '../fidelity.js'
import { literalColor } from './values.js'

/** How one tier names itself in the note it records when a background cannot be carried. */
interface BackgroundTier {
	/** The fidelity construct this tier's losses are recorded under. */
	construct: 'slide.background' | 'master.background'
	/** How a note names the thing losing its background, e.g. `"this layout"`. */
	subject: string
	/** What that thing shows instead, e.g. `"the slide takes its layout's background"`. */
	fallback: string
}

/** The slide tier: a slide's own `p:bg`, which a slide falls back from to its layout's. */
export const SLIDE_BACKGROUND: BackgroundTier = {
	construct: 'slide.background',
	subject: "this slide's background",
	fallback: "the slide takes its layout's background",
}

/** The layout/master tier, which has nothing below it to inherit from. */
export const MASTER_BACKGROUND: BackgroundTier = {
	construct: 'master.background',
	subject: "this layout's background",
	fallback: 'the layout is emitted with no background',
}

/**
 * `BackgroundProps.transparency` for a resolved background colour, or `undefined`.
 *
 * The read model reports **opacity** as a 0-1 fraction and the write option takes
 * **transparency** as a 0-100 percent, so this is the one place the two conventions meet. A
 * fully opaque colour states nothing, which keeps the IR — and therefore the emitted script —
 * free of a key that means "the default".
 */
function transparencyOf(color: { alpha?: number } | null): number | undefined {
	if (!color || color.alpha === undefined) return undefined
	const percent = Math.round((1 - color.alpha) * 100)
	return percent > 0 ? percent : undefined
}

/**
 * Map one tier's background onto the write API's `background` option.
 *
 * `undefined` means *say nothing* — the caller omits the key, and the slide or layout takes
 * whatever it would have inherited. That is the right answer for an absent background and for
 * an explicit `a:noFill`, and it is also what a recorded loss falls back to.
 * @param background - the background the read model resolved, if any
 * @param notes - where a loss this mapper cannot avoid is recorded
 * @param assets - how a picture background's part becomes bytes the script can carry
 * @param tier - which tier is asking, for the note it records
 * @returns the IR background, or `undefined`
 */
export function backgroundIr(
	background: SlideBackground | null,
	notes: NoteScope,
	assets: AssetResolver,
	tier: BackgroundTier
): BackgroundIr | undefined {
	if (!background || background.type === 'none') return undefined

	switch (background.type) {
		case 'solid': {
			if (!background.color) return undefined
			const transparency = transparencyOf(background.color)
			return transparency === undefined
				? { color: literalColor(background.color.effectiveHex) }
				: { color: literalColor(background.color.effectiveHex), transparency }
		}
		case 'image': {
			const asset = background.partName === null ? null : assets.assetFor(background.partName)
			return asset ? { data: asset } : undefined
		}
		case 'themeRef': {
			// `p:bgRef` indexes the theme's background fill list, which the write path cannot
			// author. The read model resolves it, so the flat colour survives even though the
			// reference — and therefore its response to a theme change — does not.
			const fill = background.resolvedFill
			if (fill?.type === 'solid' && fill.color) {
				notes.note(
					tier.construct,
					'flattened',
					'unwritable',
					`${tier.subject} is a theme reference (p:bgRef into the theme's background fill list), which has no write-API option; the colour it currently resolves to is baked in and stops following the theme`
				)
				const transparency = transparencyOf(fill.color)
				return transparency === undefined
					? { color: literalColor(fill.color.effectiveHex) }
					: { color: literalColor(fill.color.effectiveHex), transparency }
			}
			notes.note(
				tier.construct,
				'dropped',
				'unwritable',
				`${tier.subject} is a theme reference to a non-solid fill, which has no write-API option; ${tier.fallback}`
			)
			return undefined
		}
		default:
			notes.note(
				tier.construct,
				'dropped',
				'unsupported',
				`a ${background.type} background is not expressible through the write API's background option, so ${tier.fallback}`
			)
			return undefined
	}
}
