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
 * The same misclaim outlived that fix for a gradient and a pattern. `BackgroundProps` extends
 * `ShapeFillProps`, and the background emitter writes any fill a shape's would, so both map
 * through the shape fill's own gradient and pattern mappers.
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

import type { BackgroundFill, SlideBackground } from '../../read/api/slide-background.js'
import type { ResolvedColor } from '../../read/api/theme-context.js'
import type { AssetResolver } from './context.js'
import type { BackgroundIr } from '../ir.js'
import type { NoteScope } from '../fidelity.js'
import { gradientStops, patternOption, type BackgroundSurface } from './surface-fill.js'
import { alphaToTransparency, literalColor } from './values.js'

/** How one tier names itself in the note it records when a background cannot be carried. */
interface BackgroundTier {
	/** The fidelity construct this tier's losses are recorded under. */
	construct: BackgroundSurface
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
	if (!background) return undefined
	if (background.type !== 'themeRef') return fillIr(background, notes, assets, tier)

	// `p:bgRef` indexes the theme's background fill list, which the write path cannot author.
	// The read model resolves it, so the fill survives even though the reference — and
	// therefore its response to a theme change — does not.
	const fill = background.resolvedFill
	const ir = fill ? fillIr(fill, notes, assets, tier) : undefined
	if (fill && ir) {
		notes.note(
			tier.construct,
			'flattened',
			'unwritable',
			`${tier.subject} is a theme reference (p:bgRef into the theme's background fill list), which has no write-API option; the ${fill.type === 'solid' ? 'colour' : fill.type} it currently resolves to is baked in and stops following the theme`
		)
		return ir
	}
	notes.note(
		tier.construct,
		'dropped',
		'unwritable',
		`${tier.subject} is a theme reference whose fill does not resolve to one the write API can express, and a reference itself has no write-API option; ${tier.fallback}`
	)
	return undefined
}

/** An explicit background fill, or `undefined` for none or for one with nothing to carry. */
function fillIr(
	fill: BackgroundFill,
	notes: NoteScope,
	assets: AssetResolver,
	tier: BackgroundTier
): BackgroundIr | undefined {
	switch (fill.type) {
		case 'none':
			return undefined
		case 'solid':
			return fill.colorRef.resolved ? solidIr(fill.colorRef.resolved) : undefined
		case 'image': {
			const asset = fill.partName === null ? null : assets.assetFor(fill.partName)
			return asset ? { data: asset } : undefined
		}
		case 'gradient': {
			const gradient = gradientStops(fill.gradient, notes, tier.construct)
			return gradient ? { type: 'gradient', gradient } : undefined
		}
		case 'pattern': {
			const pattern = patternOption(fill, notes, tier.construct)
			return pattern ? (pattern as BackgroundIr) : undefined
		}
	}
}

/** A solid colour, with the transparency its `a:alpha` states. */
function solidIr(color: ResolvedColor): BackgroundIr {
	const transparency = alphaToTransparency(color.alpha)
	return transparency === undefined
		? { color: literalColor(color.effectiveHex) }
		: { color: literalColor(color.effectiveHex), transparency }
}
