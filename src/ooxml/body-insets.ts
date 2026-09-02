/**
 * PowerPoint's default `a:bodyPr` text insets.
 *
 * ECMA-376 §21.1.2.1.1 states them in prose; the XSD leaves `lIns`/`tIns`/`rIns`/`bIns`
 * optional with no schema default, so a reader that wants to know what an omitted side is
 * inset by has to carry the numbers itself. Three places did, in two units and two orders:
 * the script mapper as a points record, the round-trip canonicaliser as an inches tuple in
 * `[top, right, bottom, left]`, and `inspect.ts` as two inch scalars.
 *
 * The canonicaliser's tuple in particular has to be exactly `pointsToInches` of the mapper's
 * record, or `isDefaultMargin` stops matching and every text frame that spells its insets
 * reports four spurious differences. Nothing linked them; this module is the link.
 *
 * It is a fact about the schema, so — like `st-enums.ts` — it belongs to neither `gen/` nor
 * `read/` nor `script/`.
 */

import { POINTS_PER_INCH } from '../units.js'

/** The four default insets in points: 7.2pt (0.1in) left and right, 3.6pt (0.05in) top and bottom. */
export const BODY_INSET_DEFAULTS_PT = { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 } as const

/**
 * The same four in inches, in the write API's `margin` order — `[top, right, bottom, left]`.
 *
 * Derived rather than transcribed, so it cannot drift from {@link BODY_INSET_DEFAULTS_PT}.
 */
export const BODY_INSET_DEFAULTS_IN: readonly number[] = [
	BODY_INSET_DEFAULTS_PT.top / POINTS_PER_INCH,
	BODY_INSET_DEFAULTS_PT.right / POINTS_PER_INCH,
	BODY_INSET_DEFAULTS_PT.bottom / POINTS_PER_INCH,
	BODY_INSET_DEFAULTS_PT.left / POINTS_PER_INCH,
]
