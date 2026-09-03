/**
 * ts-pptx: units at the read model's boundary — validated on the way in, converted on the way out.
 *
 * Every read-side setter that takes a measurement has to answer the same two questions before
 * the value reaches an attribute: is it finite, and (for an extent) is it positive. `NaN` or
 * `Infinity` written verbatim makes the part invalid, and silently coercing it to something
 * drawable is exactly the footgun the project's API policy rules out — so both are a throw.
 *
 * The two halves used to be split across `table-edit.ts` (threw, then rounded) and `slide.ts`
 * (threw only, with the rounding done eight lines later at each build site), under two spellings
 * of one message. Rounding *here* is what makes the split unnecessary: a caller that has checked
 * the value already holds the integer it is going to write.
 */

import { InvalidOptionError } from '../../errors.js'
import type { InvalidOptionErrorCode } from '../../codes.js'
import { emuToPoints, FIXED_PCT_PER_PERCENT, HUNDREDTHS_PER_POINT } from '../../units.js'

/**
 * Round a measurement to whole EMU, rejecting a value that cannot be written.
 * @param value - the caller's measurement, in EMU
 * @param field - the option name, opening the message
 * @param code - the error code to raise; defaults to the generic coordinate one
 */
export function checkFiniteEmu(
	value: number,
	field: string,
	code: InvalidOptionErrorCode = 'coord/non-finite'
): number {
	if (!Number.isFinite(value))
		throw new InvalidOptionError(code, `${field} must be a finite number of EMU, got: ${String(value)}`)
	return Math.round(value)
}

/**
 * As {@link checkFiniteEmu}, and additionally rejects a non-positive value — the guard an
 * *extent* needs, since `ST_PositiveCoordinate` has no room for zero or below and a zero-size
 * shape is a degenerate result rather than a small one.
 */
export function checkPositiveEmu(value: number, field: string): number {
	const rounded = checkFiniteEmu(value, field)
	if (value <= 0) throw new InvalidOptionError('coord/not-positive', `${field} must be positive, got: ${String(value)}`)
	return rounded
}

/**
 * EMU → points, propagating "absent". The read model reports line widths, effect radii and text
 * insets in points while the DOM stores EMU, so `x === null ? null : x / EMU_PER_POINT` stood at
 * five getters across `chart.ts`, `shapes/base.ts`, `table.ts` and `text.ts`.
 *
 * The accumulator sites that read `if (v !== null) out.k = v / EMU_PER_POINT` keep their own
 * expression: their target is `number | undefined`, so propagating a `null` through here would
 * only have to be unwrapped again.
 */
export function ptFromEmu(emu: number | null): number | null {
	return emu === null ? null : emuToPoints(emu)
}

/**
 * Hundredths of a point → points, propagating "absent". DrawingML measures font size, character
 * spacing, point line spacing and bullet size in hundredths (`a:rPr/@sz`, `@spc`,
 * `a:spcPts/@val`, `a:buSzPts/@val`), and six getters across `text.ts` and `theme-context.ts`
 * each wrote the divisor out as a bare `100`.
 *
 * Naming the unit is the point: `a:buSzPct` next door is *thousandths of a percent*, so a bare
 * divisor leaves a reader to remember which attribute is which. Read against
 * {@link pctFromThousandths}.
 *
 * The two `LineSpacing` sites divide by {@link HUNDREDTHS_PER_POINT} and
 * {@link FIXED_PCT_PER_PERCENT} in place, for the reason {@link ptFromEmu} gives about
 * accumulators: their field is a plain `number` inside an object the caller's own null check
 * already guards, so a nullable return would only be unwrapped again.
 */
export function ptFromHundredths(value: number | null): number | null {
	return value === null ? null : value / HUNDREDTHS_PER_POINT
}

/**
 * Thousandths of a percent → percent, propagating "absent" — `a:buSzPct/@val`
 * (`ST_TextBulletSizePercent`), which spells 100% as `100000`. That is the same fixed-point
 * percentage the write side scales into with {@link FIXED_PCT_PER_PERCENT}, read in the other
 * direction.
 *
 * **Fixed-point only, by design.** It takes a `number`, so a caller must have parsed the
 * attribute already — and the `…PercentOrPercentString` types (`ST_Percentage` and its
 * relatives) also admit a decimal string with a literal `%`, which is the ONLY form the Strict
 * profile has. `numberValue('62.5%')` is `null`, so four getters reported a value that was
 * present as absent. Read one of those through `parsePercent`/`pctAttr` in `read/oxml/dom.ts`,
 * which takes the raw attribute and knows both spellings; this one is for the attributes whose
 * type has no string form.
 */
export function pctFromThousandths(value: number | null): number | null {
	return value === null ? null : value / FIXED_PCT_PER_PERCENT
}
