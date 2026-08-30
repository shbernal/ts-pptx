/**
 * ts-pptx: EMU at the read model's boundary — validated on the way in, converted on the way out.
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
import { EMU_PER_POINT } from '../../units.js'

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
	return emu === null ? null : emu / EMU_PER_POINT
}
