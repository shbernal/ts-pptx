/**
 * ts-pptx: an authored object's frame
 *
 * Every definer used to spell its own default for an omitted `x`/`y`/`w`/`h`, and they disagreed
 * on what "omitted" meant. A shape tested `x || (x === 0 ? 0 : 1)`, media `|| 2`, an image `|| 1`
 * after its natural size, a chart `|| '50%'`, OLE objects and 3D models `?? 4` and `?? 3`, zooms
 * `?? 0`. So `w: 0, h: 0` gave media a 2in square, a chart half the slide and an image its natural
 * size, while a shape, an OLE object or a zoom got nothing, and none of them said so. A `NaN` took
 * the default wherever `||` was used and reached the converter wherever `??` was.
 */

import { warn } from '../../diagnostics.js'
import type { Coord } from '../../types/index.js'

type Axis = 'x' | 'y' | 'w' | 'h'
const AXES: readonly Axis[] = ['x', 'y', 'w', 'h']

/** A frame as the caller wrote it: any axis may be missing or `undefined`. */
export type AuthoredFrame = { [A in Axis]?: Coord | undefined }

/** `0`, as a number or as a unit string such as `'0in'` or `'0%'`. */
function isZeroCoord(value: Coord | undefined): boolean {
	if (typeof value === 'number') return value === 0
	return typeof value === 'string' && /^\s*-?(?:0+(?:\.0*)?|\.0+)\s*(?:%|in|pt|px|emu)\s*$/.test(value)
}

/**
 * An object's frame: each axis the caller stated, else the definer's default.
 *
 * Only `undefined` and `null` are unstated. `0` is a stated extent and is kept, and anything else,
 * `NaN` included, is kept for the coordinate converter to accept or refuse. Each definer passes
 * only its own defaults. An axis with neither a stated value nor a default stays absent, for the
 * render pass to default.
 *
 * A zero `w` or `h` draws nothing along that axis, so it warns `frame/zero-extent`. A line is the
 * exception: it is drawn with no height or no width on purpose, and its definer says so.
 * @param given - the caller's frame
 * @param defaults - the definer's default for each axis it defaults
 * @param api - the method the caller called, opening the warning
 * @param zeroExtentAllowed - true for a line
 * @returns the resolved frame
 */
export function resolveAuthoredFrame<K extends Axis>(
	given: AuthoredFrame,
	defaults: Record<K, Coord>,
	api: string,
	zeroExtentAllowed = false
): Record<K, Coord> & AuthoredFrame {
	const fallback: AuthoredFrame = defaults
	const frame: AuthoredFrame = {}
	for (const axis of AXES) {
		const value = given[axis] ?? fallback[axis]
		if (value !== undefined) frame[axis] = value
	}
	const zero = (['w', 'h'] as const).filter((axis) => isZeroCoord(frame[axis]))
	if (zero.length > 0 && !zeroExtentAllowed) {
		const stated = zero.map((axis) => `${axis} is ${String(frame[axis])}`).join(' and ')
		const missing = zero.map((axis) => (axis === 'w' ? 'width' : 'height')).join(' and no ')
		warn('frame/zero-extent', `${api}: ${stated}, so the object has no ${missing} and may not be visible.`)
	}
	return frame as Record<K, Coord> & AuthoredFrame
}
