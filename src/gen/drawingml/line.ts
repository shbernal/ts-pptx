/**
 * ts-pptx: DrawingML line (stroke) properties
 *
 * Resolve the pieces of an `<a:ln>` stroke: its width, its `cap` attribute, and
 * its paint child. The paint reuses the shape fill group (`fill.ts`), because
 * DrawingML allows the same fills inside `<a:ln>` as inside a shape.
 */

import type { BorderProps, LineCap, ShapeLineProps } from '../../core-interfaces.js'
import { genXmlColorSelection, genXmlGradientFill } from './fill.js'
import { InvalidOptionError } from '../../errors.js'

/**
 * Resolve a border's line width in points, falling back to `defaultPt` when `width`
 * is not a usable number.
 * @param {BorderProps} border - border properties (may carry `width`)
 * @param {number} defaultPt - width to use when `width` is not a finite number
 * @returns {number} resolved width in points
 */
export function resolveBorderWidth(border: BorderProps, defaultPt: number): number {
	const val = border.width
	return typeof val === 'number' && !isNaN(val) ? val : defaultPt
}

/**
 * Map a friendly `LineCap` value to the OOXML `cap` attribute value (`flat`/`sq`/`rnd`).
 * @param {LineCap} [lineCap] - line cap style (defaults to `flat`)
 * @returns {string} value for the `cap` attribute on `<a:ln>`
 */
export function createLineCap(lineCap?: LineCap): string {
	if (!lineCap || lineCap === 'flat') {
		return 'flat'
	} else if (lineCap === 'square') {
		return 'sq'
	} else if (lineCap === 'round') {
		return 'rnd'
	} else {
		const neverLineCap: never = lineCap
		throw new InvalidOptionError('line/invalid-cap', `Invalid line cap: ${String(neverLineCap)}`)
	}
}

/**
 * Emit the paint child of an `<a:ln>` stroke.
 * DrawingML allows the same fill group inside `<a:ln>` as inside a shape fill, so a
 * stroke can be a gradient/pattern as well as a solid color:
 * - a `gradient` (or `type: 'gradient'`) produces a `<a:gradFill>` (gradient stroke);
 * - a `pattern`/`image` type delegates to the shared fill dispatch;
 * - otherwise a `color` produces a `<a:solidFill>`.
 * Returns '' when the line specifies no paint, so the caller emits no fill child and
 * the stroke inherits its color from the theme/placeholder.
 * @param {ShapeLineProps} line line options
 * @returns XML string
 */
export function genXmlLineFill(line: ShapeLineProps): string {
	// `type: 'none'` is an explicit *no stroke*, and it is not the same as saying nothing.
	// Omitting the paint child leaves the outline to the theme or placeholder, so a shape
	// authored with `line: { type: 'none' }` grew the theme's border instead of losing it.
	if (line.type === 'none') return '<a:noFill/>'
	// `gradient` presence selects a gradient stroke even when `type` was omitted.
	if (line.gradient || line.type === 'gradient') return genXmlGradientFill(line.gradient)
	if (line.type === 'pattern' || line.type === 'image') return genXmlColorSelection(line)
	if (line.color) return genXmlColorSelection(line)
	return ''
}
