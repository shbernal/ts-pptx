/**
 * ts-pptx: DrawingML line (stroke) properties
 *
 * Resolve the pieces of an `<a:ln>` stroke: its width, its `cap` attribute, and
 * its paint child. The paint reuses the shape fill group (`fill.ts`), because
 * DrawingML allows the same fills inside `<a:ln>` as inside a shape.
 */

import type { BorderProps, LineCap, ShapeLineProps } from '../../types/index.js'
import { genXmlColorSelection, genXmlGradientFill } from './fill.js'
import { InvalidOptionError } from '../../errors.js'
import { warnOnce } from '../../diagnostics.js'
import { checkEnumOrWarn } from '../../ooxml/check-enum.js'
import { PRESET_LINE_DASHES } from '../../ooxml/st-enums.js'
import { voidEl } from '../oxml/el.js'

/** Every key `BorderProps` defines. Keep in step with the interface in `types/style.ts`. */
const BORDER_KEYS: readonly string[] = ['type', 'dashType', 'color', 'width', 'transparency', 'cap']

/**
 * Resolve the `a:prstDash/@val` a border should emit.
 *
 * `BorderProps.type` is a three-way switch, so on its own it can only say "dashed" and
 * every dashed border collapses onto `sysDash`. `dashType` names the preset directly and
 * therefore wins when both are given; `type: 'none'` never reaches here, being decided by
 * the caller before any dash is chosen.
 *
 * A value outside `ST_PresetLineDashVal` would make the part schema-invalid, which
 * PowerPoint reports as a corrupt file rather than a mis-set option — so an unrecognized
 * one is reported and dropped back to what `type` implies.
 * @param {BorderProps} border - the caller's border properties
 * @returns {string} a value legal for `a:prstDash/@val`
 */
export function resolveBorderDash(border: BorderProps): string {
	const fromType = border.type === 'dash' ? 'sysDash' : 'solid'
	return (
		checkEnumOrWarn(border.dashType, PRESET_LINE_DASHES, 'border/invalid-dash-type', 'border: dashType') ?? fromType
	)
}

/**
 * Report an authored key that is not part of `BorderProps`, which the library would
 * otherwise drop without a sound.
 *
 * TypeScript's excess-property check already rejects a stray key on a border written
 * *inline* at the call site. It deliberately does not fire when the border is built as
 * a variable first (`const b = {...}; addTable(rows, { border: b })`) — a variable may
 * legitimately be a supertype. That exemption covers precisely the reuse pattern a
 * shared grid style encourages, so the mistake reaches here typed as valid and the
 * wrong-named value is discarded during generation. A `.pptx` gives no other signal:
 * nothing throws, the deck opens, and the border simply renders at its default weight.
 * @param border - the caller's border properties
 */
function warnUnknownBorderKeys(border: BorderProps): void {
	for (const key in border) {
		if (!Object.hasOwn(border, key) || BORDER_KEYS.includes(key)) continue
		warnOnce(
			'border/unknown-key',
			`border: unknown option \`${key}\` is ignored — valid keys are ${BORDER_KEYS.join(', ')}. ` +
				'Line thickness is `width`, in points.',
			{ received: key, valid: BORDER_KEYS }
		)
	}
}

/**
 * Resolve a border's line width in points, falling back to `defaultPt` when `width`
 * is not a usable number.
 *
 * Every `BorderProps` the library reads — table cell borders, table-style regions, and
 * every chart border — resolves its width through here, so this doubles as the one
 * place that can vet a border's shape for all of them ({@link warnUnknownBorderKeys}).
 * @param {BorderProps} border - border properties (may carry `width`)
 * @param {number} defaultPt - width to use when `width` is not a finite number
 * @returns {number} resolved width in points
 */
export function resolveBorderWidth(border: BorderProps, defaultPt: number): number {
	warnUnknownBorderKeys(border)
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
	if (line.type === 'none') return voidEl('a:noFill')
	// `gradient` presence selects a gradient stroke even when `type` was omitted.
	if (line.gradient || line.type === 'gradient') return genXmlGradientFill(line.gradient)
	if (line.type === 'pattern' || line.type === 'image') return genXmlColorSelection(line)
	if (line.color) return genXmlColorSelection(line)
	return ''
}
