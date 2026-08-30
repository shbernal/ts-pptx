/**
 * ts-pptx: DrawingML line (stroke) properties
 *
 * Resolve the pieces of an `<a:ln>` stroke: its width, its `cap` attribute, and
 * its paint child. The paint reuses the shape fill group (`fill.ts`), because
 * DrawingML allows the same fills inside `<a:ln>` as inside a shape.
 */

import type { BorderProps, LineCap, ShapeLineProps } from '../../types/index.js'
import { fillNamesPaint, genXmlColorSelection, resolveLineKind } from './fill.js'
import { InvalidOptionError } from '../../errors.js'
import { warnOnce } from '../../diagnostics.js'
import { checkEnumOrWarn } from '../../ooxml/check-enum.js'
import { PRESET_LINE_DASHES } from '../../ooxml/st-enums.js'

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
	return typeof val === 'number' && Number.isFinite(val) ? val : defaultPt
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
 *
 * `<a:ln>` takes almost the same fill group as a shape interior, so a stroke can be a gradient
 * or a pattern as well as a solid color. Which one is asked for is {@link resolveLineKind}'s
 * answer, shared with the shape interior — the stroke side used to infer `gradient` from its
 * sub-object while the interior inferred nothing, so the same `{ gradient }` spelling painted a
 * gradient outline and a black fill. "Almost" is the picture fill: `EG_LineFillProperties` has
 * no `a:blipFill`, and `resolveLineKind` is where that is refused.
 *
 * Returns '' when the line names no paint at all ({@link fillNamesPaint}), so the caller emits
 * no fill child and the stroke inherits its color from the theme or placeholder. `type: 'none'`
 * is not that case: it is an explicit *no stroke*, and a shape authored with it grew the theme's
 * border back when the two were conflated.
 * @param {ShapeLineProps} line line options
 * @returns XML string
 */
export function genXmlLineFill(line: ShapeLineProps): string {
	if (!fillNamesPaint(line)) return ''
	// Refuses a picture stroke; every other kind is dispatched by the shared builder. Both
	// `define/` rebuilds have already resolved the kind, so this only fires for a props object
	// reaching the emitter directly.
	resolveLineKind(line)
	return genXmlColorSelection(line)
}
