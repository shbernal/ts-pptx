/**
 * ts-pptx: DrawingML line (stroke) properties
 *
 * Resolve the pieces of an `<a:ln>` stroke: its width, its `cap` attribute, and
 * its paint child. The paint reuses the shape fill group (`fill.ts`), because
 * DrawingML allows the same fills inside `<a:ln>` as inside a shape.
 *
 * It also assembles the whole element for the one stroke shape that recurs: the
 * `CT_LineProperties` a `BorderProps` paints, on a chart data point, a chart or plot
 * area, and each of a table cell's six edges ({@link borderLine}).
 */

import type { BorderProps, Color, LineCap, ShapeLineProps, StrokeProps } from '../../types/index.js'
import { fillNamesPaint, genXmlColorSelection, resolveLineKind, solidPaint } from './fill.js'
import { InvalidOptionError } from '../../errors.js'
import { warnOnce } from '../../diagnostics.js'
import { checkEnumOrWarn } from '../../ooxml/check-enum.js'
import { LINE_END_TYPES, PRESET_LINE_DASHES } from '../../ooxml/st-enums.js'
import { lineWidthToEmu } from '../../units-internal.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'

/**
 * Every key `BorderProps` defines.
 *
 * Derived from the interface rather than transcribed from it: the `satisfies` makes a key added
 * to `BorderProps` and not to this list a compile error, where the old `readonly string[]` and a
 * "keep in step" comment made it a silent omission.
 */
const BORDER_KEYS: readonly string[] = Object.keys({
	type: 0,
	dashType: 0,
	color: 0,
	width: 0,
	transparency: 0,
	cap: 0,
} satisfies Record<keyof Required<BorderProps>, 0>)

/**
 * Resolve an `a:prstDash/@val`, falling back when the caller named nothing or named something
 * outside `ST_PresetLineDashVal`.
 *
 * Nine sites in the emitters reach that attribute and five of them wrote the caller's value
 * straight through, so the SAME type on the SAME attribute got two answers: a table border's
 * `dashType: 'bogusDash'` warned and fell back to solid, a shape line's went into the part. A
 * value outside the union makes the part schema-invalid, which PowerPoint reports as a corrupt
 * file rather than as a mis-set option.
 * @param value - the caller's dash, if any
 * @param fallback - what to emit when the caller named none, or named one outside the union
 * @param label - how the diagnostic names the option, e.g. `'border: dashType'`
 * @returns a value legal for `a:prstDash/@val`
 */
export function resolveDash(value: string | undefined | null, fallback: string, label: string): string {
	return checkEnumOrWarn(value, PRESET_LINE_DASHES, 'border/invalid-dash-type', label) ?? fallback
}

/**
 * Resolve the `a:prstDash/@val` a {@link StrokeProps} should emit.
 *
 * `type` is a three-way switch, so on its own it can only say "dashed" and every dashed
 * stroke collapses onto `sysDash`. `dashType` names the preset directly and therefore wins
 * when both are given; `type: 'none'` never reaches here, being decided by the caller before
 * any dash is chosen.
 * @param stroke - the caller's stroke properties
 * @param label - how the diagnostic names the option, e.g. `'border: dashType'`
 * @returns a value legal for `a:prstDash/@val`
 */
export function strokeDash(stroke: StrokeProps, label: string): string {
	return resolveDash(stroke.dashType, stroke.type === 'dash' ? 'sysDash' : 'solid', label)
}

/**
 * {@link strokeDash} under the name the table and chart-area borders call it by.
 * @param {BorderProps} border - the caller's border properties
 * @returns {string} a value legal for `a:prstDash/@val`
 */
export function resolveBorderDash(border: BorderProps): string {
	return strokeDash(border, 'border: dashType')
}

/**
 * The `<a:solidFill>` (or `<a:noFill/>`) a {@link StrokeProps} paints with.
 *
 * `type: 'none'` is the one state that is not a colour, and it is decided here rather than at
 * each of the four chart call sites that used to spell it as a separate `*Show: false` flag or
 * as a `style: 'none'` sentinel.
 * @param stroke - the caller's stroke properties
 * @param defaultColor - the colour to paint when the caller named none
 * @returns the paint element
 */
export function strokePaint(stroke: StrokeProps, defaultColor: Color): string {
	if (stroke.type === 'none') return voidEl('a:noFill')
	return genXmlColorSelection(solidPaint(stroke.color || defaultColor, stroke.transparency))
}

/**
 * Resolve an arrowhead type for `a:headEnd/@type` / `a:tailEnd/@type`, or `null` when the caller
 * named none -- the attribute is then not written and that end of the stroke stays plain.
 *
 * `ST_LineEndType` had no runtime tuple at all, so the two options were emitted verbatim and an
 * untyped caller could reach the attribute with anything.
 * @param value - the caller's arrow type, if any
 * @param label - how the diagnostic names the option, e.g. `'line: beginArrowType'`
 * @returns a value legal for the attribute, or `null` to omit it
 */
export function resolveLineEnd(value: string | undefined | null, label: string): string | null {
	return checkEnumOrWarn(value, LINE_END_TYPES, 'line/invalid-arrow-type', label)
}

/**
 * One arrowhead element (`<a:headEnd>` / `<a:tailEnd>`), or `''` when the caller named no arrow
 * or named one outside `ST_LineEndType`.
 * @param tag - `a:headEnd` or `a:tailEnd`
 * @param value - the caller's arrow type, if any
 * @param label - how the diagnostic names the option
 * @returns the element, or `''`
 */
export function lineEndEl(tag: 'a:headEnd' | 'a:tailEnd', value: string | undefined | null, label: string): string {
	const type = resolveLineEnd(value, label)
	return type === null ? '' : voidEl(tag, { type })
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

/**
 * The `<a:ln>`-shaped stroke a {@link BorderProps} paints, in the one place every border
 * emitter now derives it.
 *
 * Four builders used to turn a `BorderProps` into a line element, each with its own
 * defaults, and the differences between them were only visible by reading all four: a
 * chart data point's outline, a chart-area or plot-area border, the `<c:spPr>` line the
 * plot and chart space share, and a table cell's six edges. They agree on the shape —
 * `w` and `cap` on the element, a solid paint from `color`/`transparency`, then an
 * optional dash and whatever the schema allows after it — and disagree only on what fills
 * those slots. Those are this function's parameters, so each caller is now a named
 * argument set rather than a re-derivation.
 *
 * Two things are deliberately *not* parameters. A stroke that paints nothing is spelled
 * differently by each caller (a bare `<a:ln><a:noFill/></a:ln>` for a chart, an
 * attribute-carrying `w="0"` rule for a table cell, and nothing at all for a data point,
 * whose caller decides whether to emit it) — see {@link noStrokeLine} for the shared
 * spelling of the first. And the colour default is taken on `||`, not `??`, so an empty
 * `color` reads as *unstated* rather than reaching the colour validator as a bad value:
 * three of the four callers used to hand `''` straight through, which warns
 * `color/invalid-value` and paints `DEF_FONT_COLOR` — black — instead of the border
 * default they had named a line earlier. The chart-border builder already read it this way,
 * and `dataBorder` is the only one of the other three a caller can actually reach with an
 * empty string; the rest resolve their colour in the definition step first.
 *
 * @param name - the element's qualified name (`a:ln`, or `a:lnL`/`a:lnTlToBr`/… for a table cell)
 * @param border - the caller's border properties
 * @param spec - `defaultWidth`/`defaultColor` are used when the border states neither;
 *   `cap` is the `cap` attribute value, already resolved; `extraAttrs` follows it in
 *   document order; `dash` is the `a:prstDash/@val` to write, or `null` to omit the
 *   element; `tail` is the already-built children that follow the dash (`a:round`, and a
 *   table cell's head/tail ends), in schema order.
 * @returns the border element XML
 */
export function borderLine(
	name: string,
	border: BorderProps,
	spec: {
		defaultWidth: number
		defaultColor: string
		cap: string
		extraAttrs?: XmlAttrs
		dash: string | null
		tail: readonly string[]
	}
): string {
	return el(
		name,
		{ w: lineWidthToEmu(resolveBorderWidth(border, spec.defaultWidth)), cap: spec.cap, ...spec.extraAttrs },
		[
			genXmlColorSelection(solidPaint(border.color || spec.defaultColor, border.transparency)),
			...(spec.dash === null ? [] : [voidEl('a:prstDash', { val: spec.dash })]),
			...spec.tail,
		].map(raw)
	)
}

/**
 * The `<a:ln>` that paints no stroke at all: an explicit `<a:noFill/>` on a bare element.
 *
 * This is the chart side's spelling of "no line" and is a different statement from a
 * {@link borderLine} whose paint is absent — an omitted `a:ln` inherits the theme's
 * `a:lnRef`, where this overrides it. A table cell spells the same intent with the
 * attributes still on the element and `w="0"`, which is why that one is not this.
 * @returns the `<a:ln><a:noFill/></a:ln>` element XML
 */
export function noStrokeLine(): string {
	return el('a:ln', null, raw(voidEl('a:noFill')))
}
