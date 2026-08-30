/**
 * ts-pptx: DrawingML fills
 *
 * Emit the fill group a shape, cell, chart element, or line paint can carry:
 * `<a:solidFill>`, `<a:gradFill>`, `<a:pattFill>`, `<a:blipFill>`.
 * {@link genXmlColorSelection} is the dispatch every caller should reach for;
 * the per-kind builders are exported for the few sites that know their fill kind.
 */

import { warn } from '../../diagnostics.js'
import type {
	Color,
	GradientFillProps,
	GradientStopProps,
	PatternFillProps,
	ShapeFillProps,
	ShapeLineProps,
} from '../../types/index.js'
import { FIXED_PCT_PER_PERCENT } from '../../units.js'
import { convertRotationDegrees, transparencyToAlpha } from '../../units-internal.js'
import { createColorElement } from './color.js'
import { InvalidOptionError, UnsupportedFeatureError } from '../../errors.js'
import { el, raw, voidEl } from '../oxml/el.js'

function boolToXml(value: boolean): string {
	return value ? '1' : '0'
}

function normalizeGradientAngle(angle: number | undefined): number {
	const degrees = angle ?? 0
	if (typeof degrees !== 'number' || !Number.isFinite(degrees))
		throw new InvalidOptionError('gradient/angle-non-finite', 'Gradient angle must be a finite number.')
	// Into 0..360 rather than -360..360, which is where `convertRotationDegrees` leaves it:
	// `a:lin/@ang` is ST_PositiveFixedAngle (0..21600000), so a negative gradient angle is a
	// value PowerPoint reports as needing repair, not merely an unusual spelling.
	return convertRotationDegrees(((degrees % 360) + 360) % 360)
}

function alphaFromTransparency(transparency: number | undefined): string {
	return transparency ? voidEl('a:alpha', { val: transparencyToAlpha(transparency) }) : ''
}

function normalizeGradientStops(stops: GradientStopProps[] | undefined): GradientStopProps[] {
	if (!Array.isArray(stops) || stops.length < 2)
		throw new InvalidOptionError('gradient/too-few-stops', 'Gradient fill requires at least two stops.')

	return stops
		.map((stop) => {
			if (!stop || typeof stop.position !== 'number' || !Number.isFinite(stop.position)) {
				throw new InvalidOptionError(
					'gradient/stop-position-non-finite',
					'Gradient stop position must be a finite number from 0 to 100.'
				)
			}
			if (stop.position < 0 || stop.position > 100)
				throw new InvalidOptionError(
					'gradient/stop-position-out-of-range',
					'Gradient stop position must be from 0 to 100.'
				)
			return stop
		})
		.sort((a, b) => a.position - b.position)
}

/**
 * Create a native DrawingML gradient fill.
 * @param {GradientFillProps} gradient gradient fill options
 * @returns XML string
 */
export function genXmlGradientFill(gradient: GradientFillProps | undefined): string {
	if (!gradient || (gradient.kind !== 'linear' && gradient.kind !== 'radial')) {
		throw new UnsupportedFeatureError(
			'gradient/type-unsupported',
			'Gradient fill currently supports only linear and radial gradients.'
		)
	}
	if (typeof gradient.rotateWithShape !== 'undefined' && typeof gradient.rotateWithShape !== 'boolean') {
		throw new InvalidOptionError(
			'gradient/rotate-with-shape-not-boolean',
			'Gradient rotateWithShape must be a boolean.'
		)
	}

	const stops = normalizeGradientStops(gradient.stops)
	const rotWithShape = gradient.rotateWithShape ?? true

	const gsLst = el(
		'a:gsLst',
		null,
		stops.map((stop) =>
			raw(
				el(
					'a:gs',
					{ pos: Math.round(stop.position * FIXED_PCT_PER_PERCENT) },
					raw(createColorElement(stop.color, alphaFromTransparency(stop.transparency)))
				)
			)
		)
	)

	let shape: string
	if (gradient.kind === 'radial') {
		// `<a:path path="circle">` radiates the first stop from a focus rectangle out
		// to the edges. `fillToRect` insets place that focus: equal insets center it,
		// and the `center` percentage shifts it (l/t = center, r/b = 100 - center).
		const cx = Math.max(0, Math.min(100, gradient.center?.x ?? 50))
		const cy = Math.max(0, Math.min(100, gradient.center?.y ?? 50))
		const fillToRect = voidEl('a:fillToRect', {
			l: Math.round(cx * FIXED_PCT_PER_PERCENT),
			t: Math.round(cy * FIXED_PCT_PER_PERCENT),
			r: Math.round((100 - cx) * FIXED_PCT_PER_PERCENT),
			b: Math.round((100 - cy) * FIXED_PCT_PER_PERCENT),
		})
		shape = el('a:path', { path: 'circle' }, raw(fillToRect))
	} else {
		if (typeof gradient.scaled !== 'undefined' && typeof gradient.scaled !== 'boolean')
			throw new InvalidOptionError('gradient/scaled-not-boolean', 'Gradient scaled must be a boolean.')
		// `scaled` is absent rather than defaulted when unset, so it stays `undefined`
		// and `voidEl` drops the attribute entirely.
		shape = voidEl('a:lin', {
			ang: normalizeGradientAngle(gradient.angle),
			scaled: typeof gradient.scaled === 'boolean' ? boolToXml(gradient.scaled) : undefined,
		})
	}

	return el('a:gradFill', { rotWithShape: boolToXml(rotWithShape) }, [raw(gsLst), raw(shape)])
}

/**
 * Create a native DrawingML pattern fill.
 * @param {PatternFillProps} pattern pattern fill options
 * @returns XML string
 */
export function genXmlPatternFill(pattern: PatternFillProps | undefined): string {
	if (!pattern) throw new InvalidOptionError('pattern-fill/missing-pattern', 'Pattern fill requires a pattern object.')
	const fgColor = pattern.fgColor ?? '000000'
	const bgColor = pattern.bgColor ?? 'FFFFFF'
	return el('a:pattFill', { prst: pattern.preset }, [
		raw(el('a:fgClr', null, raw(createColorElement(fgColor)))),
		raw(el('a:bgClr', null, raw(createColorElement(bgColor)))),
	])
}

/**
 * Create a native DrawingML picture (image) fill.
 * The media relationship is registered when the object is added; this only emits
 * the `<a:blipFill>` referencing the pre-resolved rId.
 * @param {ShapeFillProps} props fill props (must carry a resolved `_imgRid`)
 * @returns XML string
 */
export function genXmlImageFill(props: ShapeFillProps | undefined): string {
	if (!props || typeof props._imgRid !== 'number') {
		warn(
			'image-fill/unresolved-media',
			'image fill is missing its resolved media reference; falling back to no fill. Provide `image: { path }` or `image: { data }`.'
		)
		return voidEl('a:noFill')
	}
	const alpha = props.transparency
	const blipInner = alpha ? voidEl('a:alphaModFix', { amt: transparencyToAlpha(alpha) }) : ''
	return el('a:blipFill', { dpi: 0, rotWithShape: 1 }, [
		// `<a:blip>` stays paired even with no `alphaModFix` child — the arity rule.
		raw(el('a:blip', { 'r:embed': `rId${props._imgRid}` }, raw(blipInner))),
		raw(voidEl('a:srcRect')),
		raw(el('a:stretch', null, raw(voidEl('a:fillRect')))),
	])
}

/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
export function genXmlColorSelection(props: Color | ShapeFillProps | ShapeLineProps): string {
	let fillType = 'solid'
	let colorVal = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		if (typeof props === 'string') colorVal = props
		else {
			if (props.type) fillType = props.type
			if (props.color) colorVal = props.color
			internalElements += alphaFromTransparency(props.transparency)
		}

		switch (fillType) {
			// `type: 'none'` is an explicit *no fill*, not the absence of a statement — the
			// same distinction `genXmlLineFill` already makes on the stroke side. Emitting
			// nothing leaves the interior to `p:style/a:fillRef` or the placeholder, so a
			// shape authored transparent came out painted in the theme's accent colour.
			case 'none':
				outText += voidEl('a:noFill')
				break
			// The other half of that distinction, and the reason it needs a name of its own:
			// on the shape/text-box path a missing `fill` already emits `<a:noFill/>`, so
			// omission cannot spell *inherit* there. `'inherit'` emits nothing, which is what
			// hands the interior back to `p:style/a:fillRef` or the placeholder.
			case 'inherit':
				outText += ''
				break
			case 'solid':
				outText += el('a:solidFill', null, raw(createColorElement(colorVal, internalElements)))
				break
			case 'gradient':
				outText += genXmlGradientFill(typeof props === 'string' ? undefined : props.gradient)
				break
			case 'pattern':
				outText += genXmlPatternFill(typeof props === 'string' ? undefined : props.pattern)
				break
			case 'image':
				outText += genXmlImageFill(typeof props === 'string' ? undefined : props)
				break
			default: // @note need a statement as having only "break" can be removed by bundlers, then triggers "no-default" js-linter
				outText += ''
				break
		}
	}

	return outText
}
