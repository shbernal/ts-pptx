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
} from '../../core-interfaces.js'
import { FIXED_PCT_PER_PERCENT } from '../../units.js'
import { convertRotationDegrees, transparencyToAlpha } from '../../units-internal.js'
import { createColorElement } from './color.js'

function boolToXml(value: boolean): string {
	return value ? '1' : '0'
}

function normalizeGradientAngle(angle: number | undefined): number {
	const degrees = angle ?? 0
	if (typeof degrees !== 'number' || !Number.isFinite(degrees))
		throw new Error('Gradient angle must be a finite number.')
	return convertRotationDegrees(((degrees % 360) + 360) % 360)
}

function gradientStopColorAdjustments(stop: GradientStopProps): string {
	let internalElements = ''
	if (stop.transparency) internalElements += `<a:alpha val="${transparencyToAlpha(stop.transparency)}"/>`
	return internalElements
}

function normalizeGradientStops(stops: GradientStopProps[] | undefined): GradientStopProps[] {
	if (!Array.isArray(stops) || stops.length < 2) throw new Error('Gradient fill requires at least two stops.')

	return stops
		.map((stop) => {
			if (!stop || typeof stop.position !== 'number' || !Number.isFinite(stop.position)) {
				throw new Error('Gradient stop position must be a finite number from 0 to 100.')
			}
			if (stop.position < 0 || stop.position > 100) throw new Error('Gradient stop position must be from 0 to 100.')
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
		throw new Error('Gradient fill currently supports only linear and radial gradients.')
	}
	if (typeof gradient.rotateWithShape !== 'undefined' && typeof gradient.rotateWithShape !== 'boolean') {
		throw new Error('Gradient rotateWithShape must be a boolean.')
	}

	const stops = normalizeGradientStops(gradient.stops)
	const rotWithShape = gradient.rotateWithShape ?? true

	let strXml = `<a:gradFill rotWithShape="${boolToXml(rotWithShape)}">`
	strXml += '<a:gsLst>'
	stops.forEach((stop) => {
		const position = Math.round(stop.position * FIXED_PCT_PER_PERCENT)
		strXml += `<a:gs pos="${position}">${createColorElement(stop.color, gradientStopColorAdjustments(stop))}</a:gs>`
	})
	strXml += '</a:gsLst>'
	if (gradient.kind === 'radial') {
		// `<a:path path="circle">` radiates the first stop from a focus rectangle out
		// to the edges. `fillToRect` insets place that focus: equal insets center it,
		// and the `center` percentage shifts it (l/t = center, r/b = 100 - center).
		const cx = Math.max(0, Math.min(100, gradient.center?.x ?? 50))
		const cy = Math.max(0, Math.min(100, gradient.center?.y ?? 50))
		const l = Math.round(cx * FIXED_PCT_PER_PERCENT)
		const t = Math.round(cy * FIXED_PCT_PER_PERCENT)
		const r = Math.round((100 - cx) * FIXED_PCT_PER_PERCENT)
		const b = Math.round((100 - cy) * FIXED_PCT_PER_PERCENT)
		strXml += `<a:path path="circle"><a:fillToRect l="${l}" t="${t}" r="${r}" b="${b}"/></a:path>`
	} else {
		if (typeof gradient.scaled !== 'undefined' && typeof gradient.scaled !== 'boolean')
			throw new Error('Gradient scaled must be a boolean.')
		const scaledAttr = typeof gradient.scaled === 'boolean' ? ` scaled="${boolToXml(gradient.scaled)}"` : ''
		strXml += `<a:lin ang="${normalizeGradientAngle(gradient.angle)}"${scaledAttr}/>`
	}
	strXml += '</a:gradFill>'

	return strXml
}

/**
 * Create a native DrawingML pattern fill.
 * @param {PatternFillProps} pattern pattern fill options
 * @returns XML string
 */
export function genXmlPatternFill(pattern: PatternFillProps | undefined): string {
	if (!pattern) throw new Error('Pattern fill requires a pattern object.')
	const fgColor = pattern.fgColor ?? '000000'
	const bgColor = pattern.bgColor ?? 'FFFFFF'
	return (
		`<a:pattFill prst="${pattern.preset}">` +
		`<a:fgClr>${createColorElement(fgColor)}</a:fgClr>` +
		`<a:bgClr>${createColorElement(bgColor)}</a:bgClr>` +
		'</a:pattFill>'
	)
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
		return '<a:noFill/>'
	}
	const alpha = props.transparency
	const blipInner = alpha ? `<a:alphaModFix amt="${Math.round((100 - alpha) * FIXED_PCT_PER_PERCENT)}"/>` : ''
	return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId${props._imgRid}">${blipInner}</a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
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
			if (props.transparency) internalElements += `<a:alpha val="${transparencyToAlpha(props.transparency)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
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
