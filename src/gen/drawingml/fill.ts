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
import type { ShapeFillPropsInternal } from '../../types/internal.js'
import { FIXED_PCT_PER_PERCENT } from '../../units.js'
import { clampRangedInput, convertRotationDegrees, transparencyToAlpha } from '../../units-internal.js'
import { alphaEl, createColorElement } from './color.js'
import { genXmlImageCropRect, STRETCH_FILL_RECT } from './src-rect.js'
import { InvalidOptionError, UnsupportedFeatureError } from '../../errors.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { xsdBool } from '../../ooxml/xsd-boolean.js'

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
	return transparency ? alphaEl(transparencyToAlpha(transparency)) : ''
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
		// `clampRangedInput`, not a bare `Math.min`/`Math.max`: those propagate `NaN` straight
		// through to `l="NaN"`, and every other percentage option on this object already throws
		// on a non-number and warns on a clamp. The stops a few lines above always did.
		const centerPct = (value: number | undefined, axis: 'x' | 'y'): number =>
			clampRangedInput(value ?? 50, 0, 100, 'gradient/center-out-of-range', `gradient.center.${axis}`)
		const cx = centerPct(gradient.center?.x, 'x')
		const cy = centerPct(gradient.center?.y, 'y')
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
			scaled: typeof gradient.scaled === 'boolean' ? xsdBool(gradient.scaled) : undefined,
		})
	}

	return el('a:gradFill', { rotWithShape: xsdBool(rotWithShape) }, [raw(gsLst), raw(shape)])
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
 * @param props fill props (must carry a resolved `_imgRid`)
 * @returns XML string
 */
export function genXmlImageFill(props: ShapeFillPropsInternal | undefined): string {
	if (!props || typeof props._imgRid !== 'number') {
		warn(
			'image-fill/unresolved-media',
			'image fill is missing its resolved media reference; falling back to no fill. Provide `image: { path }` or `image: { data }`.'
		)
		return voidEl('a:noFill')
	}
	const alpha = props.transparency
	const blipInner = alpha ? voidEl('a:alphaModFix', { amt: transparencyToAlpha(alpha) }) : ''
	// An empty `<a:srcRect/>` is the uncropped form, and stays byte-identical to what every
	// fill emitted before `crop` existed.
	const srcRect = props.image?.crop ? genXmlImageCropRect(props.image.crop, 'image fill', '') : voidEl('a:srcRect')
	return el('a:blipFill', { dpi: 0, rotWithShape: 1 }, [
		// `<a:blip>` stays paired even with no `alphaModFix` child — the arity rule.
		raw(el('a:blip', { 'r:embed': `rId${props._imgRid}` }, raw(blipInner))),
		raw(srcRect),
		raw(STRETCH_FILL_RECT),
	])
}

/** The fill kinds a `ShapeFillProps`/`ShapeLineProps` can ask for. */
export type FillKind = NonNullable<ShapeFillProps['type']>

/**
 * Resolve which fill kind a set of fill or line props asks for.
 *
 * **An explicit `type` always wins**, even against a sub-object that disagrees:
 * `{ type: 'solid', gradient }` is a solid fill whose gradient is ignored. That is the rule
 * `type: 'none'` needs in order to mean anything (`{ type: 'none', color }` must stay
 * transparent), so it is the rule everywhere rather than one kind's exception.
 *
 * With no `type`, the first sub-object present selects the kind, in declaration order —
 * `gradient`, then `pattern`, then `image`. Setting two without a `type` is already a
 * contradiction; the order exists so the answer is stated rather than emergent.
 *
 * Everything else is `'solid'`, which is what `type` documents as its default.
 *
 * This is the one answer. It used to be given in four places: the stroke emitter inferred
 * `gradient` only, the shared dispatcher inferred nothing at all (so `fill: { gradient }`
 * painted a black `<a:solidFill>` and blamed the caller's colours for it), and three
 * `define/` modules each carried their own copy of the image half.
 */
export function resolveFillKind(props: Color | ShapeFillProps | ShapeLineProps | undefined): FillKind {
	if (!props || typeof props === 'string') return 'solid'
	if (props.type) return props.type
	if (props.gradient) return 'gradient'
	if (props.pattern) return 'pattern'
	if ('image' in props && props.image) return 'image'
	return 'solid'
}

/** The paint kinds a stroke can ask for: {@link FillKind} without the one `<a:ln>` has no slot for. */
export type LineFillKind = Exclude<FillKind, 'image'>

/**
 * {@link resolveFillKind} for a stroke, which has one kind fewer.
 *
 * `<a:ln>`'s paint child is `EG_LineFillProperties` — `a:noFill`, `a:solidFill`, `a:gradFill`,
 * `a:pattFill` — with no `a:blipFill` among them, so a picture stroke has no OOXML expression
 * and `ShapeLineProps` subtracts `image` from the fill props it inherits. This is where the
 * subtraction is enforced at run time, for the JS caller TypeScript cannot stop: the request
 * is refused rather than painted as nothing, because emitting a blipFill inside `a:ln` is a
 * package PowerPoint reports as needing repair, and emitting nothing is a stroke the caller
 * asked to be a picture and got in the theme colour without being told.
 *
 * Every stroke site resolves through here — the two `define/` rebuilds and the emitter — so
 * the refusal names the call the caller made rather than surfacing at serialization time.
 */
export function resolveLineKind(props: ShapeLineProps | undefined): LineFillKind {
	const kind = resolveFillKind(props)
	if (kind === 'image')
		throw new UnsupportedFeatureError(
			'line/image-fill-unsupported',
			'A picture stroke is not expressible in OOXML: `a:ln` accepts noFill/solidFill/gradFill/pattFill only. Use `fill: { image }` for a picture interior, or give the line a solid, gradient or pattern paint.'
		)
	return kind
}

/**
 * Does this fill or line props object name a paint at all?
 *
 * The question every *optional* paint site has to answer before dispatching: an outline, a
 * slide background and a table cell each inherit when the caller said nothing, so emitting
 * a fill child for an empty props object would paint them default black.
 *
 * "Said nothing" is narrower than "is empty". `'inherit'` is the explicit spelling of the
 * same silence and is false here. `'none'` is a statement — `<a:noFill/>` — and is true.
 * A `'solid'` kind needs a `color` to be saying anything; every other kind carries its
 * payload in its own sub-object, so its presence is the statement.
 */
export function fillNamesPaint(props: Color | ShapeFillProps | ShapeLineProps | undefined): boolean {
	if (!props) return false
	if (typeof props === 'string') return props.length > 0
	const kind = resolveFillKind(props)
	if (kind === 'inherit') return false
	if (kind === 'solid') return Boolean(props.color)
	return true
}

/**
 * A solid paint for {@link genXmlColorSelection}, carrying `transparency` only when one was set.
 *
 * Every caller assembles this out of an optional source — `border.transparency`,
 * `opts.transparency` — and writing the key with an `undefined` in it would be a second spelling
 * of "no transparency": one that reads the same here, but not in a bag that is spread over
 * another. So the key is simply absent, which is this codebase's one spelling of unset.
 * @param color - the colour to paint, already resolved against whatever default applies
 * @param transparency - percent transparency, or `undefined` for none
 */
export function solidPaint(color: Color, transparency: number | undefined): ShapeFillProps {
	return transparency === undefined ? { color } : { color, transparency }
}

/**
 * Emit the fill group a set of fill/line props asks for, dispatching on
 * {@link resolveFillKind}. Returns '' for props that name nothing to paint, so a caller
 * with an optional paint can hand this whatever it has — including `undefined`.
 * @param {Color | ShapeFillProps | ShapeLineProps | undefined} props fill props
 * @returns XML string
 */
export function genXmlColorSelection(props: Color | ShapeFillProps | ShapeLineProps | undefined): string {
	let colorVal = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		const fillType = resolveFillKind(props)
		if (typeof props === 'string') colorVal = props
		else {
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
