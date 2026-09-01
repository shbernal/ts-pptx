/**
 * ts-pptx: DrawingML image sizing & cropping
 *
 * Build the `<a:srcRect>` (+ `<a:stretch>`) blipFill children for the image
 * sizing modes (`cover` / `contain` / `crop` / `stretch`), for the implicit
 * aspect-correct placement a vector source gets, and for explicit percentage crops.
 */

import { fitSrcRectPercents } from '../../media/image-size.js'
import { FIXED_PCT_PER_PERCENT } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { InvalidOptionError } from '../../errors.js'

/** Every `<a:srcRect>` below is followed by this same fill directive. */
const STRETCH = el('a:stretch', null, raw(voidEl('a:fillRect')))

/** The `<a:srcRect>` fitting `imgSize` into `boxDim`, followed by the fill directive. */
function fitSrcRect(
	type: 'cover' | 'contain',
	imgSize: { w: number; h: number },
	boxDim: { w: number; h: number }
): string {
	const { l, r, t, b } = fitSrcRectPercents(type, imgSize, boxDim)
	return voidEl('a:srcRect', { l, r, t, b }) + STRETCH
}

export const ImageSizingXml = {
	cover: function (imgSize: { w: number; h: number }, boxDim: { w: number; h: number; x: number; y: number }) {
		return fitSrcRect('cover', imgSize, boxDim)
	},
	contain: function (imgSize: { w: number; h: number }, boxDim: { w: number; h: number; x: number; y: number }) {
		return fitSrcRect('contain', imgSize, boxDim)
	},
	/** No source rectangle at all: the image fills the box, aspect ratio be damned. */
	stretch: function () {
		return STRETCH
	},
	crop: function (imgSize: { w: number; h: number }, boxDim: { w: number; h: number; x: number; y: number }) {
		const l = boxDim.x
		const r = imgSize.w - (boxDim.x + boxDim.w)
		const t = boxDim.y
		const b = imgSize.h - (boxDim.y + boxDim.h)
		if (l < 0 || r < 0 || t < 0 || b < 0) {
			const over = [
				l < 0 && `x (${l < 0 ? -l : 0} past left edge)`,
				r < 0 && `x+w (${-r} past right edge)`,
				t < 0 && `y (${-t} past top edge)`,
				b < 0 && `y+h (${-b} past bottom edge)`,
			]
				.filter(Boolean)
				.join(', ')
			throw new InvalidOptionError(
				'image/crop-window-overflows',
				`addImage sizing.type 'crop': crop window overflows image bounds — ${over}. Ensure x≥0, y≥0, x+w≤w, y+h≤h.`
			)
		}
		const lPerc = Math.round(1e5 * (l / imgSize.w))
		const rPerc = Math.round(1e5 * (r / imgSize.w))
		const tPerc = Math.round(1e5 * (t / imgSize.h))
		const bPerc = Math.round(1e5 * (b / imgSize.h))
		return voidEl('a:srcRect', { l: lPerc, r: rPerc, t: tPerc, b: bPerc }) + STRETCH
	},
}

/**
 * The blipFill children for a vector source the caller gave no `sizing` for: letterbox it to its
 * own aspect ratio inside its frame, the way `sizing: 'contain'` would — an SVG carries an
 * intrinsic ratio, and stretching a glyph into a box that disagrees with it is never the intent.
 *
 * Returns `null` when the two ratios already agree, which is the overwhelmingly common case (a
 * square icon in a square box). The caller then emits the plain `<a:stretch>` it always did, so
 * this default costs nothing — neither bytes nor a behaviour change — until it actually differs.
 * @param imgSize - the vector's intrinsic size (viewBox or width/height)
 * @param boxDim - the picture's displayed frame
 * @returns the `<a:srcRect>` + `<a:stretch>` children, or `null` if no letterboxing is needed
 */
export function genXmlVectorAspectFit(
	imgSize: { w: number; h: number },
	boxDim: { w: number; h: number }
): string | null {
	const { l, r, t, b } = fitSrcRectPercents('contain', imgSize, boxDim)
	if (l === 0 && r === 0 && t === 0 && b === 0) return null
	return fitSrcRect('contain', imgSize, boxDim)
}

/**
 * Build an explicit `<a:srcRect>` crop from percentage edge insets (0–100), emitted verbatim.
 * Each inset is the percent of the *source* image trimmed from that edge; values serialize in
 * 1000ths of a percent (DrawingML `ST_Percentage`, where 100% = 100000). Unlike `sizing: 'crop'`
 * — which derives the rect from displayed inches and the natural pixel size — this maps a
 * sub-region of the source directly, so it works for SVG and other unmeasurable formats.
 * @param crop - edge insets in percent; omitted edges default to 0
 * @param objectName - for error messages
 * @returns the `<a:srcRect>` + `<a:stretch>` blipFill children
 */
type ImageCrop = { l?: number; t?: number; r?: number; b?: number }

function normalizeImageCrop(
	crop: ImageCrop,
	label: string,
	where: string
): { l: number; t: number; r: number; b: number } {
	const edges = { l: crop.l ?? 0, t: crop.t ?? 0, r: crop.r ?? 0, b: crop.b ?? 0 }
	for (const [name, val] of Object.entries(edges)) {
		if (typeof val !== 'number' || !Number.isFinite(val) || val < 0 || val > 100) {
			throw new InvalidOptionError(
				'image/crop-inset-out-of-range',
				`${label} crop.${name} must be a percentage between 0 and 100 (got ${String(val)})${where}.`
			)
		}
	}
	if (edges.l + edges.r >= 100)
		throw new InvalidOptionError(
			'image/crop-insets-exceed-extent',
			`${label} crop: left+right insets (${edges.l}%+${edges.r}%) must be < 100%${where}.`
		)
	if (edges.t + edges.b >= 100)
		throw new InvalidOptionError(
			'image/crop-insets-exceed-extent',
			`${label} crop: top+bottom insets (${edges.t}%+${edges.b}%) must be < 100%${where}.`
		)
	return edges
}

/**
 * Build an explicit `<a:srcRect>` crop from percentage edge insets (0–100), emitted verbatim.
 * This is the shared source-rectangle half used by both picture objects and native picture fills.
 */
export function genXmlImageCropRect(crop: ImageCrop, label = 'image fill', where = ''): string {
	const edges = normalizeImageCrop(crop, label, where)
	const v = (perc: number): number => Math.round(perc * FIXED_PCT_PER_PERCENT)
	// NOTE: attribute order here is l/t/r/b, where the sizing modes above emit l/r/t/b. Attribute
	// order is byte-significant, so the two orderings are kept as they are rather than unified.
	return voidEl('a:srcRect', { l: v(edges.l), t: v(edges.t), r: v(edges.r), b: v(edges.b) })
}

export function genXmlImageCrop(crop: ImageCrop, objectName?: string): string {
	const where = objectName ? ` for image "${objectName}"` : ''
	return genXmlImageCropRect(crop, 'addImage', where) + STRETCH
}
