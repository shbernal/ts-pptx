/**
 * ts-pptx: DrawingML image sizing & cropping
 *
 * Build the `<a:srcRect>` (+ `<a:stretch>`) blipFill children for the image
 * sizing modes (`cover` / `contain` / `crop`) and for explicit percentage crops.
 */

import { fitSrcRectPercents } from '../../media/image-size.js'
import { FIXED_PCT_PER_PERCENT } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { InvalidOptionError } from '../../errors.js'

/** Every `<a:srcRect>` below is followed by this same fill directive. */
const STRETCH = el('a:stretch', null, raw(voidEl('a:fillRect')))

export const ImageSizingXml = {
	cover: function (imgSize: { w: number; h: number }, boxDim: { w: number; h: number; x: number; y: number }) {
		const { l, r, t, b } = fitSrcRectPercents('cover', imgSize, boxDim)
		return voidEl('a:srcRect', { l, r, t, b }) + STRETCH
	},
	contain: function (imgSize: { w: number; h: number }, boxDim: { w: number; h: number; x: number; y: number }) {
		const { l, r, t, b } = fitSrcRectPercents('contain', imgSize, boxDim)
		return voidEl('a:srcRect', { l, r, t, b }) + STRETCH
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
 * Build an explicit `<a:srcRect>` crop from percentage edge insets (0–100), emitted verbatim.
 * Each inset is the percent of the *source* image trimmed from that edge; values serialize in
 * 1000ths of a percent (DrawingML `ST_Percentage`, where 100% = 100000). Unlike `sizing: 'crop'`
 * — which derives the rect from displayed inches and the natural pixel size — this maps a
 * sub-region of the source directly, so it works for SVG and other unmeasurable formats.
 * @param crop - edge insets in percent; omitted edges default to 0
 * @param objectName - for error messages
 * @returns the `<a:srcRect>` + `<a:stretch>` blipFill children
 */
export function genXmlImageCrop(crop: { l?: number; t?: number; r?: number; b?: number }, objectName?: string): string {
	const where = objectName ? ` for image "${objectName}"` : ''
	const edges = { l: crop.l ?? 0, t: crop.t ?? 0, r: crop.r ?? 0, b: crop.b ?? 0 }
	for (const [name, val] of Object.entries(edges)) {
		if (typeof val !== 'number' || !isFinite(val) || val < 0 || val > 100) {
			throw new InvalidOptionError(
				'image/crop-inset-out-of-range',
				`addImage crop.${name} must be a percentage between 0 and 100 (got ${String(val)})${where}.`
			)
		}
	}
	if (edges.l + edges.r >= 100)
		throw new InvalidOptionError(
			'image/crop-insets-exceed-extent',
			`addImage crop: left+right insets (${edges.l}%+${edges.r}%) must be < 100%${where}.`
		)
	if (edges.t + edges.b >= 100)
		throw new InvalidOptionError(
			'image/crop-insets-exceed-extent',
			`addImage crop: top+bottom insets (${edges.t}%+${edges.b}%) must be < 100%${where}.`
		)
	const v = (perc: number): number => Math.round(perc * FIXED_PCT_PER_PERCENT)
	// NOTE: attribute order here is l/t/r/b, where the sizing modes above emit l/r/t/b. Attribute
	// order is byte-significant, so the two orderings are kept as they are rather than unified.
	return voidEl('a:srcRect', { l: v(edges.l), t: v(edges.t), r: v(edges.r), b: v(edges.b) }) + STRETCH
}
