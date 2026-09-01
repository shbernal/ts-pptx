/**
 * ts-pptx: DrawingML image sizing & cropping
 *
 * Build the `<a:srcRect>` (+ `<a:stretch>`) blipFill children for the image
 * sizing modes (`cover` / `contain` / `crop` / `stretch`), for the implicit
 * aspect-correct placement a vector source gets, and for explicit percentage crops.
 */

import { fitSrcRectPercents } from '../../media/image-size.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { genXmlImageCropRect, type ImageCrop } from './src-rect.js'
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
 * The blipFill children for an explicit percentage crop: the shared `<a:srcRect>` plus the
 * fill directive every sizing mode above pairs it with.
 * @param crop - edge insets in percent; omitted edges default to 0
 * @param objectName - for error messages
 * @returns the `<a:srcRect>` + `<a:stretch>` blipFill children
 */
export function genXmlImageCrop(crop: ImageCrop, objectName?: string): string {
	const where = objectName ? ` for image "${objectName}"` : ''
	return genXmlImageCropRect(crop, 'addImage', where) + STRETCH
}
