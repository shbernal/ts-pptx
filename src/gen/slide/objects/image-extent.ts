/**
 * ts-pptx: the extent an image is drawn at
 *
 * An image's frame is not always the box it is drawn in. An omitted `w` or `h` is backfilled from
 * the image's natural size, and a `sizing` box replaces the frame outright. The image renderer emits
 * that extent, and a group sizing itself around an image child has to see the same one, so both
 * read it here. This module writes no XML, which is why `gen/slide/object.ts` may import it.
 */

import type {
	PresSlideInternal,
	SlideLayoutInternal,
	SlideMasterInternal,
	SlideObject,
} from '../../../types/internal.js'
import { getImageSizeFromBase64 } from '../../../media/image-size.js'
import { pixelsToEmu } from '../../../units.js'
import { getSmartParseNumber } from '../../../units-internal.js'

/**
 * An image's natural size in pixels, measured from its embedded media.
 *
 * The bytes are not available when `addImage()` runs, but they are loaded by the time a slide is
 * serialized, which is when every caller of this runs.
 * @param obj - the image slide object
 * @param slide - the slide or layout holding it
 * @returns the natural size, or `null` when the media is not loaded or cannot be measured
 */
export function imageNaturalSize(
	obj: SlideObject,
	slide: PresSlideInternal | SlideLayoutInternal | SlideMasterInternal
): { w: number; h: number } | null {
	const relData = (slide._relsMedia || []).find((rel) => rel.rId === obj.imageRid)?.data
	return typeof relData === 'string' ? getImageSizeFromBase64(relData) : null
}

/**
 * The two boxes an image is placed with, in EMU.
 *
 * `frameW`/`frameH` are the frame with any omitted dimension backfilled from the natural size.
 * PowerPoint inserts images at 96 DPI, so natural pixels / 96 is the display size in inches.
 * `drawnW`/`drawnH` are the box actually drawn and emitted: the `sizing` box when there is one and
 * no `crop` overrides it, and the frame otherwise.
 * @param obj - the image slide object
 * @param slide - the slide or layout holding it
 * @param frame - the object's resolved frame
 * @returns the backfilled frame and the drawn box
 */
export function resolveImageExtent(
	obj: SlideObject,
	slide: PresSlideInternal | SlideLayoutInternal | SlideMasterInternal,
	frame: { cx: number; cy: number }
): { frameW: number; frameH: number; drawnW: number; drawnH: number } {
	const opts = obj.options ?? {}
	let frameW = frame.cx
	let frameH = frame.cy
	if (opts._szAuto) {
		const szAuto = opts._szAuto
		const natural = imageNaturalSize(obj, slide)
		if (natural) {
			if (szAuto.w && szAuto.h) {
				frameW = pixelsToEmu(natural.w, 96)
				frameH = pixelsToEmu(natural.h, 96)
			} else if (szAuto.h) {
				// Width supplied, derive height
				frameH = Math.round(frameW * (natural.h / natural.w))
			} else if (szAuto.w) {
				// Height supplied, derive width
				frameW = Math.round(frameH * (natural.w / natural.h))
			}
		}
	}
	// `crop` wins over `sizing`; the renderer warns about the pair.
	const sizing = opts.crop ? undefined : opts.sizing
	if (!sizing?.type) return { frameW, frameH, drawnW: frameW, drawnH: frameH }
	return {
		frameW,
		frameH,
		drawnW: sizing.w ? getSmartParseNumber(sizing.w, 'X', slide._presLayout) : frameW,
		drawnH: sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide._presLayout) : frameH,
	}
}
