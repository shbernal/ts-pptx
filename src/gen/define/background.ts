/**
 * ts-pptx: Background Definition
 *
 * Registers an `addBackground()` image as a slide media rel (color backgrounds carry no rel);
 * the `<p:bg>` XML is emitted later at slide / layout serialize time.
 */
import type { BackgroundOption } from '../../types/index.js'
import type { PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'
import { imageContentType, imageExtensionForSource } from '../../media/content-type.js'
import { getNewRelId, preencodedPath } from '../utils.js'
import { pushMediaRel } from './image-rel.js'

/**
 * Adds a background image or color to a slide definition.
 *
 * Only an image background does anything here: it is the one that needs a media relationship.
 * A colour — object or the bare-string shorthand — carries no rel and is emitted straight from
 * `slide.background` at serialize time.
 *
 * The rel id and the part name come from the shared allocators. This used to take the id as
 * `_relsMedia.length + 1`, which a hyperlink or chart rel could already hold, and to name the part
 * after the slide or layout title. Titles are caller-supplied and different ones sanitize alike,
 * so "A B" and "A-B", or a layout titled "Slide 1" beside slide 1, wrote one part that both
 * backgrounds pointed at. `nextMediaTarget` keys the name by slide, layout or master, never
 * by title, so it is unique and needs no sanitizing.
 *
 * Called on every assignment, so it replaces what an earlier one registered. A new image takes
 * over the earlier background's rel in place, keeping its id and its part name, rather than
 * pushing a second one; replacing it by removal would shift the media count every later part name
 * is read from. A colour clears the image rel id, which `slideBackgroundXml` checks before it looks
 * at a colour, so an image replaced by a colour stops painting.
 * @param {BackgroundOption} props - a bare colour, or an object with a colour or image definition
 * @param {SlideLayoutInternal} target - slide or layout that the background is set on
 */
export function addBackgroundDefinition(props: BackgroundOption | undefined, target: SlideLayoutInternal): void {
	if (!(props && typeof props === 'object' && (props.path || props.data))) {
		delete target._bkgdImgRid
		return
	}

	// The `data:` mime wins over `path`, as it does for `addImage()`: a background supplied as
	// bytes alone used to fall back on a PNG placeholder path and declare `image/png` no matter
	// what it actually carried, so `{ data: 'data:image/svg+xml;…' }` shipped SVG bytes in a part
	// the package announced as PNG.
	let strImgExtn = imageExtensionForSource(props.path || '', props.data || '')
	if (strImgExtn === 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup
	const type = imageContentType(strImgExtn)

	const previous = target._relsMedia.findIndex((rel) => rel.rId === target._bkgdImgRid)
	const replaced = target._relsMedia[previous]
	if (replaced) {
		// The record `pushMediaRel` would write, on the rel and part the earlier background holds.
		// Kept local: `props` is the caller's own object on the `slide.background =` path, not a clone.
		target._relsMedia[previous] = {
			path: props.path || preencodedPath(strImgExtn),
			type,
			extn: strImgExtn,
			data: props.data || '',
			rId: replaced.rId,
			Target: replaced.Target.replace(/\.[^./]+$/, `.${strImgExtn}`),
		}
		return
	}

	// The allocators take a slide. A layout carries the same rel lists and a slide number of its
	// own, which is all they read, and `master.ts` passes one the same way.
	const slide = target as PresSlideInternal
	const rId = getNewRelId(slide)
	pushMediaRel(slide, { kind: 'image', extn: strImgExtn, type, path: props.path, data: props.data, rId })
	target._bkgdImgRid = rId
}
