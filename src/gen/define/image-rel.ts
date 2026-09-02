/**
 * ts-pptx: the one place an image is registered as a slide media relationship.
 *
 * Three definers need the identical fifteen lines — an image *fill* on a shape or text box
 * (`registerImageFillMedia`), an `addImage()` raster (`addImageDefinition`), and the cached
 * preview raster a Zoom tile or OLE object is drawn from (`registerPreviewImage`). They had
 * three copies, and the drift had already started: only one of them still carried the comment
 * explaining the path-versus-data match below.
 */
import type { PresSlideInternal } from '../../types/internal.js'
import { getNewRelId, mediaSlideKey } from '../utils.js'
import { imageContentType } from '../../media/content-type.js'

/**
 * Push an image media rel onto `target`, reusing an identical source's package part.
 *
 * De-dup is per target and by source, not by bytes: file-path images are matched on `path`,
 * while base64 `data` images have no real path — they all share the `preencoded.<extn>`
 * placeholder — so they are matched on their payload instead, which is what stops the same
 * inline image being embedded once per use. A rel already marked `isDuplicate` is never
 * matched against, so every duplicate points at the one original rather than at a chain.
 *
 * This is a *slide-local* optimization. `package/assemble.ts` runs a second, deck-wide collapse
 * keyed on extension + bytes once every rel's data is loaded, which subsumes this one for reuse
 * across slides and for sources that only turn out identical after loading.
 * @param target - slide (or layout/master) the rel is registered on
 * @param source - the resolved image source: a `path`, a base64 `data` payload, or both
 * @param relId - the relationship id already allocated for this use
 */
export function registerImageMediaRel(
	target: PresSlideInternal,
	source: { path?: string; data?: string; extn: string },
	relId: number
): void {
	const path = source.path || ''
	const data = source.data || ''
	const type = imageContentType(source.extn)
	const dupe = target._relsMedia.find((item) => {
		if (item.isDuplicate || !item.Target || item.type !== type) return false
		return path ? item.path === path : !!data && item.data === data
	})
	target._relsMedia.push({
		path: path || 'preencoded.' + source.extn,
		type,
		extn: source.extn,
		data,
		rId: relId,
		isDuplicate: !!dupe?.Target,
		// `_relsMedia.length + 1` is read BEFORE this push, so the first rel on a slide lands on
		// `image-<key>-1`. Keep the read here rather than hoisting it.
		Target: dupe?.Target
			? dupe.Target
			: `../media/image-${mediaSlideKey(target)}-${target._relsMedia.length + 1}.${source.extn}`,
	})
}

/**
 * Push the *pair* of media rels an SVG source consumes: a rasterized PNG fallback (what a
 * renderer without SVG support paints, and what `<a:blip r:embed>` points at) and the SVG
 * itself (`asvg:svgBlip`).
 *
 * `addImage` and the picture-bullet definer both need this, and both used to take the SVG's
 * id as `pngRid + 1` on faith. That held only while nothing else allocated in between, and
 * `addImage`'s own hyperlink then took the same number a third time. Both ids now come from
 * {@link getNewRelId}, which skips every id the slide already holds.
 *
 * Neither push goes through {@link registerImageMediaRel} and neither needs to. The PNG
 * fallback is rasterized per call from a per-call `svgSize`, so two uses at different sizes
 * are genuinely two different images; the SVG source has no such excuse, but the deck-wide
 * collapse in `package/assemble.ts` keys on extension + bytes and merges them, so the same
 * SVG placed twice measures as a single `ppt/media/*.svg` part.
 *
 * @param target - slide (or layout/master) the rels are registered on
 * @param source - the resolved SVG source: a `path`, a base64 `data` payload, or both, plus
 *   the display size the PNG fallback is rasterized at when the caller knows it
 * @param pinned - ids to reuse rather than allocate. Auto-paging re-registers the same bullet
 *   on each overflow slide while sharing one options object by reference, so the pair has to
 *   keep the ids that object already carries.
 * @returns the two allocated relationship ids
 */
export function registerSvgImageRels(
	target: PresSlideInternal,
	source: { path: string; data: string; svgSize?: { w: number; h: number } },
	pinned?: { pngRid: number; svgRid: number }
): { pngRid: number; svgRid: number } {
	const { path, data } = source
	const mediaKey = mediaSlideKey(target)

	const pngRid = pinned ? pinned.pngRid : getNewRelId(target)
	target._relsMedia.push({
		path: path || data + 'png',
		type: 'image/png',
		extn: 'png',
		data,
		rId: pngRid,
		// `_relsMedia.length + 1` is read BEFORE each push, so the pair lands on consecutive names.
		Target: `../media/image-${mediaKey}-${target._relsMedia.length + 1}.png`,
		isSvgPng: true,
		...(source.svgSize ? { svgSize: source.svgSize } : {}),
	})

	const svgRid = pinned ? pinned.svgRid : getNewRelId(target)
	target._relsMedia.push({
		path: path || data || 'preencoded.svg',
		type: 'image/svg+xml',
		extn: 'svg',
		data,
		rId: svgRid,
		Target: `../media/image-${mediaKey}-${target._relsMedia.length + 1}.svg`,
	})

	return { pngRid, svgRid }
}
