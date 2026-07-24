/**
 * ts-pptx: preview/cover image registration.
 *
 * Several objects are drawn from a cached raster the library cannot render itself — a Zoom tile's
 * thumbnail of its target slide, an OLE object's picture of the embedded document. Each registers
 * that raster as an ordinary slide image rel (deduped exactly like `addImage`) and references it by
 * rId. Shared here so the two definers agree on extension sniffing, content type, and de-dup.
 */
import type { PresSlideInternal } from '../../types/internal.js'
import { getNewRelId } from '../../gen-utils.js'
import { imageContentType } from '../../media/content-type.js'

/** 32×32 solid #E7E6E6 PNG — the neutral placeholder shown when the caller supplies no cover image. */
export const PLACEHOLDER_PNG =
	'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAxSURBVFhH7c4hAQAACMAw+geFGOAJAGbi5mpRmf1Z7HEdAAAAAAAAAAAAAAAAAADAAPwOyNPFH3F8AAAAAElFTkSuQmCC'

/**
 * Register a preview/cover image as a slide media rel (deduped like `addImage`) and return its rId.
 * Falls back to the gray placeholder PNG when no cover image is supplied.
 * @param target - slide the image rel is registered on
 * @param cover - caller-supplied cover image (`path` or base64 `data`); omitted → gray placeholder
 * @returns the rId of the image rel
 */
export function registerPreviewImage(target: PresSlideInternal, cover?: { path?: string; data?: string }): number {
	const strImagePath = cover?.path || ''
	let strImageData = cover?.data || ''
	let extn = 'png'
	if (strImagePath) {
		const file = strImagePath.slice(strImagePath.lastIndexOf('/') + 1).split('?')[0] || ''
		extn = ((file.split('.').pop() || 'png').split('#')[0] || 'png').toLowerCase()
	} else if (strImageData) {
		const mime = /image\/(\w+);/.exec(strImageData)
		if (mime) extn = mime[1] ?? 'png'
	} else {
		strImageData = 'image/png;base64,' + PLACEHOLDER_PNG
	}

	const rId = getNewRelId(target)
	const mediaSlideKey =
		target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum
	const type = imageContentType(extn)
	const dupe = target._relsMedia.find((item) => {
		if (item.isDuplicate || !item.Target || item.type !== type) return false
		return strImagePath ? item.path === strImagePath : item.data === strImageData
	})
	target._relsMedia.push({
		path: strImagePath || 'preencoded.' + extn,
		type,
		extn,
		data: strImageData || '',
		rId,
		isDuplicate: !!dupe?.Target,
		Target: dupe?.Target ? dupe.Target : `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${extn}`,
	})
	return rId
}
