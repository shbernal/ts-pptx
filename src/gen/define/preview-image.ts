/**
 * ts-pptx: preview/cover image registration.
 *
 * Several objects are drawn from a cached raster the library cannot render itself — a Zoom tile's
 * thumbnail of its target slide, an OLE object's picture of the embedded document. Each registers
 * that raster as an ordinary slide image rel and references it by rId. Shared here so the two
 * definers agree on the placeholder fallback and on extension sniffing; the rel itself goes through
 * `registerImageMediaRel`, the same call `addImage` and image fills use.
 */
import type { PresSlideInternal } from '../../types/internal.js'
import { getNewRelId } from '../utils.js'
import { imageExtensionForSource } from '../../media/content-type.js'
import { registerImageMediaRel } from './image-rel.js'

/** 32×32 solid #E7E6E6 PNG — the neutral placeholder shown when the caller supplies no cover image. */
const PLACEHOLDER_PNG =
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
	if (!strImagePath && !strImageData) strImageData = 'image/png;base64,' + PLACEHOLDER_PNG
	const extn = imageExtensionForSource(strImagePath, strImageData)

	const rId = getNewRelId(target)
	registerImageMediaRel(target, { path: strImagePath, data: strImageData, extn }, rId)
	return rId
}
