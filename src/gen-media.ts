/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN } from './core-enums.js'
import type { PresSlideInternal, SlideLayoutInternal, ISlideRelMedia } from './core-interfaces.js'
import type { RuntimeAdapter } from './runtime/types.js'

type SlideMediaRelWithPath = ISlideRelMedia & { path: string }

function hasEncodingPath(rel: ISlideRelMedia): rel is SlideMediaRelWithPath {
	return typeof rel.path === 'string' && rel.path.length > 0 && !rel.path.includes('preencoded')
}

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlideInternal | SlideLayoutInternal} layout - slide layout
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(layout: PresSlideInternal | SlideLayoutInternal, runtime: RuntimeAdapter): Array<Promise<string>> {
	const imageProms: Array<Promise<string>> = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	const candidateRels = layout._relsMedia.filter(
		(rel): rel is SlideMediaRelWithPath => rel.type !== 'online' && !rel.data && hasEncodingPath(rel)
	)

	// B: PERF: Mark dupes (same `path`) to avoid loading the same media over-and-over!
	const unqPaths: string[] = []
	candidateRels.forEach(rel => {
		if (!unqPaths.includes(rel.path)) {
			rel.isDuplicate = false
			unqPaths.push(rel.path)
		} else {
			rel.isDuplicate = true
		}
	})

	// STEP 4: Read/Encode each unique media item
	candidateRels
		.filter(rel => !rel.isDuplicate)
		.forEach(rel => {
			imageProms.push(
				(async () => {
					try {
						rel.data = await runtime.loadMedia(rel)
						candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
						if (rel.isSvgPng) await runtime.createSvgPngPreview(rel)
						return 'done'
					} catch (ex) {
						rel.data = IMG_BROKEN
						candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
						throw ex
					}
				})(),
			)
		})

	// STEP 5: SVG-PNG previews
	// ......: "SVG:" base64 data still requires a png to be generated
	// ......: (`isSvgPng` flag this as the preview image, not the SVG itself)
	layout._relsMedia
		.filter(rel => rel.isSvgPng && rel.data)
		.forEach(rel => {
			imageProms.push(runtime.createSvgPngPreview(rel))
		})

	return imageProms
}
