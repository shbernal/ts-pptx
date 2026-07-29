/**
 * ts-pptx: Media Methods
 */

import { IMG_BROKEN } from '../core-enums-internal.js'
import type { PresSlideInternal, SlideLayoutInternal, SlideRelMedia } from '../types/internal.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import { warn } from '../diagnostics.js'
import { MediaError } from '../errors.js'

type SlideMediaRelWithPath = SlideRelMedia & { path: string }

function hasEncodingPath(rel: SlideRelMedia): rel is SlideMediaRelWithPath {
	return typeof rel.path === 'string' && rel.path.length > 0 && !rel.path.includes('preencoded')
}

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlideInternal | SlideLayoutInternal} layout - slide layout
 * @param {RuntimeAdapter} runtime - runtime adapter (Node/browser media loader)
 * @param {'throw' | 'placeholder'} onMediaError - failure policy: reject the export (default) or substitute a placeholder and warn
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(
	layout: PresSlideInternal | SlideLayoutInternal,
	runtime: RuntimeAdapter,
	onMediaError: 'throw' | 'placeholder' = 'throw'
): Array<Promise<string>> {
	const imageProms: Array<Promise<string>> = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	const candidateRels = layout._relsMedia.filter(
		(rel): rel is SlideMediaRelWithPath => rel.type !== 'online' && !rel.data && hasEncodingPath(rel)
	)

	// B: PERF: Mark dupes (same `path`) to avoid loading the same media over-and-over!
	const unqPaths: string[] = []
	candidateRels.forEach((rel) => {
		if (!unqPaths.includes(rel.path)) {
			rel.isDuplicate = false
			unqPaths.push(rel.path)
		} else {
			rel.isDuplicate = true
		}
	})

	// STEP 4: Read/Encode each unique media item
	candidateRels
		.filter((rel) => !rel.isDuplicate)
		.forEach((rel) => {
			imageProms.push(
				(async () => {
					try {
						rel.data = await runtime.loadMedia(rel)
						const dupes = candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path)
						dupes.forEach((dupe) => (dupe.data = rel.data))
						if (rel.isSvgPng) await runtime.createSvgPngPreview(rel)
						// A path-deduped rel can itself be an SVG-PNG preview (the same SVG *file*
						// placed 2+ times on one slide: each placement pushes its own fallback rel).
						// Such dupes are skipped by STEP 5 — its `rel.data` filter runs synchronously,
						// before this async load populates `dupe.data` — so convert them here, or the
						// fallback keeps raw SVG bytes in a `.png` part and corrupts the deck.
						await Promise.all(dupes.filter((dupe) => dupe.isSvgPng).map((dupe) => runtime.createSvgPngPreview(dupe)))
						return 'done'
					} catch (ex) {
						if (onMediaError === 'placeholder') {
							warn(
								'media/load-failed',
								`Failed to load media "${rel.path}"; embedding a broken-image placeholder. (${String(ex)})`
							)
							rel.data = IMG_BROKEN
							candidateRels
								.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path)
								.forEach((dupe) => (dupe.data = rel.data))
							return 'done'
						}
						// Default: fail-fast with an actionable error that names the failing asset and
						// chains the original cause (the raw fs/network error alone does not say which
						// media path broke). Pass `onMediaError: 'placeholder'` to degrade gracefully.
						throw new MediaError('media/load-failed', `Failed to load media "${rel.path}" during export.`, {
							cause: ex,
						})
					}
				})()
			)
		})

	// STEP 5: SVG-PNG previews
	// ......: "SVG:" base64 data still requires a png to be generated
	// ......: (`isSvgPng` flag this as the preview image, not the SVG itself)
	layout._relsMedia
		.filter((rel) => rel.isSvgPng && rel.data)
		.forEach((rel) => {
			imageProms.push(runtime.createSvgPngPreview(rel))
		})

	return imageProms
}
