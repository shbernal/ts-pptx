/**
 * ts-pptx: Media Methods
 */

import { IMG_BROKEN } from '../media/placeholders.js'
import type { PresSlideInternal, SlideLayoutInternal, SlideMasterInternal, SlideRelMedia } from '../types/internal.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import { toMediaDataUri } from '../media/base64.js'
import { warn } from '../diagnostics.js'
import { MediaError } from '../errors.js'
import { isPreencodedPath } from './utils.js'

type SlideMediaRelWithPath = SlideRelMedia & { path: string }

function hasEncodingPath(rel: SlideRelMedia): rel is SlideMediaRelWithPath {
	return typeof rel.path === 'string' && rel.path.length > 0 && !isPreencodedPath(rel.path)
}

/**
 * Write an SVG rel's PNG fallback, applying `onMediaError` to a preview that fails.
 *
 * The one place that policy is applied to a preview. The adapter has already stamped the
 * placeholder and resolved with the failure, so only the PNG fallback is ever the placeholder:
 * the SVG part keeps its bytes under either policy. A preview used to reject instead, and an SVG
 * given inline then failed the whole write whatever `onMediaError` said, while one given by path
 * was reported as `media/load-failed` and had its SVG bytes replaced with a broken image.
 */
async function writeSvgPreview(
	rel: SlideRelMedia,
	runtime: RuntimeAdapter,
	onMediaError: 'throw' | 'placeholder'
): Promise<string> {
	const failure = await runtime.createSvgPngPreview(rel)
	if (failure === null) return 'done'
	if (onMediaError === 'placeholder') {
		warn('media/svg-preview-failed', `${failure.message}; the SVG is kept and its PNG fallback is a placeholder.`)
		return 'done'
	}
	throw failure
}

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlideInternal | SlideLayoutInternal | SlideMasterInternal} layout - slide, layout or master
 * @param {RuntimeAdapter} runtime - runtime adapter (Node/browser media loader)
 * @param {'throw' | 'placeholder'} onMediaError - failure policy: reject the export (default) or substitute a placeholder and warn
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(
	layout: PresSlideInternal | SlideLayoutInternal | SlideMasterInternal,
	runtime: RuntimeAdapter,
	onMediaError: 'throw' | 'placeholder' = 'throw'
): Array<Promise<string>> {
	const imageProms: Array<Promise<string>> = []

	// STEP 0: The default video poster, resolved here rather than where `addMedia` runs. A class
	// method body is never tree-shaken, so naming the artwork in `addMediaDefinition` charged its
	// whole base64 payload to every consumer, including one who only ever wrote text boxes. The
	// import is dynamic on purpose: a static one would pull `media/playbtn.ts` back into this
	// chunk, which the write path always reaches.
	const defaultCoverRels = layout._relsMedia.filter((rel) => rel.isDefaultCover && !rel.data)
	if (defaultCoverRels.length > 0)
		imageProms.push(
			(async () => {
				const { IMG_PLAYBTN } = await import('../media/playbtn.js')
				for (const rel of defaultCoverRels) rel.data = IMG_PLAYBTN
				return 'done'
			})()
		)

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
					// The rels that point at this one's package part. Resolved before the load so both
					// the success and the placeholder arm below hand them the same bytes.
					const dupes = candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path)
					let data: string
					try {
						data = toMediaDataUri(await runtime.loadMedia(rel), rel.type)
					} catch (ex) {
						if (onMediaError === 'placeholder') {
							warn(
								'media/load-failed',
								`Failed to load media "${rel.path}"; embedding a broken-image placeholder. (${String(ex)})`
							)
							rel.data = IMG_BROKEN
							dupes.forEach((dupe) => (dupe.data = IMG_BROKEN))
							return 'done'
						}
						// Default: fail-fast with an actionable error that names the failing asset and
						// chains the original cause (the raw fs/network error alone does not say which
						// media path broke). Pass `onMediaError: 'placeholder'` to degrade gracefully.
						throw new MediaError('media/load-failed', `Failed to load media "${rel.path}" during export.`, {
							cause: ex,
						})
					}
					rel.data = data
					dupes.forEach((dupe) => (dupe.data = data))
					// Outside the load's `try`: the load succeeded, so a preview that fails is not a load
					// failure, and must not replace the SVG's own bytes with a broken image.
					if (rel.isSvgPng) await writeSvgPreview(rel, runtime, onMediaError)
					// A path-deduped rel can itself be an SVG-PNG preview (the same SVG *file*
					// placed 2+ times on one slide: each placement pushes its own fallback rel).
					// Such dupes are skipped by STEP 5 — its `rel.data` filter runs synchronously,
					// before this async load populates `dupe.data` — so convert them here, or the
					// fallback keeps raw SVG bytes in a `.png` part and corrupts the deck.
					await Promise.all(
						dupes.filter((dupe) => dupe.isSvgPng).map((dupe) => writeSvgPreview(dupe, runtime, onMediaError))
					)
					return 'done'
				})()
			)
		})

	// STEP 5: SVG-PNG previews
	// ......: "SVG:" base64 data still requires a png to be generated
	// ......: (`isSvgPng` flag this as the preview image, not the SVG itself)
	layout._relsMedia
		.filter((rel) => rel.isSvgPng && rel.data)
		.forEach((rel) => {
			imageProms.push(writeSvgPreview(rel, runtime, onMediaError))
		})

	return imageProms
}
