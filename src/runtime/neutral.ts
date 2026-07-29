import { IMG_SVG_PLACEHOLDER } from '../constants-internal.js'
import { MediaError, UnsupportedFeatureError } from '../errors.js'
import { bytesToBase64 } from '../media/base64.js'
import type { SlideRelMedia } from '../types/internal.js'
import type { RuntimeAdapter } from './types.js'

/**
 * The adapter for a runtime that is neither Node nor a browser — Deno, Bun, an edge worker,
 * anything that imports the bare `@shbernal/ts-pptx` specifier without resolving the `node` or
 * `browser` export condition.
 *
 * It implements everything that is genuinely host-neutral and refuses only what is not.
 * `fetch`, `btoa` and `TextEncoder` exist in all of these runtimes, so remote media and fonts
 * load here exactly as they do in a browser. What is missing is a *destination*: there is no
 * `node:fs` to write to and no `document` to hand a download to, so {@link RuntimeAdapter.writeFile}
 * throws a named error rather than guessing. Everything that returns bytes to the caller —
 * `write`, `stream`, `toParts` — works normally, which is the shape that suits a worker anyway.
 */
export function createNeutralRuntime(): RuntimeAdapter {
	return {
		writeFileOutputType: null,
		loadMedia,
		createSvgPngPreview,
		writeFile,
		loadFontData,
	}
}

async function loadFontData(source: string): Promise<Uint8Array> {
	const response = await fetch(source)
	if (!response.ok) throw new MediaError('font/fetch-failed', `Unable to load font (fetch): ${source}`)
	return new Uint8Array(await response.arrayBuffer())
}

/**
 * Fetch-only: a bare filesystem path has no meaning without a filesystem, and `fetch` rejecting
 * it is reported by the caller as a load failure naming the path. Returns raw base64 rather than
 * a data URI (the browser adapter's `FileReader` shape) because there is no `FileReader` here;
 * package assembly accepts either form.
 */
async function loadMedia(rel: SlideRelMedia & { path: string }): Promise<string> {
	const response = await fetch(rel.path)
	if (!response.ok) throw new MediaError('media/fetch-failed', `Unable to load image (fetch): ${rel.path}`)
	return bytesToBase64(new Uint8Array(await response.arrayBuffer()))
}

/** Rasterizing an SVG needs a canvas. Same fallback as Node: emit the placeholder PNG. */
async function createSvgPngPreview(rel: SlideRelMedia): Promise<string> {
	rel.data = IMG_SVG_PLACEHOLDER
	return 'done'
}

async function writeFile(fileName: string): Promise<string> {
	throw new UnsupportedFeatureError(
		'runtime/file-output-unavailable',
		`Cannot write "${fileName}": this runtime has neither a filesystem nor a DOM. ` +
			'Import "@shbernal/ts-pptx/node" to write to disk or "@shbernal/ts-pptx/browser" to trigger a ' +
			'download; anywhere else, use write() and persist the bytes yourself.'
	)
}
