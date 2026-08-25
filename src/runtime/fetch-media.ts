/**
 * The parts of a runtime adapter that are not runtime-specific.
 *
 * `fetch`, `btoa` and `TextEncoder` exist in Node, in every browser, and in the edge/worker
 * runtimes the neutral adapter serves, so the remote-load paths were the same code in all three
 * adapters — byte-identical in two of them and an inline branch in the third. What is genuinely
 * per-runtime stays in each adapter: `node:fs`, `FileReader`, the canvas rasterizer, the
 * `<a download>` writer.
 *
 * Nothing here touches a host global beyond `fetch`.
 */

import { IMG_SVG_PLACEHOLDER } from '../constants-internal.js'
import { MediaError } from '../errors.js'
import { bytesToBase64 } from '../media/base64.js'
import type { SlideRelMedia } from '../types/internal.js'

/**
 * Whether `source` names something to `fetch` rather than something to read from a filesystem.
 *
 * A real scheme test, not `source.startsWith('http')`. That prefix matched a *relative path* —
 * `httpdocs/logo.png`, `http-icons/x.svg` — and sent it to `fetch`, which then failed with a
 * fetch error naming a file that exists on disk. It was also case-sensitive, so
 * `HTTPS://example.com/f.ttf` missed and went to `fs.readFile`.
 *
 * Only `http:` and `https:` count. A `file:` or `data:` URL falls through to the adapter's
 * local branch, where it fails — the same as before this was tightened, and deliberately left
 * that way: making either work is a feature with its own decisions (what a `file:` URL means
 * off Node, whether a `data:` URL should bypass loading entirely), not a side effect of fixing
 * a prefix test.
 */
export function isRemote(source: string): boolean {
	const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(source)?.[1]?.toLowerCase()
	return scheme === 'http' || scheme === 'https'
}

/** Fetch a font file into raw bytes for `registerFontMetrics`. */
export async function fetchFontBytes(source: string): Promise<Uint8Array> {
	const response = await fetch(source)
	if (!response.ok) throw new MediaError('font/fetch-failed', `Unable to load font (fetch): ${source}`)
	return new Uint8Array(await response.arrayBuffer())
}

/**
 * Fetch a media file as raw base64 (no `data:` prefix) — see {@link RuntimeAdapter.loadMedia}
 * for why both encodings are accepted.
 *
 * Encodes through `bytesToBase64` rather than `Buffer`, so the same function serves a runtime
 * that has no `Buffer`. The Node adapter still uses `Buffer` on its *filesystem* branch, where
 * it is both available and faster on a large payload.
 */
export async function fetchMediaBase64(path: string): Promise<string> {
	const response = await fetch(path)
	if (!response.ok) throw new MediaError('media/fetch-failed', `Unable to load image (fetch): ${path}`)
	return bytesToBase64(new Uint8Array(await response.arrayBuffer()))
}

/**
 * The SVG-preview fallback for a runtime with no canvas: stamp the placeholder PNG.
 *
 * An SVG placed on a slide needs a raster fallback beside it for viewers that will not render
 * the SVG. Rasterizing needs a canvas, which only the browser adapter has, so everywhere else
 * the fallback is a fixed placeholder image rather than a missing part.
 */
export async function placeholderSvgPreview(rel: SlideRelMedia): Promise<string> {
	rel.data = IMG_SVG_PLACEHOLDER
	return 'done'
}
