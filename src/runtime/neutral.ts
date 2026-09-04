import { UnsupportedFeatureError } from '../errors.js'
import type { RuntimeAdapter } from './types.js'
import { fetchFontBytes, fetchMediaBase64, placeholderSvgPreview } from './fetch-media.js'

/**
 * The adapter for a runtime that is neither Node nor a browser — Deno, Bun, an edge worker,
 * anything that imports the bare `pptx-ts` specifier without resolving the `node` or
 * `browser` export condition.
 *
 * It implements everything that is genuinely host-neutral and refuses only what is not.
 * `fetch`, `btoa` and `TextEncoder` exist in all of these runtimes, so remote media and fonts
 * load here exactly as they do in a browser — which is why every method but `writeFile` is now
 * the shared implementation in {@link ./fetch-media} rather than a copy of it. What is missing
 * is a *destination*: there is no `node:fs` to write to and no `document` to hand a download to,
 * so {@link RuntimeAdapter.writeFile} throws a named error rather than guessing. Everything that
 * returns bytes to the caller — `write`, `stream`, `toParts` — works normally, which is the
 * shape that suits a worker anyway.
 *
 * Media loading is fetch-only here: a bare filesystem path has no meaning without a filesystem,
 * and `fetch` rejecting it is reported by the caller as a load failure naming the path.
 */
export function createNeutralRuntime(): RuntimeAdapter {
	return {
		writeFileOutputType: null,
		loadMedia: (rel) => fetchMediaBase64(rel.path),
		createSvgPngPreview: placeholderSvgPreview,
		writeFile,
		loadFontData: fetchFontBytes,
	}
}

async function writeFile(fileName: string): Promise<string> {
	throw new UnsupportedFeatureError(
		'runtime/file-output-unavailable',
		`Cannot write "${fileName}": this runtime has neither a filesystem nor a DOM. ` +
			'Import "pptx-ts/node" to write to disk or "pptx-ts/browser" to trigger a ' +
			'download; anywhere else, use write() and persist the bytes yourself.'
	)
}
