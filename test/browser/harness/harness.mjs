// The browser side of the adapter lane: loads the *shipped* `dist/browser.js` over a
// plain `<script type="module">` — no bundler anywhere in the path — and exposes one
// function for the specs to drive through `page.evaluate`.
//
// Loading dist directly is deliberate. The site's demos page already covers the bundled
// story (Rollup resolving the `browser` export condition and tree-shaking it); what it
// cannot tell you is whether the file that ships is itself browser-loadable. Here, a
// single `node:*` import anywhere in the `dist/browser.js` chunk graph is a hard page
// failure rather than something a bundler quietly polyfills.
//
// The relative specifier below resolves because scripts/browser-harness-server.mjs
// serves the repo with its real layout, so the URL depth and the on-disk depth agree —
// which is also what lets `pnpm run typecheck:test` resolve it to `dist/browser.d.ts`.
import TsPptx from '../../../dist/browser.js'
import { buildDeckBase64, DECKS } from './decks.mjs'

/**
 * Media and font sources as URLs. The Node side of each comparison passes the same files
 * as filesystem paths; that substitution is the entire difference under test.
 *
 * The two `missing*` entries point *inside* a served prefix, so what comes back is a real
 * 404 from a real server rather than a stubbed `fetch` — the adapter's `!response.ok`
 * arms are reached the way a consumer would reach them.
 */
const ASSETS = {
	png: '/demos/common/images/logo_square.png',
	svg: '/demos/common/images/lock-green.svg',
	font: '/test/read/fixtures/fonts/Silkscreen-Regular.ttf',
	missingPng: '/demos/common/images/no-such-image.png',
	missingFont: '/test/read/fixtures/fonts/no-such-font.ttf',
	brokenSvg: '/test/browser/harness/broken.svg',
	zeroSizeSvg: '/test/browser/harness/zero-size.svg',
}

/**
 * Build a deck and report the outcome as plain data.
 *
 * Errors are flattened rather than thrown because a rejection crossing the
 * `page.evaluate` boundary arrives as a bare message: the class and its `code` — which
 * are the API, and the thing worth asserting — do not survive. Failure paths are
 * therefore a return value here, and the spec asserts on `code`.
 *
 * @param {string} name a key of `DECKS` in ./decks.mjs
 * @returns {Promise<{ok: true, base64: string} | {ok: false, name: string, code: string, message: string, causeCode: string, causeMessage: string}>}
 */
async function build(name) {
	try {
		return { ok: true, base64: await buildDeckBase64(new TsPptx(), name, ASSETS) }
	} catch (err) {
		// The media pipeline wraps a loader failure in `media/load-failed` and chains the
		// original as `cause`, so the adapter's own code is one level down. Both are
		// reported: the outer says which stage failed, the inner says which loader.
		const cause = err?.cause
		return {
			ok: false,
			name: String(err?.name ?? ''),
			code: String(err?.code ?? ''),
			message: String(err?.message ?? err),
			causeCode: String(cause?.code ?? ''),
			causeMessage: String(cause?.message ?? ''),
		}
	}
}

/**
 * Build a deck and hand it to `writeFile` — the object-URL `<a download>` path.
 *
 * The site's demos page already drives that path, but through a Vite bundle. Doing it here
 * too is what lets one coverage measurement over `dist/browser.js` account for all four
 * adapter functions (see adapter-coverage.spec.mjs); without it `writeFile` would be
 * covered in a fixture where the file is not loaded as a file at all.
 *
 * @param {string} name a key of `DECKS` in ./decks.mjs
 * @returns {Promise<string>} the file name `writeFile` reports
 */
async function download(name) {
	const pres = new TsPptx()
	await DECKS[name](pres, ASSETS)
	return await pres.writeFile({ fileName: 'harness.pptx' })
}

Object.assign(window, { harness: { assets: ASSETS, build, download } })
