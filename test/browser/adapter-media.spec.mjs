import fs from 'node:fs'
import path from 'node:path'
import { diffParts, explodePackage, listParts } from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'
import { expect, test } from './fixtures.mjs'
import { buildDeckInHarness, buildDeckInNode, NODE_ASSETS, openHarness, packageBytes } from './helpers.mjs'

/**
 * `loadMedia` and `createSvgPngPreview` (src/runtime/browser.ts) in a real browser.
 *
 * These two functions are why the demo fixture is not enough: `quarterly-review` draws
 * every asset it shows, so no deck in the demo ever asks the runtime to load one. Here
 * the harness loads the shipped `dist/browser.js` unbundled and builds decks whose only
 * purpose is to reach them — from real URLs, with real 404s.
 *
 * Two of the three assertions below are cross-runtime, using the byte-identity gate's own
 * explode/normalize/diff (`scripts/pptx-parts.mjs`) so there is one definition of "the
 * same bytes" in the repo. The raster case must come out identical; the SVG case must
 * not, and the test pins exactly how it differs — Node has no rasterizer, so its
 * `createSvgPngPreview` writes a fixed placeholder where the browser writes a real PNG.
 * That is a documented divergence, and the way to keep it documented is to assert it.
 */

const OUT_ROOT = path.join(ROOT, '.tmp', 'browser-adapter')

/** Explode both packages of one deck into `.tmp/` and hand back the two directories. */
async function bothRuntimes(page, deck) {
	const browser = await buildDeckInHarness(page, deck)
	expect(browser.ok, `the harness failed to build the "${deck}" deck: ${browser.message}`).toBe(true)
	const nodeBase64 = await buildDeckInNode(deck)

	// Left on disk deliberately: on a failure the two exploded trees are what makes a
	// named part diffable, exactly as `.tmp/byte-identity/` is for the Node gate.
	return {
		nodeDir: await explodePackage(packageBytes(nodeBase64), path.join(OUT_ROOT, deck, 'node')),
		browserDir: await explodePackage(packageBytes(browser.base64), path.join(OUT_ROOT, deck, 'browser')),
	}
}

/** The media parts of an exploded package, sorted. */
const mediaParts = (dir) => listParts(dir).filter((part) => part.startsWith('ppt/media/'))

test.beforeEach(async ({ page }) => {
	await openHarness(page)
})

test('loadMedia: a raster image fetched in the browser lands as the same bytes Node reads off disk', async ({
	page,
}) => {
	const { nodeDir, browserDir } = await bothRuntimes(page, 'raster')

	// The two implementations return *different strings* for the same image — Node hands
	// back raw base64, the browser a `FileReader` data URI with a `Content-Type`-derived
	// mime prefix. Everything downstream of the adapter has to reconcile that, including
	// the image sizer: the deck pins no width or height, so the emitted extent is measured
	// from whatever the loader returned. If either reconciliation were wrong this diff
	// would name `ppt/media/…` or `ppt/slides/slide1.xml`.
	const diffs = diffParts(nodeDir, browserDir)
	expect(diffs, `node vs browser raster deck differs:\n  ${diffs.join('\n  ')}`).toEqual([])

	// Agreement between two runtimes is not the same as agreement with the file. Compare
	// the embedded part against the source bytes so the test cannot pass on two identical
	// wrong answers.
	expect(mediaParts(browserDir)).toHaveLength(1)
	const embedded = fs.readFileSync(path.join(browserDir, mediaParts(browserDir)[0]))
	expect(embedded.equals(fs.readFileSync(NODE_ASSETS.png))).toBe(true)
})

test('createSvgPngPreview: the browser rasterizes the PNG fallback Node can only stub', async ({ page }) => {
	const { nodeDir, browserDir } = await bothRuntimes(page, 'svg')

	// An SVG consumes two rels: the SVG itself and a PNG fallback. The SVG part must
	// match — it is the same bytes through the same loader — and the PNG part must not.
	const diffs = diffParts(nodeDir, browserDir)
	expect(diffs).toHaveLength(1)
	expect(diffs[0]).toMatch(/^CHANGED {2}ppt\/media\/.*\.png$/)

	const pngPart = mediaParts(browserDir).find((part) => part.endsWith('.png'))
	const browserPng = fs.readFileSync(path.join(browserDir, pngPart))
	const nodePng = fs.readFileSync(path.join(nodeDir, pngPart))

	// A real PNG, by its signature, and materially bigger than the placeholder — i.e. the
	// canvas actually drew the 500×500 artwork rather than handing back an empty bitmap.
	expect([...browserPng.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47])
	expect(browserPng.byteLength).toBeGreaterThan(nodePng.byteLength * 4)

	// And the SVG itself is untouched by either runtime.
	const svgPart = mediaParts(browserDir).find((part) => part.endsWith('.svg'))
	expect(fs.readFileSync(path.join(browserDir, svgPart)).equals(fs.readFileSync(NODE_ASSETS.svg))).toBe(true)
})

test('loadMedia: a 404 fails the export with the adapter code, not a bare fetch error', async ({ page }) => {
	const outcome = await buildDeckInHarness(page, 'missingImage')
	expect(outcome.ok).toBe(false)
	// The pipeline wraps the loader failure so the message names the asset; the adapter's
	// own code survives as the cause. Both halves are the API, so both are asserted.
	expect(outcome.code).toBe('media/load-failed')
	expect(outcome.causeCode).toBe('media/fetch-failed')
	expect(outcome.message).toContain('no-such-image.png')
})

test('createSvgPngPreview: an undecodable SVG fails rather than shipping a blank fallback', async ({ page }) => {
	const outcome = await buildDeckInHarness(page, 'brokenSvg')
	expect(outcome.ok).toBe(false)
	expect(outcome.code).toBe('media/load-failed')
	// `image.onerror` — the fetch succeeded, so this is the decode arm, not the network one.
	expect(outcome.causeCode).toBe('media/svg-preview-failed')
})

test('createSvgPngPreview: a zero-dimension SVG is caught by the h/w guard', async ({ page }) => {
	const outcome = await buildDeckInHarness(page, 'zeroSizeSvg')
	expect(outcome.ok).toBe(false)
	expect(outcome.causeCode).toBe('media/svg-preview-failed')
	// Distinguishes this arm from the `onerror` one above: same code, different reason,
	// and the reason is what says the guard fired instead of the decode failing.
	expect(outcome.causeMessage).toContain('h/w=0')
})
