import { DECKS } from './harness/decks.mjs'
import { expect, test } from './fixtures.mjs'
import { buildDeckInHarness, openHarness } from './helpers.mjs'

/**
 * The browser lane's own coverage gate.
 *
 * This is the per-function half of the lane's coverage story, and it stays that way now
 * that the other half exists. `scripts/coverage-merge.mjs` folds what every spec here
 * executed into the Node report, so the repo's headline numbers finally count
 * `src/runtime/browser.ts` where it runs — but a percentage cannot say "every adapter
 * function was entered", and a merged report that drifts by a tenth says nothing about
 * *which* function stopped running. That is this file's job.
 *
 * The measurement is the one available where the code actually runs: Chromium's own V8
 * block coverage for the `dist/browser.js` script, collected across every scenario the
 * harness has. Two assertions, and they fail for different reasons —
 *
 *   - **every adapter function ran.** Add a fifth function to `RuntimeAdapter` and this
 *     stays green until it is added to the list below, which is the moment to ask what
 *     covers it. Delete the only test for an existing one and it goes red immediately.
 *   - **the file's executed share stays above a floor.** Catches the other shape of
 *     regression: a function that is still entered, but whose interesting arms are not.
 *
 * Chromium-only by construction (`page.coverage` is a CDP feature). That is not a
 * constraint the lane feels, because the lane is Chromium-only on purpose — see
 * docs/runtime-and-package-support.md "Which Browsers The Lane Runs".
 */

/**
 * Every function in `src/runtime/browser.ts`, by the name it keeps in the bundle.
 *
 * `writeFile` is why the harness has a `download` scenario at all: the demo drives that
 * path too, but through a Vite bundle, where this file is not loaded as a file and no
 * per-URL coverage entry exists for it.
 */
const ADAPTER_FUNCTIONS = ['createBrowserRuntime', 'loadFontData', 'loadMedia', 'createSvgPngPreview', 'writeFile']

/**
 * Floor for the executed share of `dist/browser.js`. Measured 92.74 when this landed;
 * pinned a notch below with a couple of points of slack, like every other gate in the
 * repo. Ratchet it upward as coverage improves — never down to make a red run green.
 *
 * What keeps it off 100: `createSvgPngPreview`'s missing-2d-context arm (unreachable in a
 * browser that has a canvas) and `loadMedia`'s `FileReader.onerror` (a `Blob` read that
 * does not fail). Reaching either means stubbing DOM constructors, which asserts about the
 * stub rather than about the browser.
 *
 * `tableToSlides` used to be on that list too, described as out of scope because it needs a
 * rendered `<table>`. It is now covered — by the `html-table` project, against its own page
 * (test/browser/harness/table.html). It does not move this number and should not: the
 * measurement here is per-page V8 coverage of the *adapter* harness, whose DOM has no table
 * in it, so `tableToSlides` is simply not reachable from these scenarios. Its coverage
 * reaches the repo's headline figures the same way everything else in the lane does, through
 * `scripts/coverage-merge.mjs`. The scope line has not moved: what `html-table` proves is
 * that a measurement is taken and honoured, not that the browser's layout is correct — see
 * test/browser/table-widths.spec.mjs.
 */
const MIN_EXECUTED_PCT = 90

/** V8 reports uncovered spans as ranges with `count === 0`; everything else ran. */
function executedPct(entry) {
	const uncovered = entry.functions.flatMap((fn) => fn.ranges.filter((range) => range.count === 0))
	// Ranges nest (a function's body encloses its own dead arms), so union them by
	// walking sorted starts rather than summing — otherwise a dead arm inside a dead
	// function is counted twice and the percentage undershoots.
	let unreached = 0
	let cursor = 0
	for (const range of uncovered.sort((a, b) => a.startOffset - b.startOffset)) {
		const from = Math.max(range.startOffset, cursor)
		if (range.endOffset > from) {
			unreached += range.endOffset - from
			cursor = range.endOffset
		}
	}
	const total = entry.source.length
	return ((total - unreached) / total) * 100
}

test('every RuntimeAdapter function runs, and dist/browser.js stays above its coverage floor', async ({
	page,
	jsCoverage,
}) => {
	// `jsCoverage` (test/browser/fixtures.mjs) started collection before this body ran, so
	// the harness page's own load is measured rather than being the one thing that is not.
	await openHarness(page)

	// Every scenario the harness knows, so the gate cannot silently stop covering one
	// that a sibling spec drops. Failures are outcomes here, not throws — the specs next
	// door assert *which* failure each one is; this only needs the code to have run.
	for (const deck of Object.keys(DECKS)) await buildDeckInHarness(page, deck)

	const downloadPromise = page.waitForEvent('download')
	await page.evaluate(() => window['harness'].download('fonts'))
	await downloadPromise

	// The fixture caches this, so its own teardown writes these same entries out for the
	// merge instead of collecting a second, empty round.
	const entries = await jsCoverage.stop()
	const entry = entries.find((script) => script.url.endsWith('/dist/browser.js'))
	expect(entry, `no coverage entry for dist/browser.js; got:\n  ${entries.map((e) => e.url).join('\n  ')}`).toBeTruthy()

	const ran = new Set(
		entry.functions.filter((fn) => fn.ranges.some((range) => range.count > 0)).map((fn) => fn.functionName)
	)
	expect(ADAPTER_FUNCTIONS.filter((name) => !ran.has(name))).toEqual([])

	const pct = executedPct(entry)
	expect(pct, `dist/browser.js executed ${pct.toFixed(2)}% (floor ${MIN_EXECUTED_PCT}%)`).toBeGreaterThanOrEqual(
		MIN_EXECUTED_PCT
	)
})
