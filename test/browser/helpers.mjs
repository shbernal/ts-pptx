import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { expect } from '@playwright/test'
import { ROOT } from '../../scripts/script-utils.mjs'
import { buildDeckBase64 } from './harness/decks.mjs'

/**
 * Drive `demos/vite-demo` through one deck build and hand back the downloaded bytes.
 *
 * This is the whole point of the browser lane: the demo imports the *same* showcase
 * module `pnpm demos:build quarterly-review` runs, so what comes back here is the deck
 * as a browser assembled it — through `src/runtime/browser.ts`'s `writeFile`, the
 * object-URL `<a download>` path that no Node test can reach.
 */
export async function buildDeckInBrowser(page) {
	await page.goto('./')

	const button = page.getByRole('button', { name: /^Build / })
	await expect(button).toBeEnabled()

	const downloadPromise = page.waitForEvent('download')
	await button.click()

	// Wait on the page's own outcome before the download, so a build that threw is
	// reported as its error message rather than as a 30s "no download event" timeout.
	// `App.tsx` renders `role="status"` on success and `role="alert"` on failure.
	const outcome = page.locator('[role="status"], [role="alert"]')
	await expect(outcome).toBeVisible({ timeout: 30_000 })
	const alert = page.getByRole('alert')
	if (await alert.count()) throw new Error('the demo failed to build the deck: ' + (await alert.innerText()))

	const download = await downloadPromise
	const file = await download.path()
	return { bytes: new Uint8Array(await readFile(file)), fileName: download.suggestedFilename() }
}

// --- the runtime-adapter harness (the `runtime-adapter` Playwright project) ---

/**
 * The same media and font sources the harness page uses, as filesystem paths.
 *
 * Keep in step with `ASSETS` in `harness/harness.mjs`: same files, addressed the way each
 * runtime addresses them. That substitution — a URL for a path — is the only intended
 * difference between the two sides of every cross-runtime comparison below.
 */
export const NODE_ASSETS = {
	png: path.join(ROOT, 'demos', 'common', 'images', 'logo_square.png'),
	svg: path.join(ROOT, 'demos', 'common', 'images', 'lock-green.svg'),
	font: path.join(ROOT, 'test', 'read', 'fixtures', 'fonts', 'Silkscreen-Regular.ttf'),
	missingPng: path.join(ROOT, 'demos', 'common', 'images', 'no-such-image.png'),
	missingFont: path.join(ROOT, 'test', 'read', 'fixtures', 'fonts', 'no-such-font.ttf'),
	brokenSvg: path.join(ROOT, 'test', 'browser', 'harness', 'broken.svg'),
	zeroSizeSvg: path.join(ROOT, 'test', 'browser', 'harness', 'zero-size.svg'),
}

/** Load the harness page and fail with the page's own reason if it did not come up. */
export async function openHarness(page) {
	await page.goto('./')
	await page.waitForFunction(() => !!window['harness'] || !!window['harnessError'])
	const failure = await page.evaluate(() => window['harnessError'])
	// The likeliest cause by far is `dist/browser.js` (or a chunk it reaches) acquiring an
	// import the browser cannot resolve — a bare specifier missing from the page's import
	// map, or a `node:*` builtin. Both are findings about the shipped package, so they are
	// reported as the error the browser gave rather than as a missing global.
	if (failure) throw new Error('the adapter harness failed to load: ' + failure)
}

/**
 * Build one deck from `harness/decks.mjs` in the page.
 * @returns the harness's flattened outcome — `{ok:true, base64}` or `{ok:false, code, …}`.
 */
export async function buildDeckInHarness(page, deck) {
	return await page.evaluate((name) => window['harness'].build(name), deck)
}

/** Build the same deck in Node, against `dist/node.js`, for the comparison. */
export async function buildDeckInNode(deck) {
	const { default: TsPptx } = await import('../../dist/node.js')
	return await buildDeckBase64(new TsPptx(), deck, NODE_ASSETS)
}

/** Decode a package the harness returned as base64. */
export function packageBytes(base64) {
	return new Uint8Array(Buffer.from(base64, 'base64'))
}

// --- the rendered-table harness (the `html-table` Playwright project) ---

/**
 * Load the rendered-table page, failing with the page's own reason if it did not come up.
 *
 * A separate page from `openHarness`'s, on the same server: it renders a real `<table>` so
 * `tableToSlides` reads a non-zero `offsetWidth`. See `harness/table.mjs` for why the two
 * fixtures are not one.
 */
export async function openTableHarness(page) {
	await page.goto('./table.html')
	await page.waitForFunction(() => !!window['tableHarness'] || !!window['harnessError'])
	const failure = await page.evaluate(() => window['harnessError'])
	if (failure) throw new Error('the rendered-table harness failed to load: ' + failure)
}

/**
 * The two width bases the live page reports for one fixture.
 * @returns {Promise<{measured: number[], css: string[]}>}
 */
export async function tableBases(page, scenario) {
	return await page.evaluate((name) => window['tableHarness'].bases(name), scenario)
}

/**
 * Convert one fixture table in the page.
 * @returns the harness's flattened outcome — `{ok:true, base64}` or `{ok:false, code, message}`.
 */
export async function buildTableInHarness(page, scenario) {
	return await page.evaluate((name) => window['tableHarness'].build(name), scenario)
}

/**
 * Convert the same fixture in Node, against a DOM that renders nothing.
 *
 * happy-dom is the same DOM `test/regression/html/html-to-slides-node.test.js` drives, and the
 * point of building here too is that `offsetWidth` is `0` for every cell — so the widths
 * come from the *other* basis. That contrast is the assertion, not an incidental detail.
 */
export async function buildTableInNode(scenario) {
	const { Window } = await import('happy-dom')
	const { tableToSlides } = await import('../../dist/html.js')
	const { default: TsPptx } = await import('../../dist/node.js')
	const { TABLE_HTML, TABLE_ID } = await import('./harness/table-fixture.mjs')

	const win = new Window()
	win.document.body.innerHTML = TABLE_HTML[scenario]
	const pres = new TsPptx()
	tableToSlides(pres, win.document.getElementById(TABLE_ID))
	return /** @type {string} */ (await pres.write({ outputType: 'base64' }))
}
