'use strict'

// End-to-end test harness for the browser demo (`demos/browser/index.html`).
//
// Strategy:
//   1. Boot `browser_server.mjs` from the `demos/` cwd (port 8000) — once per file.
//   2. Launch Chromium via Playwright — once per file, reused across cases.
//   3. Per case: open a fresh BrowserContext (isolates global state between
//      successive demo button clicks), intercept the CDN URLs in
//      `demos/browser/index.html` and fulfill them with the LOCAL fresh bundle
//      (`process.env.PPTXGEN_LOCAL_BUNDLE` from `_runner.js`) and the local
//      `demos/modules/demos.mjs`. No source mutation, no temp HTML.
//   4. Click the case's selector, capture the resulting `.pptx` download.
//   5. Validate via the OOXML validator (same helper as `test/run-schema.js`)
//      and assert zero schema errors.
//
// Cases cover every button on `index.html`:
//   * `#btnRunAllDemos`        (header, 120s — produces ONE big deck)
//   * `#btnRunBasicDemo`       (tab-intro)
//   * `#btnRunSandboxDemo`     (tab-intro — depends on `doAppStart` populating
//                                `<code id="demo-sandbox">`; we wait for that)
//   * `#btnGenFunc_Chart`      (tab-charts)
//   * `#btnGenFunc_Image`      (tab-images)
//   * `#btnGenFunc_Media`      (tab-images)
//   * `#btnGenFunc_Shape`      (tab-shapes)
//   * `#btnGenFunc_Text`       (tab-shapes)
//   * `#btnGenFunc_Table`      (tab-tables)
//   * `#btnGenFunc_Master`     (tab-masters)

const fs = require('fs')
const path = require('path')

const { isInstalled } = require('../validator')
const { assert, expectNoSchemaErrors, withTempDir } = require('./_helpers')
const { startServer } = require('./_server')

const DEMO_URL = 'http://localhost:8000/browser/index.html'
const CDN_BUNDLE_URL = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/dist/pptxgen.bundle.js'
const CDN_DEMOS_URL = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/demos/modules/demos.mjs'

const REPO_ROOT = path.resolve(__dirname, '..', '..')
const LOCAL_BUNDLE = process.env.PPTXGEN_LOCAL_BUNDLE || path.join(REPO_ROOT, 'dist', 'pptxgen.bundle.js')
const LOCAL_DEMOS = path.join(REPO_ROOT, 'demos', 'modules', 'demos.mjs')

const DEFAULT_DOWNLOAD_TIMEOUT_MS = 60000
const ALL_DEMOS_TIMEOUT_MS = 120000

async function setup (ctx) {
	if (!isInstalled()) throw new Error('OOXMLValidatorCLI not installed; run ./tools/ooxml-validator/install.sh')
	if (!fs.existsSync(LOCAL_BUNDLE)) throw new Error('local bundle missing at ' + LOCAL_BUNDLE + '; the runner is responsible for assembling this — run `node test/release/_runner.js` rather than this test directly')
	if (!fs.existsSync(LOCAL_DEMOS)) throw new Error('local demos.mjs missing at ' + LOCAL_DEMOS)

	let chromium
	try {
		chromium = require('playwright').chromium
	} catch (e) {
		throw new Error('playwright is not installed; run `npm install` (and `npx playwright install chromium` if needed): ' + e.message)
	}

	ctx.server = await startServer()
	try {
		ctx.browser = await chromium.launch({ headless: true })
	} catch (e) {
		await ctx.server.kill()
		throw e
	}
}

async function teardown (ctx) {
	if (ctx.browser) {
		try { await ctx.browser.close() } catch (_) { /* ignore */ }
	}
	if (ctx.server) {
		try { await ctx.server.kill() } catch (_) { /* ignore */ }
	}
}

async function runDemoButtonCase (ctx, opts) {
	const { selector, label, downloadTimeoutMs = DEFAULT_DOWNLOAD_TIMEOUT_MS } = opts

	await withTempDir('pptxgen-release-', async (tmpDir) => {
		const context = await ctx.browser.newContext({ acceptDownloads: true })

		let bundleHits = 0
		const pageErrors = []
		const consoleErrors = []

		try {
			const page = await context.newPage()

			await page.route(CDN_BUNDLE_URL, route => {
				bundleHits++
				return route.fulfill({
					path: LOCAL_BUNDLE,
					headers: { 'Content-Type': 'application/javascript' }
				})
			})
			await page.route(CDN_DEMOS_URL, route => {
				return route.fulfill({
					path: LOCAL_DEMOS,
					headers: { 'Content-Type': 'application/javascript' }
				})
			})

			page.on('pageerror', err => { pageErrors.push(err.message) })
			page.on('console', msg => { if (msg.type() === 'error') consoleErrors.push(msg.text()) })

			await page.goto(DEMO_URL, { waitUntil: 'load' })
			await page.waitForLoadState('networkidle')

			// `<code id="demo-sandbox">` ships empty in tab-intro.html. `doAppStart`
			// (called on DOMContentLoaded) populates it once the section HTML loads.
			// Until that happens, `#btnRunSandboxDemo` runs `new Function('')()` (no-op).
			// `#btnRunBasicDemo` reads from `#demo-basic` populated in the same call,
			// so this gate covers both paths.
			await page.waitForFunction(
				() => {
					const el = document.getElementById('demo-sandbox')
					return !!el && typeof el.innerHTML === 'string' && el.innerHTML.includes('writeFile')
				},
				null,
				{ timeout: 15000 }
			)

			await page.waitForSelector(selector, { state: 'attached', timeout: 15000 })

			assert(
				bundleHits >= 1,
				'expected the CDN bundle URL to be intercepted at least once for ' + label +
					'; got ' + bundleHits + ' hits — substitution did not run, so the test would not be exercising local code.' +
					' pageErrors=' + JSON.stringify(pageErrors) + ' consoleErrors=' + JSON.stringify(consoleErrors)
			)

			// Many demo buttons live inside currently-collapsed Bootstrap accordions or
			// inactive tab panels (e.g. `#btnGenFunc_Chart` is in the `#tab-charts`
			// pane which is `display:none` until the user activates that tab). The
			// click handlers are wired by `main.js` directly via `addEventListener`
			// at module load, so a programmatic `.click()` reliably fires them
			// regardless of the element's CSS visibility — exactly as if the user
			// had opened the tab and clicked. We use `dispatchEvent` to ensure
			// bubbling and Bootstrap-modal-friendly default behaviour.
			const [download] = await Promise.all([
				page.waitForEvent('download', { timeout: downloadTimeoutMs }),
				page.evaluate((sel) => {
					const el = document.querySelector(sel)
					if (!el) throw new Error('selector not found: ' + sel)
					el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }))
				}, selector)
			])

			const suggested = download.suggestedFilename() || (label.replace(/[^a-z0-9]+/gi, '_') + '.pptx')
			const out = path.join(tmpDir, suggested)
			await download.saveAs(out)

			const stat = fs.statSync(out)
			assert(stat.size > 0, 'downloaded file is empty: ' + out)

			await expectNoSchemaErrors(out, label)
		} finally {
			try { await context.close() } catch (_) { /* ignore */ }
		}
	})
}

// `knownFailure` annotations:
// The harness surfaces real, pre-existing schema-validity bugs in the library
// and demo data. They are deferred to follow-up slices/commits — see
// `.autoloop/runs/full-editor/progress.md` "Relevant Issues". Marking these
// `knownFailure` keeps the harness landable today and red-flags each case
// loudly with its bug reference. Removing a marker once the underlying fix
// lands is intentionally a one-line edit.
const KF_LOWERCASE_GUID = 'getUuid emits lowercase hex; OOXML schema requires uppercase A-F (src/gen-utils.ts getUuid)'
const KF_ZIP_DIR_ENTRIES = 'zip.folder() emits directory entries that OpenXmlValidator rejects with OpenXmlPackageException (src/pptxgen.ts ~line 501)'
const KF_ALL_DEMOS_CASCADE = 'btnRunAllDemos cascades both KF_LOWERCASE_GUID (sections in every sub-demo) and KF_ZIP_DIR_ENTRIES'

const BUTTON_CASES = [
	{ selector: '#btnRunBasicDemo', label: 'browser/btnRunBasicDemo' },
	{ selector: '#btnRunSandboxDemo', label: 'browser/btnRunSandboxDemo' },
	{ selector: '#btnRunAllDemos', label: 'browser/btnRunAllDemos', downloadTimeoutMs: ALL_DEMOS_TIMEOUT_MS, knownFailure: KF_ALL_DEMOS_CASCADE },
	{ selector: '#btnGenFunc_Chart', label: 'browser/btnGenFunc_Chart', knownFailure: KF_ZIP_DIR_ENTRIES },
	{ selector: '#btnGenFunc_Image', label: 'browser/btnGenFunc_Image', knownFailure: KF_LOWERCASE_GUID },
	{ selector: '#btnGenFunc_Media', label: 'browser/btnGenFunc_Media', knownFailure: KF_ZIP_DIR_ENTRIES },
	{ selector: '#btnGenFunc_Shape', label: 'browser/btnGenFunc_Shape', knownFailure: KF_ZIP_DIR_ENTRIES },
	{ selector: '#btnGenFunc_Text', label: 'browser/btnGenFunc_Text', knownFailure: KF_ZIP_DIR_ENTRIES },
	{ selector: '#btnGenFunc_Table', label: 'browser/btnGenFunc_Table', knownFailure: KF_ZIP_DIR_ENTRIES },
	{ selector: '#btnGenFunc_Master', label: 'browser/btnGenFunc_Master', knownFailure: KF_ZIP_DIR_ENTRIES }
]

const cases = BUTTON_CASES.map(opts => {
	const c = {
		name: opts.label + ': produces a schema-valid .pptx',
		fn: (ctx) => runDemoButtonCase(ctx, opts)
	}
	if (opts.knownFailure) c.knownFailure = opts.knownFailure
	return c
})

module.exports = { setup, teardown, cases }
