'use strict'

// End-to-end smoke test for the browser demo (`demos/browser/index.html`).
//
// Strategy:
//   1. Boot `browser_server.mjs` from the `demos/` cwd (port 8000).
//   2. Launch Chromium via Playwright.
//   3. Intercept the CDN URLs in `demos/browser/index.html` and fulfill them
//      with the LOCAL `dist/pptxgen.bundle.js` and `demos/modules/demos.mjs`
//      so the test exercises the local code path, not the published library.
//      No source mutation, no temp HTML.
//   4. Click `#btnRunBasicDemo` and capture the resulting `.pptx` download.
//   5. Validate the file via the OOXML validator (same helper used by
//      `test/run-schema.js`) and assert zero schema errors.
//
// This is the "smoke" slice (S3a). The full button matrix lands in S3b.

const fs = require('fs')
const os = require('os')
const path = require('path')

const { runValidatorOnFile, isInstalled } = require('../validator')
const { startServer } = require('./_server')

const DEMO_URL = 'http://localhost:8000/browser/index.html'
const CDN_BUNDLE_URL = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/dist/pptxgen.bundle.js'
const CDN_DEMOS_URL = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/demos/modules/demos.mjs'

const REPO_ROOT = path.resolve(__dirname, '..', '..')
const LOCAL_BUNDLE = process.env.PPTXGEN_LOCAL_BUNDLE || path.join(REPO_ROOT, 'dist', 'pptxgen.bundle.js')
const LOCAL_DEMOS = path.join(REPO_ROOT, 'demos', 'modules', 'demos.mjs')

function assert (cond, msg) {
	if (!cond) throw new Error('assertion failed: ' + msg)
}

async function expectNoSchemaErrors (filePath, label) {
	const errors = await runValidatorOnFile(filePath)
	if (errors.length === 0) return
	const summary = errors
		.slice(0, 5)
		.map(e => `  - [${e.ErrorType}] ${e.Description} (path: ${(e.Path && e.Path.PartUri) || '?'})`)
		.join('\n')
	const more = errors.length > 5 ? `\n  ...(${errors.length - 5} more)` : ''
	assert(false, `${label}: ${errors.length} schema error(s):\n${summary}${more}`)
}

async function withBrowser (fn) {
	// Lazy require so this file can be loaded by the runner even if Playwright
	// is unavailable (the test will then fail with a clear message).
	let chromium
	try {
		chromium = require('playwright').chromium
	} catch (e) {
		throw new Error('playwright is not installed; run `npm install` (and `npx playwright install chromium` if needed): ' + e.message)
	}
	const browser = await chromium.launch({ headless: true })
	try {
		return await fn(browser)
	} finally {
		await browser.close()
	}
}

async function runBasicDemoCase () {
	if (!isInstalled()) throw new Error('OOXMLValidatorCLI not installed; run ./tools/ooxml-validator/install.sh')
	if (!fs.existsSync(LOCAL_BUNDLE)) throw new Error('local bundle missing at ' + LOCAL_BUNDLE + '; the runner is responsible for assembling this — run `node test/release/_runner.js` rather than this test directly')
	if (!fs.existsSync(LOCAL_DEMOS)) throw new Error('local demos.mjs missing at ' + LOCAL_DEMOS)

	const server = await startServer()
	const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptxgen-release-'))

	try {
		await withBrowser(async (browser) => {
			const context = await browser.newContext({ acceptDownloads: true })
			let bundleHits = 0
			let demosHits = 0
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
					demosHits++
					return route.fulfill({
						path: LOCAL_DEMOS,
						headers: { 'Content-Type': 'application/javascript' }
					})
				})

				// Surface page errors in the test output for easier debugging.
				const pageErrors = []
				page.on('pageerror', err => { pageErrors.push(err.message) })

				await page.goto(DEMO_URL, { waitUntil: 'load' })
				await page.waitForLoadState('networkidle')
				await page.waitForSelector('#btnRunBasicDemo', { state: 'visible', timeout: 15000 })

				assert(bundleHits >= 1, 'expected the CDN bundle URL to be intercepted at least once; got ' + bundleHits + ' hits — substitution did not run, so the test would not be exercising local code. pageerrors=' + JSON.stringify(pageErrors))

				const [download] = await Promise.all([
					page.waitForEvent('download', { timeout: 60000 }),
					page.click('#btnRunBasicDemo')
				])

				const suggested = download.suggestedFilename() || 'basic-demo.pptx'
				const out = path.join(tmpDir, suggested)
				await download.saveAs(out)

				const stat = fs.statSync(out)
				assert(stat.size > 0, 'downloaded file is empty: ' + out)

				await expectNoSchemaErrors(out, 'browser/btnRunBasicDemo')
			} finally {
				await context.close()
			}
		})
	} finally {
		try { fs.rmSync(tmpDir, { recursive: true, force: true }) } catch (_) { /* ignore */ }
		await server.kill()
	}
}

module.exports = [
	{
		name: 'browser/btnRunBasicDemo: produces a schema-valid .pptx',
		fn: runBasicDemoCase
	}
]
