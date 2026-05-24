'use strict'

// End-to-end test for the Web Worker demo (`demos/browser/worker_test.html`).
//
// Why this is a separate file from `browser.test.js`:
//   * The worker page does NOT load pptxgenjs from the CDN; instead, the
//     dedicated Web Worker (`demos/browser/js/pptxgenjs_worker.js`) issues
//     `importScripts('./pptxgen.bundle.js')`, fetching `demos/browser/js/
//     pptxgen.bundle.js` directly from disk.
//   * Playwright `page.route()` / `context.route()` does NOT reliably
//     intercept `importScripts` requests originating from dedicated Web
//     Workers (open issues microsoft/playwright#5952 and #6403). So the
//     CDN-substitution strategy `browser.test.js` uses cannot work here.
//
// Strategy:
//   1. Pre-flight: refuse to start if `demos/browser/js/pptxgen.bundle.js` is
//      already dirty in git — we don't want to stomp on local edits.
//   2. Setup: read the original bytes of `demos/browser/js/pptxgen.bundle.js`
//      into memory, then overwrite with the fresh bundle the runner built at
//      `process.env.PPTXGEN_LOCAL_BUNDLE`.
//   3. Run a single case: open `worker_test.html`, click `#generatePptWorker`,
//      capture the resulting `.pptx` download, validate, assert zero schema
//      errors.
//   4. Teardown (always — `try/finally` in the runner): write the original
//      bytes back. Belt-and-braces: ALSO run `git checkout -- <path>` if the
//      file is tracked, in case the in-memory restore was somehow corrupted.
//      Verify the file matches its committed state via `git status`.
//
// The worker demo includes `slide.addImage({ path: "https://raw.github..." })`
// — i.e., this test has a transient internet dependency. Documented in the
// commit body. Fails loudly if offline.

const { spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const { isInstalled } = require('../validator')
const { assert, expectNoSchemaErrors, withTempDir } = require('./_helpers')
const { startServer } = require('./_server')

const WORKER_URL = 'http://localhost:8000/browser/worker_test.html'

const REPO_ROOT = path.resolve(__dirname, '..', '..')
const DEMO_BUNDLE = path.join(REPO_ROOT, 'demos', 'browser', 'js', 'pptxgen.bundle.js')
const LOCAL_BUNDLE = process.env.PPTXGEN_LOCAL_BUNDLE || path.join(REPO_ROOT, 'dist', 'pptxgen.bundle.js')

const DOWNLOAD_TIMEOUT_MS = 90000

function gitStatusPorcelain (file) {
	const r = spawnSync('git', ['-P', 'status', '--porcelain', '--', file], { cwd: REPO_ROOT, encoding: 'utf8' })
	if (r.status !== 0) {
		// Not in a git repo or git unavailable — best-effort fallback.
		return null
	}
	return r.stdout
}

function gitIsTracked (file) {
	const r = spawnSync('git', ['-P', 'ls-files', '--error-unmatch', '--', file], { cwd: REPO_ROOT, encoding: 'utf8' })
	return r.status === 0
}

function gitCheckoutFile (file) {
	const r = spawnSync('git', ['-P', 'checkout', '--', file], { cwd: REPO_ROOT, encoding: 'utf8' })
	return r.status === 0
}

async function setup (ctx) {
	if (!isInstalled()) throw new Error('OOXMLValidatorCLI not installed; run ./tools/ooxml-validator/install.sh')
	if (!fs.existsSync(LOCAL_BUNDLE)) throw new Error('local bundle missing at ' + LOCAL_BUNDLE + '; run `node test/release/_runner.js` rather than this test directly')
	if (!fs.existsSync(DEMO_BUNDLE)) throw new Error('demo worker bundle missing at ' + DEMO_BUNDLE)

	// Pre-flight: refuse to swap if the demo bundle is already dirty.
	const status = gitStatusPorcelain(DEMO_BUNDLE)
	if (status === null) {
		// Not in a git repo — no safety net; proceed but warn.
		console.log('  [worker] git unavailable; bundle restore will rely on in-memory snapshot only')
	} else if (status.trim() !== '') {
		throw new Error('refusing to swap ' + path.relative(REPO_ROOT, DEMO_BUNDLE) + ': file is already modified (git status: ' + JSON.stringify(status.trim()) + '). Commit or stash your changes first.')
	}

	let chromium
	try {
		chromium = require('playwright').chromium
	} catch (e) {
		throw new Error('playwright is not installed; run `npm install` (and `npx playwright install chromium` if needed): ' + e.message)
	}

	// Snapshot original bytes BEFORE overwriting — this is the primary restore path.
	ctx.originalBundleBytes = fs.readFileSync(DEMO_BUNDLE)
	ctx.bundleSwapped = false
	ctx.tracked = gitIsTracked(DEMO_BUNDLE)

	const freshBytes = fs.readFileSync(LOCAL_BUNDLE)
	fs.writeFileSync(DEMO_BUNDLE, freshBytes)
	ctx.bundleSwapped = true

	ctx.server = await startServer()
	try {
		ctx.browser = await chromium.launch({ headless: true })
	} catch (e) {
		await ctx.server.kill()
		throw e
	}
}

async function teardown (ctx) {
	if (ctx && ctx.browser) {
		try { await ctx.browser.close() } catch (_) { /* ignore */ }
	}
	if (ctx && ctx.server) {
		try { await ctx.server.kill() } catch (_) { /* ignore */ }
	}

	// Restore the demo bundle.
	if (ctx && ctx.bundleSwapped && ctx.originalBundleBytes) {
		try {
			fs.writeFileSync(DEMO_BUNDLE, ctx.originalBundleBytes)
		} catch (e) {
			console.log('  [worker] failed to restore ' + path.relative(REPO_ROOT, DEMO_BUNDLE) + ' from in-memory snapshot: ' + e.message)
		}
		// Belt-and-braces: if tracked by git, reset to the committed bytes too.
		if (ctx.tracked) {
			const ok = gitCheckoutFile(DEMO_BUNDLE)
			if (!ok) {
				console.log('  [worker] WARNING: `git checkout` failed for ' + path.relative(REPO_ROOT, DEMO_BUNDLE) + '; verify working tree manually')
			}
		}
		// Final check: the file should now be clean.
		const status = gitStatusPorcelain(DEMO_BUNDLE)
		if (status !== null && status.trim() !== '') {
			throw new Error('worker teardown failed to restore ' + path.relative(REPO_ROOT, DEMO_BUNDLE) + ' (git status: ' + JSON.stringify(status.trim()) + ')')
		}
	}
}

async function runWorkerCase (ctx) {
	await withTempDir('pptxgen-release-worker-', async (tmpDir) => {
		const context = await ctx.browser.newContext({ acceptDownloads: true })

		const pageErrors = []
		const consoleErrors = []

		try {
			const page = await context.newPage()

			page.on('pageerror', err => { pageErrors.push(err.message) })
			page.on('console', msg => { if (msg.type() === 'error') consoleErrors.push(msg.text()) })

			await page.goto(WORKER_URL, { waitUntil: 'load' })
			await page.waitForLoadState('networkidle')
			await page.waitForSelector('#generatePptWorker', { state: 'visible', timeout: 15000 })

			const [download] = await Promise.all([
				page.waitForEvent('download', { timeout: DOWNLOAD_TIMEOUT_MS }),
				page.click('#generatePptWorker')
			])

			const suggested = download.suggestedFilename() || 'worker_demo.pptx'
			const out = path.join(tmpDir, suggested)
			await download.saveAs(out)

			const stat = fs.statSync(out)
			assert(
				stat.size > 0,
				'downloaded file is empty: ' + out +
					' pageErrors=' + JSON.stringify(pageErrors) +
					' consoleErrors=' + JSON.stringify(consoleErrors)
			)

			await expectNoSchemaErrors(out, 'worker/generatePptWorker')
		} finally {
			try { await context.close() } catch (_) { /* ignore */ }
		}
	})
}

module.exports = {
	setup,
	teardown,
	cases: [
		{
			name: 'worker/generatePptWorker: produces a schema-valid .pptx',
			fn: runWorkerCase
		}
	]
}
