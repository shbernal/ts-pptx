import fs from 'node:fs'
import path from 'node:path'
import { test as base } from '@playwright/test'
import { ROOT } from '../../scripts/script-utils.mjs'

/**
 * The browser lane's coverage collector.
 *
 * Every spec in the `runtime-adapter` project imports `test` from here instead of from
 * `@playwright/test`, and gets one extra thing for it: Chromium's V8 coverage for the
 * shipped `dist/*.js` files the harness served, written to `.tmp/browser-coverage/` for
 * `scripts/coverage-merge.mjs` to fold into the Node report.
 *
 * Why a fixture rather than a line in each spec: what the merge wants is *every* scenario
 * the lane runs, and a per-spec opt-in is a thing to forget. Auto-use means a new
 * adapter spec contributes its coverage by existing, and the only way to leave the merge
 * short is to import `test` from the wrong module — which is visible in the diff.
 *
 * Only `/dist/*.js` URLs are kept. The harness page's own module, the import-mapped
 * dependencies it pulls from `node_modules/` and anything the browser loaded of its own
 * accord are not this package's shipped code and are not in the Node report's file set
 * either; keeping them would put files in the merged denominator that the Node lane never
 * had a chance to cover.
 *
 * The `demo` project does not use this `test`, and could not benefit from it: Vite
 * re-bundles the package into the app's own chunks, so nothing there is served as
 * `dist/browser.js` and no per-file coverage entry for it exists. What the demo proves it
 * proves by byte-identity (test/browser/cross-runtime-bytes.spec.mjs), not by coverage.
 */

/** Where `scripts/coverage-merge.mjs` looks. Cleared per run by `coverage-setup.mjs`. */
export const BROWSER_COVERAGE_DIR = path.join(ROOT, '.tmp', 'browser-coverage')

/** A shipped file, addressed the way the harness server serves the repo. */
const DIST_SCRIPT = /^\/dist\/[^/]+\.js$/

/**
 * What a spec gets when it takes the `jsCoverage` fixture: one call, giving the entries
 * Chromium collected for this test. Calling it more than once is safe and returns the same
 * entries — see the fixture body.
 * @typedef {{ stop(): Promise<any[]> }} JsCoverage
 */

export const test = base.extend(
	/**
	 * The cast is what tells `extend` the shape it is adding — `.mjs` has nowhere to write
	 * the type argument `base.extend<{jsCoverage: JsCoverage}>(…)` would take in TypeScript,
	 * and without it `pnpm run typecheck:test` sees a fixture that no spec is allowed to
	 * name.
	 * @type {import('@playwright/test').Fixtures<{jsCoverage: JsCoverage}, {}, import('@playwright/test').PlaywrightTestArgs & import('@playwright/test').PlaywrightTestOptions, import('@playwright/test').PlaywrightWorkerArgs & import('@playwright/test').PlaywrightWorkerOptions>}
	 */ ({
		/**
		 * Starts V8 coverage before the test body — before `beforeEach` hooks too, which is
		 * why `openHarness` in a hook is still measured — and writes what it collected on the
		 * way out.
		 *
		 * A test that wants to *assert* on the coverage (adapter-coverage.spec.mjs) takes the
		 * fixture and calls `stop()` itself; the result is cached, so the teardown's own call
		 * gets the same entries rather than a second, empty collection. Playwright's
		 * `stopJSCoverage` can only be called once per page.
		 */
		jsCoverage: [
			async ({ page }, use, testInfo) => {
				if (!page.coverage) throw new Error('page.coverage is a Chromium API; this lane is Chromium-only by design')
				await page.coverage.startJSCoverage({ resetOnNavigation: false })

				let entries = null
				const collector = {
					async stop() {
						entries ??= await page.coverage.stopJSCoverage()
						return entries
					},
				}

				await use(collector)

				writeEntries(await collector.stop(), testInfo)
			},
			{ auto: true },
		],
	})
)

/**
 * One file per test, named by Playwright's own stable test id so a re-run overwrites its
 * predecessor instead of stacking a second copy of the same scenario into the merge.
 *
 * `source` is not stored — the merge reads the file off disk, which is the same file the
 * harness server streamed. `sourceLength` is kept as the check on that claim: V8's ranges
 * are byte offsets into what the browser parsed, so a length that disagrees with the file
 * on disk means the two are not the same build and every offset below is meaningless.
 */
function writeEntries(entries, testInfo) {
	const kept = []
	for (const entry of entries) {
		const { pathname } = new URL(entry.url)
		if (!DIST_SCRIPT.test(pathname)) continue
		kept.push({ path: pathname, sourceLength: entry.source?.length ?? null, functions: entry.functions })
	}
	if (!kept.length) return

	fs.mkdirSync(BROWSER_COVERAGE_DIR, { recursive: true })
	const file = path.join(BROWSER_COVERAGE_DIR, `${testInfo.testId}.json`)
	const record = {
		spec: path.relative(ROOT, testInfo.file).replace(/\\/g, '/'),
		title: testInfo.titlePath.join(' > '),
		project: testInfo.project.name,
		entries: kept,
	}
	fs.writeFileSync(file, JSON.stringify(record) + '\n')
}

export { expect } from '@playwright/test'
