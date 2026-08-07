import fs from 'node:fs'
import { BROWSER_COVERAGE_DIR } from './fixtures.mjs'

/**
 * Playwright `globalSetup`: empty `.tmp/browser-coverage/` before the lane runs.
 *
 * The collector names each file after a stable test id, so a full re-run overwrites its
 * own output — but a *narrowed* run does not. `playwright test adapter-media` would leave
 * the previous run's font and coverage scenarios sitting in the directory, and
 * `scripts/coverage-merge.mjs` would fold them in as if this run had produced them. The
 * merged number would then describe a lane that never ran, which is the one thing a
 * coverage gate must never do.
 *
 * Clearing here rather than in the merge script keeps that property with the run that
 * produces the data: whatever is in the directory afterwards is exactly what this
 * invocation of Playwright executed.
 */
export default function clearBrowserCoverage() {
	fs.rmSync(BROWSER_COVERAGE_DIR, { recursive: true, force: true })
}
