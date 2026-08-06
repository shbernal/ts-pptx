import { defineConfig, devices } from '@playwright/test'

/**
 * The browser lane.
 *
 * The Node suite (`vitest.config.ts`) proves the emission core against `dist/node.js`.
 * This proves the same core in a real browser, driving `demos/vite-demo` — which imports
 * the *same* showcase module `pnpm demos:build quarterly-review` runs. Nothing here is a
 * second copy of the deck: the demo is the only fixture, and the assertions live in
 * `test/browser/`.
 *
 * What this lane deliberately does NOT cover is live-DOM layout fidelity — real
 * `offsetWidth` after layout, the resolved cascade, browser-chosen fonts. Runtime support
 * and layout fidelity are separate claims (see docs/project-target.md "Out Of Active
 * Scope"); this lane moves only the first one.
 *
 * Run it with `pnpm run test:browser`, which builds `dist/` and the demo first.
 */

// Vite's preview default. Pinned rather than left to chance because `base` below has to
// agree with it, and because `--strictPort` must fail loudly instead of silently serving
// the wrong app on the next free port.
const PORT = 4173

// Bound explicitly below with `--host 127.0.0.1`. Left to itself, `vite preview` binds
// `localhost`, which on Windows resolves to `::1` **only** — a v4 probe of the same port
// gets nothing and the webServer wait times out at 60s with no clue why. Naming the v4
// loopback on both sides makes the lane behave the same on Windows and on CI's Linux
// instead of depending on how the host orders its loopback records.
const HOST = '127.0.0.1'

// `demos/vite-demo/vite.config.ts` sets `base: '/TsPptx/demos/vite/'` for the published
// GitHub Pages demo, and `vite preview` honours it — the app is NOT at `/`.
const BASE_URL = `http://${HOST}:${PORT}/TsPptx/demos/vite/`

export default defineConfig({
	testDir: './test/browser',
	// Under `.tmp/` (gitignored) with the rest of this repo's generated artifacts, rather
	// than Playwright's default `test-results/` at the root, which nothing ignores.
	outputDir: './.tmp/playwright',
	// One deck build per test, each ~seconds; there is no value in parallelising two of
	// them across browser processes, and serial output is far easier to read on failure.
	workers: 1,
	fullyParallel: false,
	forbidOnly: !!process.env['CI'],
	// No retries: every assertion here is deterministic by construction (`FIXED_MTIME`
	// pins the zip, the demo takes no network). A retry would convert a real flake —
	// which is a finding — into a green run.
	retries: 0,
	reporter: [['list']],
	use: {
		baseURL: BASE_URL,
		trace: 'retain-on-failure',
	},
	// Chromium only, deliberately. The adapter surface this exercises (`fetch`,
	// `FileReader`, canvas, object URLs, `<a download>`) is uncontroversial across
	// engines, and a matrix costs CI time for a divergence nobody has observed. Add
	// Firefox/WebKit when something concrete surfaces, not pre-emptively.
	projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
	webServer: {
		// `pnpm run test:browser` builds the demo; this only serves it. Keeping the build
		// out of the webServer command means a failed build fails as a build, with the
		// bundler's own output, instead of as a webServer timeout.
		command: `pnpm --dir demos/vite-demo exec vite preview --host ${HOST} --port ${PORT} --strictPort`,
		url: BASE_URL,
		reuseExistingServer: !process.env['CI'],
		timeout: 60_000,
	},
})
