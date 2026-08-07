import { defineConfig, devices } from '@playwright/test'

/**
 * The browser lane.
 *
 * The Node suite (`vitest.config.ts`) proves the emission core against `dist/node.js`.
 * This proves the same core in a real browser, across two fixtures that answer different
 * questions:
 *
 *   - **demo** — drives `demos/vite-demo`, which imports the *same* showcase module
 *     `pnpm demos:build quarterly-review` runs. This is the bundled story: a real
 *     consumer, Vite resolving the `browser` export condition, Rollup tree-shaking it.
 *     It exercises `writeFile` (the object-URL `<a download>` path) and proves the
 *     emission core is runtime-invariant, but it never loads an asset — the deck draws
 *     every one of them — so it cannot reach the rest of the adapter.
 *   - **runtime-adapter** — loads the shipped `dist/browser.js` unbundled off a static
 *     server (scripts/browser-harness-server.mjs) and drives it with decks built to hit
 *     `loadMedia`, `createSvgPngPreview` and `loadFontData`, including their failure
 *     arms. Real URLs, real 404s, no stubbed `fetch`.
 *   - **html-table** — renders a real `<table>` on that same server and converts it, so
 *     `tableToSlides` reads a non-zero `offsetWidth`. That is the one width basis no Node
 *     lane can produce: happy-dom reports `0` for every cell, so the measured arm of
 *     `pickColWidthBasis` had never executed anywhere.
 *
 * What this lane deliberately does NOT cover is live-DOM layout **fidelity** — whether the
 * browser's numbers are the right numbers, or whether two engines agree on them. The
 * distinction is fine but load-bearing: `html-table` asserts that a measurement is taken
 * and honoured proportionally, never that it matches what the page painted. Runtime
 * support and layout fidelity are separate claims (see docs/project-target.md "Out Of
 * Active Scope"); this lane moves only the first one.
 *
 * Run it with `pnpm run test:browser`, which builds `dist/` and the demo first.
 */

// Vite's preview default. Pinned rather than left to chance because `base` below has to
// agree with it, and because `--strictPort` must fail loudly instead of silently serving
// the wrong app on the next free port.
const PORT = 4173

// The harness server, on the next port up. Its own origin rather than a route on the
// preview server: `vite preview` serves the demo's build output, and mounting repo paths
// into it would mean the demo's config decides what the adapter tests can reach.
const HARNESS_PORT = 4174

// Bound explicitly below with `--host 127.0.0.1`. Left to itself, `vite preview` binds
// `localhost`, which on Windows resolves to `::1` **only** — a v4 probe of the same port
// gets nothing and the webServer wait times out at 60s with no clue why. Naming the v4
// loopback on both sides makes the lane behave the same on Windows and on CI's Linux
// instead of depending on how the host orders its loopback records.
const HOST = '127.0.0.1'

// `demos/vite-demo/vite.config.ts` sets `base: '/TsPptx/demos/vite/'` for the published
// GitHub Pages demo, and `vite preview` honours it — the app is NOT at `/`.
const BASE_URL = `http://${HOST}:${PORT}/TsPptx/demos/vite/`

// The harness page sits at its real repo path, so the relative `../../../dist/browser.js`
// inside `harness.mjs` is the same specifier on disk and over HTTP — one path that both
// the browser and `pnpm run typecheck:test` resolve.
const HARNESS_URL = `http://${HOST}:${HARNESS_PORT}/test/browser/harness/`

export default defineConfig({
	testDir: './test/browser',
	// Empties `.tmp/browser-coverage/` so what the collector leaves there is exactly what
	// this run executed — see test/browser/coverage-setup.mjs for why a narrowed run makes
	// that worth doing at the start rather than at merge time.
	globalSetup: './test/browser/coverage-setup.mjs',
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
		trace: 'retain-on-failure',
	},
	// Chromium only, and the decision is written down rather than left as a default —
	// see docs/runtime-and-package-support.md "Which Browsers The Lane Runs". Short
	// version: the APIs in play (`fetch`, `FileReader`, canvas, object URLs,
	// `<a download>`) are uncontroversial across engines, so a matrix would cost CI time
	// per push to re-answer a question nothing has raised. Add Firefox/WebKit when a
	// concrete divergence surfaces, not pre-emptively.
	//
	// Projects rather than separate configs: they differ only in which page they drive, and
	// all of them must run on one `pnpm run test:browser`.
	//
	// **Every project matches by filename prefix, and none of them matches by exclusion.**
	// `demo` used to be spelled `testIgnore: ['adapter-*.spec.mjs']`, which quietly meant
	// "everything not yet invented" — the first spec added under a third prefix would have
	// run a second time against the demo's baseURL and failed for a reason having nothing to
	// do with what it tests. A positive match makes the pairing between a spec's name and
	// its fixture explicit, and makes adding a prefix a visible edit here.
	projects: [
		{
			name: 'demo',
			testMatch: ['deck-*.spec.mjs', 'cross-runtime-*.spec.mjs'],
			use: { ...devices['Desktop Chrome'], baseURL: BASE_URL },
		},
		{
			name: 'runtime-adapter',
			testMatch: ['adapter-*.spec.mjs'],
			use: { ...devices['Desktop Chrome'], baseURL: HARNESS_URL },
		},
		{
			// Same server as `runtime-adapter`, different page and a different question — the
			// two are kept apart so neither fixture's DOM has anything in it the other put
			// there. See test/browser/harness/table.mjs.
			name: 'html-table',
			testMatch: ['table-*.spec.mjs'],
			use: { ...devices['Desktop Chrome'], baseURL: HARNESS_URL },
		},
	],
	webServer: [
		{
			// `pnpm run test:browser` builds the demo; this only serves it. Keeping the build
			// out of the webServer command means a failed build fails as a build, with the
			// bundler's own output, instead of as a webServer timeout.
			command: `pnpm --dir demos/vite-demo exec vite preview --host ${HOST} --port ${PORT} --strictPort`,
			url: BASE_URL,
			reuseExistingServer: !process.env['CI'],
			timeout: 60_000,
		},
		{
			// `process.execPath`, not `node`: the repo's scripts avoid resolving a bare
			// command name on Windows, where the PATH entry is a `.cmd` shim.
			command: `"${process.execPath}" scripts/browser-harness-server.mjs --host ${HOST} --port ${HARNESS_PORT}`,
			url: HARNESS_URL,
			reuseExistingServer: !process.env['CI'],
			timeout: 30_000,
		},
	],
})
