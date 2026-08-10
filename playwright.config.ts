import { defineConfig, devices } from '@playwright/test'

/**
 * The browser lane.
 *
 * The Node suite (`vitest.config.ts`) proves the emission core against `dist/node.js`.
 * This proves the same core in a real browser, across two fixtures that answer different
 * questions:
 *
 *   - **demo** — drives the site's own demos page, which imports the *same* showcase module
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
 *     `pickColWidthBasis` had never executed anywhere. It also runs the conversions that
 *     basis feeds — auto-paging a table too tall for one slide — and compares them against
 *     the same conversion on a DOM that renders nothing, which is how the long-dismissed
 *     `tableToSlides` overflow report was shown to have nothing to do with the DOM.
 *
 * What this lane deliberately does NOT cover is live-DOM layout **fidelity** — whether the
 * browser's numbers are the right numbers, or whether two engines agree on them. The
 * distinction is fine but load-bearing: `html-table` asserts that a measurement is taken
 * and honoured proportionally, never that it matches what the page painted. Runtime
 * support and layout fidelity are separate claims (see docs/project-target.md "Out Of
 * Active Scope"); this lane moves only the first one.
 *
 * Run it with `pnpm run test:browser`, which builds `dist/` and the site first.
 */

// `vitepress preview`'s default, restated because `webServer.url` below has to agree with
// it. The preview server takes no `--strictPort`: `listen` on a taken port throws
// EADDRINUSE and the command exits, which is the behaviour that flag would have bought.
const PORT = 4173

// The harness server, on the next port up. Its own origin rather than a route on the
// preview server: the preview serves the site's build output, and mounting repo paths into
// it would mean the site's config decides what the adapter tests can reach.
const HARNESS_PORT = 4174

// `vitepress preview` takes no `--host` — it listens on every interface — so naming the v4
// loopback here is a client-side choice, and it is the one that behaves the same on Windows
// and on CI's Linux. `localhost` would resolve to `::1` only on Windows, and a v4 probe of
// the same port gets nothing: a 60s webServer timeout with no clue why.
const HOST = '127.0.0.1'

// The published site's own base (`docs/.vitepress/config.mts`), which `vitepress preview`
// honours — the site is NOT at `/`. The demos page is a route under it, so the baseURL is
// the site root and the spec navigates to `./demos`.
const BASE_URL = `http://${HOST}:${PORT}/ts-pptx/`

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
	// Playwright's default is 30s, which the `demo` project no longer fits: it drives a page
	// of the real site, so a run is the async chunk downloading, then a deck built in the
	// tab, then that deck read back and rendered. Raised for all three projects rather than
	// per-project — a limit that differs by project is a limit nobody remembers the shape of.
	timeout: 60_000,
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
			// `pnpm run test:browser` builds the site; this only serves it. Keeping the build
			// out of the webServer command means a failed build fails as a build, with the
			// bundler's own output, instead of as a webServer timeout.
			command: `pnpm exec vitepress preview docs --port ${PORT}`,
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
