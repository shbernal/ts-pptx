import { coverageConfigDefaults, defineConfig } from 'vitest/config'

// The suite runs against the built package (`pnpm run build` then `vitest run`),
// so tests import from `dist/`, not `src/`. v8 collects coverage for the code it
// actually executes — the bundled `dist/` output — and remaps line/branch data
// back to `src/` via the sourcemaps tsdown emits. Instrumenting `src/**` instead
// would report ~8% because almost nothing under `src/` is executed directly.
//
// Thresholds are pinned a notch below the current measured numbers so an
// accidental coverage regression fails CI without the gate being flaky. Ratchet
// them upward as coverage improves; never loosen them to make a red build pass.
export default defineConfig({
	test: {
		// The schema fixtures are `describe.concurrent` and each concurrent test
		// spawns an OOXMLValidatorCLI (.NET) process; `test/read` spawns validators
		// too, so with a bare `vitest run` the real process ceiling is
		// workers × maxConcurrency. Cap it deliberately rather than leaving the
		// default (5) to interact with the worker pool by accident. If CI turns
		// flaky or OOMs, lower this — do not re-serialize the suite, that is the
		// 50s → ~15s the concurrency bought.
		maxConcurrency: 8,
		coverage: {
			provider: 'v8',
			include: ['dist/**/*.js'],
			// Browser-only entry points are out of the Node suite's scope (they call
			// `fetch`, `document`, DOM layout APIs that cannot run headless). They map
			// to their own bundled chunks, so exclude those chunks rather than the
			// shared `pptxgen` chunk. Live-DOM feature code (`tableToSlides` /
			// `html-dom.ts`) is imported only by the browser entry, so it bundles into
			// these chunks too and needs no in-source `v8 ignore` fence. See
			// docs/project-target.md "Out Of Active Scope".
			exclude: [
				...coverageConfigDefaults.exclude,
				'dist/browser.js', // src/browser.ts — browser entry
				'dist/browser-*.js', // src/runtime/browser.ts — browser runtime adapter
			],
			// `json-summary` writes coverage/coverage-summary.json (per-file + total
			// rollup) and `json` writes coverage/coverage-final.json (raw per-line map)
			// so agents and ratchet scripts can read coverage without scraping the HTML.
			reporter: ['text-summary', 'text', 'html', 'json-summary', 'json'],
			thresholds: {
				statements: 90,
				branches: 76,
				functions: 96,
				lines: 93,
			},
		},
	},
})
