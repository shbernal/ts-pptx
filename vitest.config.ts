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
		coverage: {
			provider: 'v8',
			include: ['dist/**/*.js'],
			// Browser-only entry points are out of the Node suite's scope (they call
			// `fetch`, `document`, DOM layout APIs that cannot run headless). They map
			// to their own bundled chunks, so exclude those chunks rather than the
			// shared `pptxgen` chunk. Partial-file browser code that shares a chunk with
			// tested code (e.g. `genTableToSlides`) is fenced with `v8 ignore` comments
			// at the source instead. See docs/project-target.md "Out Of Active Scope".
			exclude: [
				...coverageConfigDefaults.exclude,
				'dist/browser.js', // src/browser.ts — browser entry
				'dist/browser-*.js', // src/runtime/browser.ts — browser runtime adapter
				'dist/standalone.js', // src/standalone.ts — browser IIFE bundle (not exercised)
			],
			// `json-summary` writes coverage/coverage-summary.json (per-file + total
			// rollup) and `json` writes coverage/coverage-final.json (raw per-line map)
			// so agents and ratchet scripts can read coverage without scraping the HTML.
			reporter: ['text-summary', 'text', 'html', 'json-summary', 'json'],
			thresholds: {
				statements: 89,
				branches: 74,
				functions: 95,
				lines: 92,
			},
		},
	},
})
