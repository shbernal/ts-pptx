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
//
// `branches` trails the other three by design. The read model guards every
// element lookup (`x ? … : null`) whether or not the schema lets `x` be absent,
// so a standing share of the branch count is unreachable on any valid package —
// see docs/testing.md "Branches that are not worth covering" for which of those
// to leave alone and which are real input worth a test. Two files carry that
// reasoning in full, per remaining arm: test/read/chrome-read-edges.test.js for
// src/read/api/chrome.ts, and test/read/import-slide-preserve.test.js for
// src/read/oxml/theme.ts.
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
		// Vitest's default 5s is a per-test *wall-clock* budget, which stops being a
		// property of the test once validators run concurrently: a fixture that
		// validates a large deck spends most of those 5s queued behind its peers for
		// CPU, not working. Two of them (the hierarchical-chartEx and
		// carryMasterGraphics decks) sat right on the line and failed intermittently.
		// Raised so the timeout is what it is meant to be — a hang detector, not a
		// performance assertion. Lower `maxConcurrency` before lowering this.
		testTimeout: 30_000,
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
				// Raised 91 -> 92 once the table auto-pager landed: measured 93.21, so the gate
				// keeps well over a point of slack.
				statements: 92,
				// Raised 79 -> 80 once measure-fit/text-fit landed: measured 81.42, so the
				// gate keeps a full point of slack. Ratchet upward only — if a change drops a
				// number below its gate, that is a finding to explain, never a gate to lower.
				branches: 80,
				// Left alone deliberately: measured 98.33, so the next notch would keep 0.33,
				// and a gate with less slack than a point is a gate that goes red on noise.
				functions: 97,
				// Raised 94 -> 95 once the zoom/background definers landed: measured 96.00, the
				// first time this axis has cleared a full point behind the notch.
				lines: 95,
			},
		},
	},
})
