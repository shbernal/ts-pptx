import { configDefaults, coverageConfigDefaults, defineConfig } from 'vitest/config'

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
		// `test/browser/**` belongs to Playwright (`playwright.config.ts`, `pnpm run
		// test:browser`), not to Vitest. Its specs are named `*.spec.mjs`, which Vitest's
		// default `include` matches, so without this it would collect them and fail on
		// `@playwright/test`'s fixtures. Excluded by directory rather than by filename so
		// the two harnesses never race for a file on the strength of what it is called.
		exclude: [...configDefaults.exclude, 'test/browser/**'],
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
			// `dist/browser.js` used to be excluded here, with `dist/browser-*.js`
			// alongside it, on the grounds that the browser entry and its runtime adapter
			// "cannot run headless". Both are gone as of the browser lane: the adapter is
			// exercised in a real Chromium (test/browser/adapter-*.spec.mjs), and the
			// premise was false in a second way — tsdown bundles `src/runtime/browser.ts`
			// *into* `dist/browser.js`, so `dist/browser-*.js` never matched anything.
			//
			// The number below is therefore the Node suite's alone, and that is the
			// honest reading of it: the browser lane's coverage is not merged in (see
			// docs/testing.md "Browser Lane"). What the Node suite does cover of that file
			// is real — three regression tests import the entry — and what it does not is
			// the four adapter functions, which is exactly what the browser lane asserts.
			//
			// `html-dom.ts` (the `tableToSlides` conversion) used to be excluded by these
			// same globs too, on the grounds that only the browser entry imported it. That
			// is no longer true: the `ts-pptx/html` entry imports it too, so tsdown emits
			// it as its own shared `dist/html-dom-*.js` chunk — also not excluded. The
			// Node suite executes it against a real DOM
			// (test/regression/html-to-slides-node.test.js), so it is covered code now,
			// not unreachable code.
			//
			// Nothing of this repo's own is excluded any more. The defaults are restated
			// rather than left implicit so that a future exclusion has an obvious home —
			// and so that adding one is a visible edit, not the absence of an edit.
			exclude: [...coverageConfigDefaults.exclude],
			// `json-summary` writes coverage/coverage-summary.json (per-file + total
			// rollup) and `json` writes coverage/coverage-final.json (raw per-line map)
			// so agents and ratchet scripts can read coverage without scraping the HTML.
			reporter: ['text-summary', 'text', 'html', 'json-summary', 'json'],
			thresholds: {
				// Raised 91 -> 92 once the table auto-pager landed: measured 93.21, so the gate
				// keeps well over a point of slack.
				statements: 92,
				// Raised 80 -> 81 once the text and chart definers landed: measured 82.79, so the
				// gate keeps 1.79. 82 was available on the raw number and declined — it would
				// keep 0.79, under the point of slack every notch here is required to leave.
				// Ratchet upward only — if a change drops a number below its gate, that is a
				// finding to explain, never a gate to lower.
				branches: 81,
				// Left at 97, and now the tightest gate here: dropping the `dist/browser.js`
				// exclusion put `src/runtime/browser.ts`'s 13 functions into the denominator
				// with 1 of them reachable from Node, so the measured number fell 98.33 ->
				// 97.35 and the slack fell with it, to 0.35.
				//
				// That is the price of an honest denominator, not a regression, and the way
				// to buy the slack back is to merge the browser lane's coverage into this
				// report — not to lower the gate. Until then those 12 functions are gated
				// where they actually run, by test/browser/adapter-coverage.spec.mjs.
				functions: 97,
				// Raised 94 -> 95 once the zoom/background definers landed: measured 96.00, the
				// first time this axis has cleared a full point behind the notch.
				lines: 95,
			},
		},
	},
})
