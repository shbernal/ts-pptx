import fs from 'node:fs'
import os from 'node:os'
import { configDefaults, coverageConfigDefaults, defineConfig } from 'vitest/config'

// ---------------------------------------------------------------------------
// Worker ceiling
//
// Vitest sizes its fork pool from the CPU (`availableParallelism() - 1`), which
// made this suite's footprint a property of the developer's machine rather than
// of this repo — and the wrong property. Cores decide how fast it *could* run;
// memory decides whether the host survives it. A faster CPU made the spike
// bigger, never the run safer.
//
// Measured on the batched suite (see docs/testing.md "Suite cost and the worker
// ceiling"), peak RSS is close to linear in the pool size:
//
//     4 workers 1.66 GB   6 workers 2.27 GB   8 workers 2.68 GB   11 workers 3.40 GB
//
// which is BASE_MB + PER_WORKER_MB × workers to within ~50 MB across all four
// points. So the pool can be solved for directly from the memory actually free
// at startup instead of being guessed at or pinned to a timid constant. On an
// idle machine this lands on the CPU bound and nothing is given up; next to a
// browser and an agent it scales itself down instead of pushing the host into
// swap.
//
// This is a ceiling, not a reservation — the numbers are a model of observed
// peaks, so treat them as a budget to stay under, and re-measure if the suite's
// shape changes materially.
// ---------------------------------------------------------------------------

const BASE_MB = 700
const PER_WORKER_MB = 250
/** Leave the rest of `available` to the desktop, the editor, and page cache. */
const BUDGET_FRACTION = 0.6

/**
 * Memory this machine could actually hand over right now, in MB.
 *
 * `os.freemem()` is the portable answer but means different things per platform:
 * on Linux it is `MemFree`, which excludes reclaimable page cache and so badly
 * understates what is available (measured here: 9.1 GB free against 12.4 GB
 * available). Prefer `MemAvailable` where the kernel publishes it — that is the
 * number this budget wants — and fall back to `freemem()` elsewhere, which is
 * already the right quantity on Windows and macOS.
 */
function availableMemoryMb(): number {
	try {
		const meminfo = fs.readFileSync('/proc/meminfo', 'utf8')
		const match = /^MemAvailable:\s+(\d+) kB$/m.exec(meminfo)
		if (match) return Number(match[1]) / 1024
	} catch {
		// Not Linux, or /proc is not mounted — fall through.
	}
	return os.freemem() / (1024 * 1024)
}

function resolveMaxWorkers(): number {
	// Escape hatch first, for pinning the pool where a predictable cost beats an adaptive
	// one — bisecting a concurrency-dependent failure, or a CI job that wants its footprint
	// fixed rather than dependent on whatever else the runner is doing. Nothing under
	// `.github/workflows/` sets it today; this comment used to assert that CI did, which
	// was never true. The hosted runners are small enough that `byCpu` binds well below the
	// memory budget anyway.
	const override = Number(process.env.VITEST_MAX_WORKERS)
	if (Number.isInteger(override) && override > 0) return override

	const byCpu = Math.max(1, os.availableParallelism() - 1)
	const budgetMb = availableMemoryMb() * BUDGET_FRACTION
	const byMemory = Math.floor((budgetMb - BASE_MB) / PER_WORKER_MB)
	// Never resolve to zero: one worker at a time is slow, but it is the only
	// setting that still makes progress on a machine with nothing left to give.
	const workers = Math.max(1, Math.min(byCpu, byMemory))
	if (workers < byCpu) {
		process.stderr.write(
			`[vitest] pool capped to ${workers} workers (CPU allows ${byCpu}) — ` +
				`${Math.round(availableMemoryMb())} MB available, budget ${Math.round(budgetMb)} MB\n`
		)
	}
	return workers
}

const maxWorkers = resolveMaxWorkers()

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
		// Bound the pool by memory, not by core count — see the header above.
		// `maxWorkers` is the whole knob: Vitest 4 has no `minWorkers` (it is not in
		// `InlineConfig`, and setting it is a type error rather than a no-op), and
		// the pool grows to this ceiling on demand rather than being preallocated.
		maxWorkers,
		// Share one module registry per worker instead of rebuilding it per test file.
		//
		// `dist/` is over 1 MB of JavaScript (`text-*.js` alone is 610 KB), 78 files import
		// `dist/node.js` and 64 import `dist/read.js`. Under Vitest's default isolation each
		// of the 235 files re-evaluates that graph from scratch, and the cost is not
		// marginal: measured across three paired runs at `VITEST_MAX_WORKERS=4`, the
		// `import` phase falls 136.0/114.5/84.8s to 34.4/30.6/22.0s — 3.5-4x, with no
		// overlap between the two groups — taking wall clock down 22-37%. Two smaller wins
		// come along with it, both from module state that now survives a file boundary:
		// the validator's batch queue can join requests across files instead of spawning
		// one .NET validator per file, and corpus.js's `irFor` memo stops being rebuilt
		// per file.
		//
		// What this gives up is a real guarantee, not a formality: isolation is what made
		// cross-file state leakage impossible rather than merely absent. Two things replace
		// it. `test/setup-globals.js` resets the one process-global the library owns
		// (`setDiagnosticHandler`) after every test, so the leak channel is closed by
		// construction. And `sequence.shuffle.files` below means a suite that grows an
		// order dependence fails on it instead of hiding behind a stable file order.
		//
		// The other module-level state under `src/` is idempotent caching — `math.ts`'s
		// lazy temml/mathml2omml handles and `measure/font-metrics.ts`'s heuristic
		// singleton — which is better shared than rebuilt.
		isolate: false,
		sequence: {
			// Randomize file order so that `isolate: false` cannot quietly acquire an
			// order dependence. Vitest prints the seed on failure; re-run with
			// `--sequence.seed=<n>` to reproduce one. Tests *within* a file stay in source
			// order on purpose — `captureDiagnostics` and the warn-capturing schema
			// fixtures rest on that, and shuffling them would trade a real guarantee for
			// nothing.
			shuffle: { files: true, tests: false },
		},
		setupFiles: ['./test/setup-globals.js'],
		// The schema fixtures are `describe.concurrent` and `test/read` validates
		// too. This used to be half of a `workers × maxConcurrency` process ceiling:
		// every concurrent test spawned its own .NET validator process, so this number
		// multiplied directly into RAM. It does not any more — `ooxml-validate` batches
		// validation requests and holds at most ONE validator child per worker
		// regardless of what this is set to, so the validator cost is `workers × one
		// oracle` and this knob only governs in-process test interleaving. Measured on
		// the full suite: 7 concurrent oracle children at the peak, whatever this says.
		// Raising it no longer buys spawn parallelism;
		// lower `maxWorkers` (or let the memory budget do it) if the suite needs to
		// shrink.
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
			// The number this config reports is still the Node suite's alone — one run of
			// one lane can only report what it executed — but it is no longer the number
			// the repo is judged on. `scripts/coverage-merge.mjs` folds the browser lane's
			// V8 coverage into it, and `pnpm run coverage:gate` checks the merged result;
			// see docs/testing.md "Merged coverage". What the Node suite covers of that
			// file is real — three regression tests import the entry — and the twelve
			// adapter functions it cannot reach are covered where they run.
			//
			// `html-dom.ts` (the `tableToSlides` conversion) used to be excluded by these
			// same globs too, on the grounds that only the browser entry imported it. That
			// is no longer true: the `ts-pptx/html` entry imports it too, so tsdown emits
			// it as its own shared `dist/html-dom-*.js` chunk — also not excluded. The
			// Node suite executes it against a real DOM
			// (test/regression/html/html-to-slides-node.test.js), so it is covered code now,
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
			// These four are the *Node suite's* floor, and only that. They may not be
			// lowered and they fail if this suite goes backwards — but the point-of-slack
			// rule is no longer held against them, because it cannot honestly be: this
			// report's denominator includes `src/runtime/browser.ts`, whose adapter needs
			// `fetch`, `FileReader` and a canvas. That is not missing tests, it is a
			// missing runtime, and no amount of Node testing can buy slack back here.
			//
			// The doctrine moved to the report that has a collector for every line it
			// counts: `scripts/coverage-gates.json`, checked by scripts/coverage-gate.mjs
			// against the Node suite and browser lane merged (`pnpm run coverage:gate`).
			// That is where a notch must clear its number by a full point, where ratchets
			// happen, and where the rule now *fails a build* rather than living in a
			// comment. See the header of scripts/coverage-gate.mjs for why prose was not
			// enough.
			thresholds: {
				// Raised 91 -> 92 once the table auto-pager landed: measured 93.21.
				statements: 92,
				// Raised 80 -> 81 once the text and chart definers landed: measured 82.79.
				// Ratchet upward only — if a change drops a number below its gate, that is a
				// finding to explain, never a gate to lower.
				branches: 81,
				// Left at 97. Dropping the `dist/browser.js` exclusion put
				// `src/runtime/browser.ts`'s 13 functions into this denominator with 1 of them
				// reachable from Node, so the Node-only number fell 98.33 -> 97.35. It reads
				// 97.77 now that the public accessors have tests
				// (test/regression/api/public-accessors.test.js), and 98.29 merged.
				functions: 97,
				// Raised 94 -> 95 once the zoom/background definers landed, measured 96.00 at
				// the time; the same exclusion drop took it to 95.67, and it reads 95.74 now.
				// 96.11 merged.
				lines: 95,
			},
		},
	},
})
