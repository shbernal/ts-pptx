import { defineConfig } from 'vitest/config'

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
			include: ['dist/**'],
			reporter: ['text-summary', 'text', 'html'],
			thresholds: {
				statements: 80,
				branches: 66,
				functions: 88,
				lines: 84,
			},
		},
	},
})
