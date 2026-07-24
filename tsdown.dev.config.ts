import { defineConfig } from 'tsdown'

// Fast inner-loop build for the edit → test cycle. Unlike the production
// `tsdown.config.ts`, this config skips the two most expensive steps:
//   1. `.d.ts` emit (`dts: false`)          — type declarations aren't needed to run tests
//   2. the browser `standalone` IIFE bundle — never exercised by the Node suite
//
// It still emits every Node-side entry the regression/read/measure suites import
// from `dist/` (node, index, core, inspect, measure, read, math, zip), so a
// running `pnpm run watch:dev` keeps `dist/` current for the full `test/regression`
// (and sibling) suites — the pieces a one-assertion change actually touches.
//
// Use it via `pnpm run watch:dev` in one terminal + `pnpm run test:watch:fast`
// in another. See docs/testing.md "Fast inner loop".
export default defineConfig({
	dts: false,
	deps: {
		neverBundle: ['node:fs', 'node:https'],
	},
	fixedExtension: false,
	format: 'esm',
	sourcemap: true,
	target: 'es2024',
	treeshake: true,
	clean: true,
	entry: {
		index: 'src/index.ts',
		inspect: 'src/inspect.ts',
		measure: 'src/measure.ts',
		read: 'src/read.ts',
		math: 'src/math.ts',
		zip: 'src/zip.ts',
		node: 'src/node.ts',
	},
})
