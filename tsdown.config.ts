import { defineConfig } from 'tsdown'

const shared = {
	dts: {
		sourcemap: true,
	},
	deps: {
		neverBundle: ['node:fs', 'node:https'],
	},
	fixedExtension: false,
	format: 'esm',
	sourcemap: true,
	target: 'es2024',
	treeshake: true,
} as const

export default defineConfig([
	{
		...shared,
		clean: true,
		entry: {
			index: 'src/index.ts',
			inspect: 'src/inspect.ts',
			measure: 'src/measure.ts',
			read: 'src/read.ts',
			math: 'src/math.ts',
			zip: 'src/zip.ts',
			node: 'src/node.ts',
			browser: 'src/browser.ts',
		},
	},
])
