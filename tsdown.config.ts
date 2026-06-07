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
			core: 'src/core.ts',
			inspect: 'src/inspect.ts',
			node: 'src/node.ts',
			browser: 'src/browser.ts',
		},
	},
	{
		...shared,
		clean: false,
		deps: {
			alwaysBundle: ['jszip'],
			neverBundle: ['node:fs', 'node:https'],
			onlyBundle: false,
		},
		entry: {
			standalone: 'src/standalone.ts',
		},
		platform: 'browser',
	},
])
