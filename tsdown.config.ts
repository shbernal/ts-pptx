import { defineConfig, type UserConfig } from 'tsdown'

// Annotated rather than `as const`: the annotation contextually types the string
// literals (`format: 'esm'`) without also making the arrays readonly, which
// `deps.neverBundle` rejects.
const shared: UserConfig = {
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
}

export default defineConfig([
	{
		...shared,
		clean: true,
		entry: {
			index: 'src/index.ts',
			inspect: 'src/inspect.ts',
			measure: 'src/measure.ts',
			read: 'src/read.ts',
			script: 'src/script.ts',
			math: 'src/math.ts',
			zip: 'src/zip.ts',
			node: 'src/node.ts',
			browser: 'src/browser.ts',
		},
	},
])
