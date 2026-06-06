import commonjs from '@rollup/plugin-commonjs'
import resolve from '@rollup/plugin-node-resolve'
import typescript from '@rollup/plugin-typescript'
import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

export async function readPackageJson() {
	return JSON.parse(await fs.readFile(path.join(ROOT, 'package.json'), 'utf8'))
}

export function createExternalPredicate(pkg) {
	const externalPackages = new Set([...Object.keys(pkg.dependencies || {}), ...Object.keys(pkg.peerDependencies || {})])

	return (id) => {
		if (id.startsWith('node:')) return true
		for (const dep of externalPackages) {
			if (id === dep || id.startsWith(dep + '/')) return true
		}
		return false
	}
}

export function createRollupPlugins() {
	return [
		resolve({ preferBuiltins: true }),
		commonjs(),
		typescript({
			tsconfig: path.join(ROOT, 'tsconfig.build.json'),
			compilerOptions: {
				declaration: false,
				declarationDir: undefined,
				outDir: undefined,
				sourceMap: false,
			},
		}),
	]
}

export function createRollupInputOptions(pkg) {
	return {
		input: 'src/pptxgen.ts',
		external: createExternalPredicate(pkg),
		plugins: createRollupPlugins(),
	}
}

export async function createRollupConfig() {
	const pkg = await readPackageJson()
	return {
		...createRollupInputOptions(pkg),
		output: [
			{
				file: './src/bld/pptxgen.iife.js',
				format: 'iife',
				name: 'PptxGenJS',
				globals: { jszip: 'JSZip' },
			},
			{ file: './src/bld/pptxgen.cjs.js', format: 'cjs', exports: 'default' },
			{ file: './src/bld/pptxgen.es.js', format: 'es' },
		],
	}
}
