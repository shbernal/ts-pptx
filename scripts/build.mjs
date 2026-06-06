#!/usr/bin/env node
import { rollup } from 'rollup'
import { minify } from 'terser'
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, createRollupInputOptions, readPackageJson } from './rollup-options.mjs'

const BLD_DIR = path.join(ROOT, 'src', 'bld')
const DIST_DIR = path.join(ROOT, 'dist')
const LIBS_DIR = path.join(ROOT, 'libs')
const DEMO_BROWSER_DIR = path.join(ROOT, 'demos', 'browser', 'js')
const DEMO_NODE_DIST_DIR = path.join(ROOT, 'demos', 'node', 'node_modules', 'pptxgenjs', 'dist')
const DEMO_VITE_PKG_DIR = path.join(ROOT, 'demos', 'vite-demo', 'node_modules', 'pptxgenjs')

const BLD_IIFE = path.join(BLD_DIR, 'pptxgen.iife.js')
const BLD_ESM = path.join(BLD_DIR, 'pptxgen.js')
const BLD_LEGACY_FILES = [path.join(BLD_DIR, 'pptxgen.cjs.js'), path.join(BLD_DIR, 'pptxgen.es.js')]

const DIST_ESM = path.join(DIST_DIR, 'pptxgen.js')
const DIST_MIN = path.join(DIST_DIR, 'pptxgen.min.js')
const DIST_BUNDLE = path.join(DIST_DIR, 'pptxgen.bundle.js')
const DIST_LEGACY_FILES = [path.join(DIST_DIR, 'pptxgen.cjs.js'), path.join(DIST_DIR, 'pptxgen.es.js')]

process.chdir(ROOT)

const pkg = await readPackageJson()

function header() {
	return '/* PptxGenJS ' + pkg.version + ' @ ' + new Date().toISOString() + ' */\n'
}

function rel(file) {
	return path.relative(ROOT, file)
}

async function exists(file) {
	try {
		await fs.access(file)
		return true
	} catch {
		return false
	}
}

async function assertFile(file, hint) {
	if (!(await exists(file))) {
		throw new Error('expected ' + rel(file) + (hint ? ' ' + hint : ''))
	}
}

async function removeFiles(files) {
	await Promise.all(files.map((file) => fs.rm(file, { force: true })))
}

async function createRollupBundle() {
	return rollup(createRollupInputOptions(pkg))
}

async function buildBld() {
	await fs.mkdir(BLD_DIR, { recursive: true })
	await removeFiles(BLD_LEGACY_FILES)
	const bundle = await createRollupBundle()
	try {
		await bundle.write({
			file: BLD_IIFE,
			format: 'iife',
			name: 'PptxGenJS',
			globals: { jszip: 'JSZip' },
		})
		await bundle.write({
			file: BLD_ESM,
			format: 'es',
		})
	} finally {
		await bundle.close()
	}
	logFiles('Built', [BLD_IIFE, BLD_ESM])
}

async function writeDistEntries() {
	await buildBld()
	await fs.mkdir(DIST_DIR, { recursive: true })
	await removeFiles(DIST_LEGACY_FILES)
	await fs.writeFile(DIST_ESM, header() + (await fs.readFile(BLD_ESM, 'utf8')))
	logFiles('Wrote', [DIST_ESM])
}

async function libFiles() {
	const files = (await fs.readdir(LIBS_DIR))
		.filter((file) => file.endsWith('.js'))
		.sort()
		.map((file) => path.join(LIBS_DIR, file))
	if (files.length === 0) {
		throw new Error('expected at least one JS file under ' + rel(LIBS_DIR))
	}
	return files
}

async function concatScripts(files) {
	const chunks = []
	for (const file of files) {
		chunks.push(await fs.readFile(file, 'utf8'))
		chunks.push('\n;\n')
	}
	return chunks.join('')
}

async function minifyToFile(sourceCode, outFile) {
	const fileName = path.basename(outFile)
	const result = await minify(sourceCode, {
		compress: true,
		mangle: true,
		format: {
			comments: false,
			preamble: header().trimEnd(),
		},
		sourceMap: {
			filename: fileName,
			url: fileName + '.map',
		},
	})
	if (!result.code) throw new Error('terser did not produce code for ' + rel(outFile))
	if (!result.map) throw new Error('terser did not produce a source map for ' + rel(outFile))
	await fs.mkdir(path.dirname(outFile), { recursive: true })
	await fs.writeFile(outFile, result.code + '\n')
	await fs.writeFile(outFile + '.map', result.map + '\n')
}

async function buildDist() {
	await writeDistEntries()
	await minifyToFile(await fs.readFile(BLD_IIFE, 'utf8'), DIST_MIN)
	await minifyToFile(await concatScripts([...(await libFiles()), BLD_IIFE]), DIST_BUNDLE)
	logFiles('Wrote', [DIST_MIN, DIST_MIN + '.map', DIST_BUNDLE, DIST_BUNDLE + '.map'])
}

async function copyBrowserBundle() {
	await assertFile(DIST_BUNDLE, 'before copying browser bundle')
	await assertFile(DIST_BUNDLE + '.map', 'before copying browser bundle map')
	await fs.mkdir(DEMO_BROWSER_DIR, { recursive: true })
	await fs.copyFile(DIST_BUNDLE, path.join(DEMO_BROWSER_DIR, 'pptxgen.bundle.js'))
	await fs.copyFile(DIST_BUNDLE + '.map', path.join(DEMO_BROWSER_DIR, 'pptxgen.bundle.js.map'))
	logFiles('Copied', [
		path.join(DEMO_BROWSER_DIR, 'pptxgen.bundle.js'),
		path.join(DEMO_BROWSER_DIR, 'pptxgen.bundle.js.map'),
	])
}

async function copyNodeDemo() {
	await assertFile(DIST_ESM, 'before copying Node demo ESM entry')
	await fs.mkdir(DEMO_NODE_DIST_DIR, { recursive: true })
	await removeFiles([path.join(DEMO_NODE_DIST_DIR, 'pptxgen.cjs.js'), path.join(DEMO_NODE_DIST_DIR, 'pptxgen.es.js')])
	await fs.copyFile(DIST_ESM, path.join(DEMO_NODE_DIST_DIR, 'pptxgen.js'))
	logFiles('Copied', [path.join(DEMO_NODE_DIST_DIR, 'pptxgen.js')])
}

async function copyViteDemo(opts = {}) {
	const copyCode = opts.code !== false
	const copyTypes = opts.types !== false
	const destDist = path.join(DEMO_VITE_PKG_DIR, 'dist')
	const destTypes = path.join(DEMO_VITE_PKG_DIR, 'types')
	const copied = []
	if (copyCode) {
		await assertFile(DIST_ESM, 'before copying Vite demo ESM entry')
		await fs.mkdir(destDist, { recursive: true })
		await removeFiles([path.join(destDist, 'pptxgen.es.js')])
		await fs.copyFile(DIST_ESM, path.join(destDist, 'pptxgen.js'))
		copied.push(path.join(destDist, 'pptxgen.js'))
	}
	if (copyTypes) {
		const typesFile = path.join(ROOT, 'types', 'pptxgen.d.ts')
		await assertFile(typesFile, 'before copying Vite demo types')
		await fs.mkdir(destTypes, { recursive: true })
		await copyDir(path.join(ROOT, 'types'), destTypes)
		copied.push(destTypes)
	}
	logFiles('Copied', copied)
}

async function copyDir(from, to) {
	await fs.rm(to, { recursive: true, force: true })
	await fs.mkdir(to, { recursive: true })
	const entries = await fs.readdir(from, { withFileTypes: true })
	for (const entry of entries) {
		const src = path.join(from, entry.name)
		const dest = path.join(to, entry.name)
		if (entry.isDirectory()) await copyDir(src, dest)
		else if (entry.isFile()) await fs.copyFile(src, dest)
	}
}

async function ship() {
	await buildDist()
	await copyBrowserBundle()
	await copyViteDemo()
	await copyNodeDemo()
}

function logFiles(verb, files) {
	for (const file of files) {
		console.log(verb + ' ' + rel(file))
	}
}

async function main() {
	const command = process.argv[2] || 'build'
	switch (command) {
		case 'build':
			await buildBld()
			break
		case 'dist':
			await buildDist()
			break
		case 'ship':
			await ship()
			break
		case 'copy:browser':
			await copyBrowserBundle()
			break
		case 'copy:node':
			await copyNodeDemo()
			break
		case 'copy:vite':
			await copyViteDemo()
			break
		case 'copy:vite-types':
			await copyViteDemo({ code: false })
			break
		default:
			throw new Error(
				'unknown build command "' +
					command +
					'"; expected build, dist, ship, copy:browser, copy:node, copy:vite, or copy:vite-types'
			)
	}
}

main().catch((err) => {
	console.error(err && err.stack ? err.stack : err)
	process.exit(1)
})
