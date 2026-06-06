#!/usr/bin/env node
import { rollup } from 'rollup'
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, createRollupInputOptions, readPackageJson } from './rollup-options.mjs'

const BLD_DIR = path.join(ROOT, 'src', 'bld')
const DIST_DIR = path.join(ROOT, 'dist')
const DEMO_NODE_DIST_DIR = path.join(ROOT, 'demos', 'node', 'node_modules', 'pptxgenjs', 'dist')
const DEMO_VITE_PKG_DIR = path.join(ROOT, 'demos', 'vite-demo', 'node_modules', 'pptxgenjs')

const BLD_ESM = path.join(BLD_DIR, 'pptxgen.js')

const DIST_ESM = path.join(DIST_DIR, 'pptxgen.js')
const DIST_STALE_FILES = [
	path.join(DIST_DIR, 'pptxgen.cjs.js'),
	path.join(DIST_DIR, 'pptxgen.es.js'),
	path.join(DIST_DIR, 'pptxgen.min.js'),
	path.join(DIST_DIR, 'pptxgen.min.js.map'),
	path.join(DIST_DIR, 'pptxgen.bundle.js'),
	path.join(DIST_DIR, 'pptxgen.bundle.js.map'),
]

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
	await fs.rm(BLD_DIR, { recursive: true, force: true })
	await fs.mkdir(BLD_DIR, { recursive: true })
	const bundle = await createRollupBundle()
	try {
		await bundle.write({
			file: BLD_ESM,
			format: 'es',
		})
	} finally {
		await bundle.close()
	}
	logFiles('Built', [BLD_ESM])
}

async function writeDistEntries() {
	await buildBld()
	await fs.mkdir(DIST_DIR, { recursive: true })
	await removeFiles(DIST_STALE_FILES)
	await fs.writeFile(DIST_ESM, header() + (await fs.readFile(BLD_ESM, 'utf8')))
	logFiles('Wrote', [DIST_ESM])
}

async function buildDist() {
	await writeDistEntries()
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
				'unknown build command "' + command + '"; expected build, dist, ship, copy:node, copy:vite, or copy:vite-types'
			)
	}
}

main().catch((err) => {
	console.error(err && err.stack ? err.stack : err)
	process.exit(1)
})
