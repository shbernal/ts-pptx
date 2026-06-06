#!/usr/bin/env node
import { rollup } from 'rollup'
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, createRollupInputOptions, readPackageJson } from './rollup-options.mjs'

const BLD_DIR = path.join(ROOT, 'src', 'bld')
const DIST_DIR = path.join(ROOT, 'dist')

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
		default:
			throw new Error('unknown build command "' + command + '"; expected build or dist')
	}
}

main().catch((err) => {
	console.error(err && err.stack ? err.stack : err)
	process.exit(1)
})
