#!/usr/bin/env node
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, run } from './script-utils.mjs'

const ALL_TARGETS = ['node', 'vite']
const requestedTargets = process.argv.slice(2)
const targets = requestedTargets.length > 0 ? requestedTargets : ALL_TARGETS

for (const target of targets) {
	if (!ALL_TARGETS.includes(target)) {
		throw new Error('unknown demo smoke target "' + target + '"; expected node or vite')
	}
}

async function fileStat(file) {
	try {
		return await fs.stat(file)
	} catch {
		return null
	}
}

async function assertGeneratedPptx(file, before) {
	const after = await fileStat(file)
	if (!after?.isFile()) throw new Error('expected demo output file: ' + path.relative(ROOT, file))
	if (after.size === 0) throw new Error('demo output file is empty: ' + path.relative(ROOT, file))
	if (before && after.mtimeMs <= before.mtimeMs) {
		throw new Error('demo output file was not refreshed: ' + path.relative(ROOT, file))
	}
}

async function smokeNodeDemo() {
	const outputFile = path.join(ROOT, 'demos', 'node', 'output', 'PptxGenJS_Demo_Text.pptx')
	const before = await fileStat(outputFile)
	await run('pnpm', ['--dir', 'demos/node', 'run', 'demo-text'])
	await assertGeneratedPptx(outputFile, before)
}

async function smokeViteDemo() {
	await run('pnpm', ['--dir', 'demos/vite-demo', 'run', 'build'])
}

await run('pnpm', ['run', 'build'])

for (const target of targets) {
	if (target === 'node') await smokeNodeDemo()
	else if (target === 'vite') await smokeViteDemo()
}

console.log('Demo smoke tests passed: ' + targets.join(', '))
