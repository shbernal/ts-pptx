#!/usr/bin/env node
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT } from './rollup-options.mjs'
import { run } from './script-utils.mjs'

const ALL_TARGETS = ['node', 'vite']
const requestedTargets = process.argv.slice(2)
const targets = requestedTargets.length > 0 ? requestedTargets : ALL_TARGETS

for (const target of targets) {
	if (!ALL_TARGETS.includes(target)) {
		throw new Error('unknown demo smoke target "' + target + '"; expected node or vite')
	}
}

async function pptxFiles(dir) {
	try {
		return new Set((await fs.readdir(dir)).filter((file) => file.endsWith('.pptx')))
	} catch {
		return new Set()
	}
}

async function removeNewPptxFiles(dir, before) {
	for (const file of await pptxFiles(dir)) {
		if (!before.has(file)) await fs.rm(path.join(dir, file), { force: true })
	}
}

async function smokeNodeDemo() {
	const demoDir = path.join(ROOT, 'demos', 'node')
	const before = await pptxFiles(demoDir)
	try {
		await run('pnpm', ['--dir', 'demos/node', 'run', 'demo-text'])
	} finally {
		await removeNewPptxFiles(demoDir, before)
	}
}

async function smokeViteDemo() {
	await run('pnpm', ['--dir', 'demos/vite-demo', 'run', 'build'])
}

await run('pnpm', ['run', 'build:dist'])
await run('pnpm', ['run', 'types:build'])

for (const target of targets) {
	if (target === 'node') await smokeNodeDemo()
	else if (target === 'vite') await smokeViteDemo()
}

console.log('Demo smoke tests passed: ' + targets.join(', '))
