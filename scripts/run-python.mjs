#!/usr/bin/env node
// Resolves a usable Python 3 and runs a repo script with it.
//
// The docs:* scripts used to hardcode `python3`, which breaks on Windows: the
// Microsoft Store ships a `python3.exe` app-execution alias that is not an
// interpreter and exits 9009. So probe candidates in a platform-aware order and
// only accept one that actually reports a Python 3 version.
import { spawnSync } from 'node:child_process'
import path from 'node:path'

import { ROOT } from './script-utils.mjs'

const MIN_MAJOR = 3
const MIN_MINOR = 9
const MIN_VERSION = `${MIN_MAJOR}.${MIN_MINOR}`

const PROBE = 'import sys; print("%d.%d" % sys.version_info[:2])'

// TSPPTX_PYTHON is authoritative — never silently fall back past a typo in it.
// PYTHON is a shared convention (node-gyp and friends), so treat it as a hint.
const override = process.env.TSPPTX_PYTHON

function candidates() {
	if (override) return [{ command: override, args: [] }]
	const list = []
	if (process.env.PYTHON) list.push({ command: process.env.PYTHON, args: [] })
	if (process.platform === 'win32') {
		// `py -3` first: the launcher is a real interpreter shim, unlike the Store alias.
		list.push({ command: 'py', args: ['-3'] }, { command: 'python', args: [] }, { command: 'python3', args: [] })
	} else {
		list.push({ command: 'python3', args: [] }, { command: 'python', args: [] })
	}
	return list
}

function probe(candidate) {
	let result
	try {
		result = spawnSync(candidate.command, [...candidate.args, '-c', PROBE], {
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'pipe'],
			windowsHide: true,
		})
	} catch {
		return null
	}
	if (result.error || result.status !== 0) return null
	const match = /^(\d+)\.(\d+)$/m.exec((result.stdout || '').trim())
	if (!match) return null
	const major = Number(match[1])
	const minor = Number(match[2])
	if (major !== MIN_MAJOR || minor < MIN_MINOR) return null
	return { ...candidate, version: `${major}.${minor}` }
}

function resolvePython() {
	for (const candidate of candidates()) {
		const found = probe(candidate)
		if (found) return found
	}
	return null
}

const [script, ...scriptArgs] = process.argv.slice(2)
if (!script) {
	console.error('usage: node scripts/run-python.mjs <script.py> [args...]')
	process.exit(2)
}

const python = resolvePython()
if (!python) {
	const tried = candidates()
		.map((candidate) => [candidate.command, ...candidate.args].join(' '))
		.join(', ')
	console.error(
		override
			? `TSPPTX_PYTHON is set to "${override}", which is not a working Python ${MIN_VERSION}+ interpreter.`
			: `No Python ${MIN_VERSION}+ interpreter found (tried: ${tried}).\n` +
					'Install Python 3 from https://www.python.org/downloads/ (not the Microsoft Store alias),\n' +
					'or point TSPPTX_PYTHON at an interpreter, e.g. TSPPTX_PYTHON=C:\\Python313\\python.exe'
	)
	process.exit(1)
}

const target = path.isAbsolute(script) ? script : path.join(ROOT, script)
const result = spawnSync(python.command, [...python.args, target, ...scriptArgs], {
	cwd: process.cwd(),
	stdio: 'inherit',
	windowsHide: true,
})
if (result.error) {
	console.error(result.error.message)
	process.exit(1)
}
process.exit(result.status ?? 1)
