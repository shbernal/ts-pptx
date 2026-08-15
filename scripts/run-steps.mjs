#!/usr/bin/env node
// Run a list of package.json scripts as one sequence, without a package-manager
// process per step.
//
// Why this exists: the gates are composites, and spelling a composite as
// `pnpm run a && pnpm run b && …` charges a full pnpm startup for every step.
// Measured in this repo, that wrapper is ~0.7–1.3s regardless of the work it
// wraps — it is pnpm's own CLI boot plus a second `node`, not dependency
// resolution, so it does not shrink with the size of the job. `verify` expanded
// to 13 such invocations: ~13s of an ~85s gate spent starting package managers.
//
// The alternative — inlining each step's real command into the composite —
// would create a second, drifting copy of every gate, which is precisely what
// AGENTS.md warns against. So this runner keeps package.json as the single
// definition of what each step *is* and only removes the layer that re-launches
// a package manager to read it.
//
// Usage:
//   node scripts/run-steps.mjs <script-name>…     run each, in order, stop on failure
//   node scripts/run-steps.mjs --list <name>…     print the expansion, run nothing

import { spawn } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { ROOT, runCli } from './script-utils.mjs'

/** @type {Record<string, string>} */
const scripts = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).scripts

/** This runner's own entry, as it appears in a package.json script body. */
const SELF = /^node\s+scripts\/run-steps\.mjs\s+(.*)$/

/**
 * Flatten a script name into the leaf shell commands it ultimately runs.
 *
 * Two forms are followed rather than executed, because both are indirection this
 * runner exists to remove:
 *   - `pnpm run <name>`, when `<name>` is one of our own scripts;
 *   - `node scripts/run-steps.mjs <names…>`, i.e. a composite built on this runner.
 *
 * Everything else is a leaf and is handed to a shell as written. Note the
 * deliberate narrowness: `pnpm --dir demos/showcases run build` is NOT followed,
 * because it names a script in a different workspace that this package.json does
 * not define.
 * @param {string} name - a key in package.json `scripts`
 * @param {string[]} trail - names already being expanded, for cycle detection
 * @returns {{step: string, command: string}[]}
 */
function expand(name, trail = []) {
	const body = scripts[name]
	if (body === undefined) throw new Error(`no such script: ${name}\n(known: ${Object.keys(scripts).join(', ')})`)
	if (trail.includes(name)) throw new Error(`script cycle: ${[...trail, name].join(' -> ')}`)

	const out = []
	for (const segment of body.split('&&').map((s) => s.trim())) {
		if (segment === '') continue
		const selfMatch = SELF.exec(segment)
		if (selfMatch) {
			for (const child of (selfMatch[1] ?? '').split(/\s+/).filter(Boolean))
				out.push(...expand(child, [...trail, name]))
			continue
		}
		const runName = /^pnpm\s+run\s+(\S+)$/.exec(segment)?.[1]
		if (runName !== undefined && scripts[runName] !== undefined) {
			out.push(...expand(runName, [...trail, name]))
			continue
		}
		out.push({ step: name, command: segment })
	}
	return out
}

/**
 * Run one leaf command through a shell, with the local bin directory on PATH.
 *
 * `shell: true` is what lets a bare bin name (`tsc`, `vitest`, `vitepress`) resolve
 * the same way it does under a package manager, including to a `.CMD` shim on
 * Windows, which `spawn` refuses to exec directly.
 * @param {string} command
 * @returns {Promise<number>}
 */
function runLeaf(command) {
	const binDir = path.join(ROOT, 'node_modules', '.bin')
	return new Promise((resolve, reject) => {
		const child = spawn(command, {
			cwd: ROOT,
			shell: true,
			stdio: 'inherit',
			env: { ...process.env, PATH: binDir + path.delimiter + (process.env.PATH ?? '') },
		})
		child.on('error', reject)
		child.on('exit', (code, signal) => resolve(signal ? 1 : (code ?? 1)))
	})
}

await runCli(async () => {
	const argv = process.argv.slice(2)
	const listOnly = argv[0] === '--list'
	const names = listOnly ? argv.slice(1) : argv
	if (names.length === 0) {
		console.error('usage: node scripts/run-steps.mjs [--list] <script-name>…')
		return 1
	}

	const steps = names.flatMap((n) => expand(n))

	if (listOnly) {
		for (const { step, command } of steps) console.log(step.padEnd(24), command)
		return 0
	}

	const started = process.hrtime.bigint()
	/** @type {{step: string, ms: number}[]} */
	const timings = []
	// Composites overlap: `verify:full` is `verify` plus `docs:build`, and both
	// reach `docs:api` (6.5s) and `ensure-dist`. Every step in these gates is a
	// check or a regenerator whose second consecutive run cannot say anything the
	// first did not, so an exact command repeat inside ONE invocation is skipped.
	//
	// The assumption is that no step here mutates another's inputs — true today,
	// and the skip is logged rather than silent so that a step which starts doing
	// so is visible at the point it would be wrongly elided. Order is otherwise
	// untouched: this drops repeats, it never reorders or parallelises.
	const alreadyRun = new Set()
	for (const { step, command } of steps) {
		if (alreadyRun.has(command)) {
			console.log(`· ${step} — skipped, already run in this invocation (${command})`)
			continue
		}
		alreadyRun.add(command)
		const stepStart = process.hrtime.bigint()
		const code = await runLeaf(command)
		const ms = Number(process.hrtime.bigint() - stepStart) / 1e6
		timings.push({ step, ms })
		if (code !== 0) {
			console.error(`\n✗ ${step} failed (exit ${code})\n  ${command}`)
			return code
		}
	}

	// A per-step breakdown, so the cost of a gate is visible where it is paid
	// rather than having to be measured from outside every time it grows.
	const total = Number(process.hrtime.bigint() - started) / 1e6
	const merged = new Map()
	for (const { step, ms } of timings) merged.set(step, (merged.get(step) ?? 0) + ms)
	const summary = [...merged].map(([step, ms]) => `${step} ${(ms / 1000).toFixed(1)}s`).join(', ')
	console.log(`\n✓ ${names.join(' ')} — ${(total / 1000).toFixed(1)}s (${summary})`)
	return 0
})
