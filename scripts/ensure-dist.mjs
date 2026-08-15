#!/usr/bin/env node
// Freshness guard for `dist/`. Rebuilds only when the build inputs are newer
// than the build outputs; otherwise it is a ~50ms no-op.
//
// Why this exists: the test suite runs against the built package (tests import
// from `dist/`, not `src/`), so every test script used to front-load an
// unconditional `pnpm run build`. That cost 3.3s on every invocation and drove a
// parallel set of `:fast` twin scripts that skipped the build and silently
// tested stale output when `dist/` was not current. One guard replaces both:
// aggregates and individual scripts call it, and nobody prefixes
// `pnpm run build &&` any more.
//
// Accuracy bias: this compares mtimes, which is imprecise in one harmless
// direction and must never be imprecise in the other.
//   - Acceptable: a `git checkout` / `git stash` rewrites source mtimes and
//     triggers a needless 3.3s rebuild.
//   - Not acceptable: judging a stale `dist/` fresh, which would run the whole
//     suite against old code. So the input set includes the build configs and
//     the lockfile (a dependency bump changes the bundle without touching
//     `src/`), and the output check requires the `.d.ts` as well as the `.js` —
//     a build interrupted after emitting JS but before types is not "fresh".
//
// Usage:
//   node scripts/ensure-dist.mjs                rebuild if stale
//   node scripts/ensure-dist.mjs --check        fail if stale, never build
//                                               (for CI legs that want the build to
//                                               be an explicit, separately-timed step)
//   node scripts/ensure-dist.mjs --if-missing   build only when dist/ is absent
//                                               (for `prepare` — see the flag's comment)

import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, parseCliOrExit, run } from './script-utils.mjs'

const INPUT_FILES = ['tsdown.config.ts', 'tsconfig.base.json', 'tsconfig.json', 'package.json', 'pnpm-lock.yaml']
const INPUT_DIR = 'src'
const OUTPUT_FILES = ['dist/index.js', 'dist/index.d.ts']

/**
 * Newest mtime under `dir`, recursively. Returns 0 if the tree is absent.
 * @param {string} dir
 * @returns {Promise<number>}
 */
async function newestMtimeIn(dir) {
	let newest = 0
	let entries
	try {
		entries = await fs.readdir(dir, { withFileTypes: true })
	} catch {
		return 0
	}
	for (const entry of entries) {
		const full = path.join(dir, entry.name)
		if (entry.isDirectory()) {
			newest = Math.max(newest, await newestMtimeIn(full))
		} else if (entry.isFile()) {
			const { mtimeMs } = await fs.stat(full)
			newest = Math.max(newest, mtimeMs)
		}
	}
	return newest
}

/**
 * Mtime of one file, or 0 when it does not exist.
 * @param {string} file
 * @returns {Promise<number>}
 */
async function mtimeOf(file) {
	try {
		return (await fs.stat(file)).mtimeMs
	} catch {
		return 0
	}
}

/** @returns {Promise<string[]>} the `OUTPUT_FILES` that are not on disk at all. */
async function missingOutputs() {
	const outputs = await Promise.all(OUTPUT_FILES.map((f) => mtimeOf(path.join(ROOT, f))))
	return OUTPUT_FILES.filter((_, i) => outputs[i] === 0)
}

/** @returns {Promise<string | null>} why `dist/` is stale, or null if it is current. */
async function stale() {
	const outputs = await Promise.all(OUTPUT_FILES.map((f) => mtimeOf(path.join(ROOT, f))))
	const missing = OUTPUT_FILES.filter((_, i) => outputs[i] === 0)
	if (missing.length > 0) return 'missing build output: ' + missing.join(', ')

	const oldestOutput = Math.min(...outputs)
	const inputMtimes = await Promise.all([
		newestMtimeIn(path.join(ROOT, INPUT_DIR)),
		...INPUT_FILES.map((f) => mtimeOf(path.join(ROOT, f))),
	])
	const newestInput = Math.max(...inputMtimes)
	if (newestInput > oldestOutput) {
		const which = [INPUT_DIR + '/', ...INPUT_FILES][inputMtimes.indexOf(newestInput)]
		return 'build input is newer than dist/: ' + which
	}
	return null
}

/**
 * Run the `build` script through whichever package manager is running this script.
 *
 * Not a hardcoded `pnpm`, because of `prepare`. This repo is a pnpm repo, but a consumer's
 * `npm i github:shbernal/ts-pptx#<sha>` runs `prepare` in *their* toolchain — and `pnpm` is
 * declared here only as `packageManager`, never installed as a dependency, so `run()` finds
 * no local bin and falls through to a `pnpm.cmd` shim that a plain-npm consumer does not
 * have. `npm_execpath` is set by npm, pnpm and yarn alike and points at the JS entry of the
 * one actually in use, so this asks the right one and keeps the `build` script — rather than
 * a transcribed `tsdown` call — as the single definition of what a build is.
 */
function runBuild() {
	const execpath = process.env.npm_execpath
	// Only a JS entry is spawnable this way; a `.cmd`/`.ps1` shim would need a shell.
	if (execpath && /\.[cm]?js$/.test(execpath)) return run(process.execPath, [execpath, 'run', 'build'])
	return run('pnpm', ['run', 'build'])
}

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: `Build \`dist/\` if it is out of date with respect to its inputs.

  node scripts/ensure-dist.mjs                build when stale
  node scripts/ensure-dist.mjs --check        report staleness, never build (exit 1 if stale)
  node scripts/ensure-dist.mjs --if-missing   build only when there is no dist/ at all

Options:
  --check       fail instead of building — for CI, where a stale dist/ is a mistake
  --if-missing  build an absent dist/, but leave a stale one alone — for \`prepare\`
  -h, --help    show this message`,
	options: { check: { type: 'boolean', default: false }, 'if-missing': { type: 'boolean', default: false } },
})

// `--if-missing` asks a different question from the freshness check: not "is this build
// current?" but "is there a build at all?".
//
// It exists for the `prepare` script, which runs in two unrelated places. In this repo it
// runs on every install, where rebuilding a *stale* dist/ would be wrong — `pnpm run build`
// and `pnpm run watch` would then build twice, and every other script already front-loads
// its own unconditional `ensure-dist`. In a consumer's `npm i github:shbernal/ts-pptx#<sha>`
// it is the only build that will ever run: `dist/` is gitignored, so the checkout npm packs
// has no build output at all unless `prepare` produces one. Absent means build; stale is
// somebody else's question.
if (values['if-missing'] && (await missingOutputs()).length === 0) process.exit(0)

const reason = await stale()
if (reason === null) process.exit(0)

if (values.check) {
	console.error('dist/ is not current (' + reason + '). Run: pnpm run build')
	process.exit(1)
}

console.log('dist/ is not current (' + reason + ') — building')
await runBuild()
