#!/usr/bin/env node
/**
 * Bundle-size budget — how big the browser entry has got.
 *
 * Nothing measured shipped size before this. That was tolerable while browser support was
 * "works by construction"; it is not tolerable alongside a claim that the browser is a
 * first-class runtime, because a size promise nobody measures is a promise that quietly
 * stops being true. A single dependency pulled into the wrong chunk can double the entry
 * without failing one test in the repo.
 *
 *   node scripts/bundle-size-ratchet.mjs            # check (exit 1 over budget)
 *   node scripts/bundle-size-ratchet.mjs --freeze   # rewrite budget.json from dist/
 *   node scripts/bundle-size-ratchet.mjs --list     # per-chunk breakdown
 *
 * **What is measured.** For each budgeted entry, the transitive closure of its *relative*
 * imports — the files tsdown emits and the package ships — gzipped at level 9. Bare
 * specifiers are excluded because they are the consumer's dependencies, resolved and paid
 * for by the consumer's bundler: for the browser entry those are `fflate` and
 * (dynamically, on first font registration) `opentype.js`. Dynamic *relative* imports are
 * included — they ship in the package whether or not a consumer's bundler splits them.
 *
 * **This number is a proxy, not a download size.** `dist/` is unminified, and every real
 * browser consumer runs it through a bundler that minifies before serving. So the figure
 * here is comfortably larger than what anyone downloads. What makes it a useful gate is
 * that it moves when the code moves: a dependency reaching the browser entry, or a chunk
 * split going wrong, shows up as a step change. Do not quote it as "the browser bundle
 * is N kB"; docs/runtime-and-package-support.md says what it does mean.
 *
 * **How it differs from the raw-xml ratchet**, whose mechanics this otherwise copies.
 * That one fails when a count drops, because a lower number is the goal and banking it is
 * the point. Bytes are not like that: every real feature costs some, and a build whose
 * output moved 40 bytes is not a finding. So `--freeze` writes the measurement plus
 * `HEADROOM_PCT`, ordinary work passes inside it, and only a step change fails. The
 * mirror-image nag — re-freeze when an entry comes in `SLACK_PCT` under — exists so a
 * genuine win gets banked instead of silently becoming next year's headroom.
 */

import fs from 'node:fs'
import path from 'node:path'
import zlib from 'node:zlib'
import { ROOT } from './script-utils.mjs'

const DIST = path.join(ROOT, 'dist')
const BUDGET = path.join(ROOT, 'scripts', 'bundle-size-budget.json')

/**
 * Entries under budget, by their `dist/` file name.
 *
 * The browser entry is the one a consumer pays for over the wire, so it is the one with a
 * number. The Node entries are read off a local disk and the read/inspect entries are
 * opt-in subpaths; adding either is a line here plus a `--freeze`.
 */
const ENTRIES = ['browser.js']

/** Room `--freeze` leaves above the measurement, so ordinary work is not a re-freeze. */
const HEADROOM_PCT = 5

/** Re-freeze is only worth asking for when an entry comes in this far under budget. */
const SLACK_PCT = 15

/** Relative specifiers, static and dynamic. Bare ones are the consumer's to resolve. */
const RELATIVE_IMPORT = /(?:from\s*|import\s*\()\s*["'](\.[^"']*)["']/g

const argv = process.argv.slice(2)
const freeze = argv.includes('--freeze')
const list = argv.includes('--list')

/**
 * Every emitted file an entry pulls in, including itself.
 * @param {string} entry file name under `dist/`
 * @returns {string[]} file names, sorted
 */
function closureOf(entry) {
	const seen = new Set()
	const queue = [entry]
	while (queue.length) {
		const name = queue.shift()
		if (!name || seen.has(name)) continue
		const file = path.join(DIST, name)
		if (!fs.existsSync(file)) throw new Error(`dist/${name} is missing — run \`pnpm run build\` first`)
		seen.add(name)
		const text = fs.readFileSync(file, 'utf8')
		for (const [, specifier] of text.matchAll(RELATIVE_IMPORT)) {
			if (specifier) queue.push(path.posix.normalize(path.posix.join(path.posix.dirname(name), specifier)))
		}
	}
	return [...seen].sort()
}

/** @param {string} name @returns {number} gzipped bytes */
const gzipBytes = (name) => zlib.gzipSync(fs.readFileSync(path.join(DIST, name)), { level: 9 }).byteLength

/** @type {Map<string, {files: Array<{name: string, gzip: number}>, gzip: number}>} */
const measured = new Map()
for (const entry of ENTRIES) {
	const files = closureOf(entry).map((name) => ({ name, gzip: gzipBytes(name) }))
	measured.set(entry, { files, gzip: files.reduce((sum, file) => sum + file.gzip, 0) })
}

/** @param {number} bytes */
const kb = (bytes) => (bytes / 1024).toFixed(1) + ' kB'

if (list) {
	for (const [entry, { files, gzip }] of measured) {
		console.log(`${entry}: ${kb(gzip)} gzipped, ${files.length} file(s)`)
		for (const file of [...files].sort((a, b) => b.gzip - a.gzip))
			console.log(`  ${kb(file.gzip).padStart(9)}  ${file.name}`)
	}
	process.exit(0)
}

if (freeze) {
	/** @type {Record<string, number>} */
	const frozen = {}
	// Rounded up to a whole kB so the budget reads as a decision someone made rather than
	// as a build artifact copied into a file.
	for (const [entry, { gzip, files }] of measured) {
		frozen[entry] = Math.ceil((gzip * (1 + HEADROOM_PCT / 100)) / 1024) * 1024
		console.log(`bundle size: froze ${entry} at ${kb(frozen[entry])} (measured ${kb(gzip)}, ${files.length} file(s))`)
	}
	fs.writeFileSync(BUDGET, JSON.stringify(frozen, null, '\t') + '\n')
	process.exit(0)
}

/** @type {Record<string, number>} */
const budgetFile = JSON.parse(fs.readFileSync(BUDGET, 'utf8'))
const relBudget = path.relative(ROOT, BUDGET).split(path.sep).join('/')

const missing = [...measured.keys()].filter((entry) => typeof budgetFile[entry] !== 'number')
if (missing.length) {
	console.error(`bundle size FAILED — no budget for ${missing.join(', ')} in ${relBudget}.`)
	console.error('\n  pnpm run bundle-size:freeze')
	process.exit(1)
}

const checked = [...measured].map(([entry, { gzip, files }]) => ({
	entry,
	gzip,
	files,
	budget: budgetFile[entry] ?? 0,
}))
const over = checked.filter((row) => row.gzip > row.budget)
const under = checked.filter((row) => row.gzip < row.budget * (1 - SLACK_PCT / 100))

if (over.length) {
	console.error('bundle size FAILED — an entry grew past its budget:\n')
	for (const { entry, gzip, files, budget } of over) {
		console.error(`  ${entry}: ${kb(gzip)} gzipped (budget ${kb(budget)}), ${files.length} file(s)`)
		for (const file of [...files].sort((a, b) => b.gzip - a.gzip).slice(0, 5))
			console.error(`    ${kb(file.gzip).padStart(9)}  ${file.name}`)
	}
	console.error('\nRun `pnpm run bundle-size:list` for the full breakdown. If the growth is')
	console.error(`intended, raise the number in ${relBudget} in the same commit and say what`)
	console.error('bought the bytes; if it is not, something reached the browser entry that should not.')
	process.exit(1)
}

if (under.length) {
	console.log(`bundle size: ${SLACK_PCT}%+ under budget — bank it by lowering ${relBudget}:\n`)
	for (const { entry, gzip, budget } of under) console.log(`  ${entry}: ${kb(budget)} -> ${kb(gzip)}`)
	console.log('\n  pnpm run bundle-size:freeze')
	process.exit(1)
}

for (const { entry, gzip, files, budget } of checked)
	console.log(`bundle size: ok — ${entry} ${kb(gzip)} gzipped (budget ${kb(budget)}), ${files.length} file(s)`)
