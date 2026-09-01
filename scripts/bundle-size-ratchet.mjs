#!/usr/bin/env node
/**
 * Bundle-size budget — how big each published entry point has got.
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
 * imports — the files tsdown emits and the package ships — minified, then gzipped at
 * level 9. Bare specifiers are excluded because they are the consumer's dependencies,
 * resolved and paid for by the consumer's bundler: for the browser entry those are
 * `fflate` and (dynamically, on first font registration) `opentype.js`. Dynamic *relative*
 * imports are included — they ship in the package whether or not a consumer's bundler
 * splits them.
 *
 * **Why minify first, when `dist/` ships unminified.** Because the alternative measures
 * documentation. Almost half of `dist/` by weight is doc comments, and they do not survive
 * any consumer's build — every real browser consumer runs this through a bundler that
 * minifies before serving. Gating on the unminified bytes therefore charges a commit for
 * prose and credits it for deletions of prose, and that is not a hypothetical: over the
 * twenty-three commits from v3.7.0 to 147951de, the "state it once" refactors *removed*
 * 14.5 kB of code from the browser closure while adding 26.9 kB of comments explaining the
 * consolidations, and the gate booked the net as a 10.2 kB regression that all but failed
 * the entry. Read closure, same window: +25.3 kB of comments against +7.8 kB of code. A
 * gate whose sign can be opposite to the truth is worse than no gate, because a re-freeze
 * looks like the answer every time.
 *
 * Minifying makes the number mean the thing its budget claims. It is still a proxy — a
 * consumer's bundler tree-shakes across the closure, which this deliberately does not, so
 * the figure stays an upper bound — but it is now within sight of a download rather than
 * three times it, and it moves only when code moves. That also buys sensitivity for free:
 * a 1 kB regression is 1.5% of the minified reader entry against 0.5% of the raw one, so
 * it lands well inside {@link HEADROOM_PCT} instead of hiding under it.
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
import esbuild from 'esbuild'
import { ROOT, isMain, parseCli, runCli } from './script-utils.mjs'

const DIST = path.join(ROOT, 'dist')
const BUDGET = path.join(ROOT, 'scripts', 'bundle-size-budget.json')

/**
 * Entries under budget, by their `dist/` file name — every entry `package.json`'s `exports`
 * publishes, so nothing a consumer can import is unwatched.
 *
 * The browser entry is the one a consumer pays for over the wire and was for a while the only
 * one with a number, on the reasoning that the Node entries are read off a local disk and the
 * rest are opt-in subpaths. What that left out is that *reaching* an entry is what costs: a
 * dependency landing in a chunk `./read` pulls in is paid for by every consumer of `./read`,
 * and until it also reached `browser.js` no gate here could see it. Per-entry is the shape that
 * catches it, and it is available for free — an entry file name carries no content hash, unlike
 * the chunks (`text-Bc8hqWTD.js`), so a budget keyed on these does not churn on every build.
 *
 * Adding an entry is a line here plus a `--freeze`. The closures overlap heavily (the shared
 * chunks are counted once per entry that reaches them), which is fine: this is a per-entry
 * question — "what does importing *this* cost" — not a package total. Together they do reach
 * every `.js` file the build emits, which is the property that makes "unwatched bytes" a thing
 * that cannot happen rather than a thing nobody has checked lately.
 */
const ENTRIES = [
	'browser.js',
	'html.js',
	'index.js',
	'inspect.js',
	'math.js',
	'measure.js',
	'node.js',
	'read.js',
	'script.js',
	'zip.js',
]

/** Room `--freeze` leaves above the measurement, so ordinary work is not a re-freeze. */
const HEADROOM_PCT = 5

/** Re-freeze is only worth asking for when an entry comes in this far under budget. */
const SLACK_PCT = 15

/**
 * ...and this far under in absolute terms, which is what keeps the small entries usable.
 *
 * `--freeze` rounds the budget up to a whole kB, and on a 5 kB entry that rounding alone is
 * larger than {@link SLACK_PCT} of it: `zip.js` froze at 6 kB, measured 5 kB, and was
 * immediately 16% under — a nag no re-freeze could clear, because the next freeze rounds to the
 * same 6 kB. A percentage of a tiny number is noise; asking for a re-freeze over 1 kB is asking
 * for a gate to be switched off.
 */
const SLACK_MIN_BYTES = 2048

/**
 * Relative specifiers, static and dynamic. Bare ones are the consumer's to resolve.
 *
 * Three forms, and the third is the one that matters: `from './x'` (import and
 * `export … from` alike), `import('./x')`, and the side-effect `import './x'`.
 * Missing that last form is the failure mode a ratchet cannot afford — a
 * side-effect import drops its file *and everything below it* from the closure,
 * so the measurement falls and the check passes. `\(?` is what admits both the
 * call form and the bare one.
 */
export const RELATIVE_IMPORT = /(?:\bfrom\s*|\bimport\s*\(?\s*)["'](\.[^"']*)["']/g

/**
 * One block comment, built from a string rather than written as a regex literal.
 *
 * A regex that matches block comments has to spell the closing delimiter, and a star-slash inside
 * a regex literal is where naive scanners lose the thread: es-module-lexer, which is what Vite
 * runs over this file when the test suite imports it, took it for the end of a comment and
 * rejected everything after it as a syntax error — while `node --check` and esbuild both accepted
 * the very same file. Constructing the pattern here costs one layer of backslashes and takes the
 * question off the table.
 */
const BLOCK_COMMENT = new RegExp('/\\*[\\s\\S]*?\\*/', 'g')

/**
 * Blank out block comments and whole-line `//` comments, so a specifier that is only *written
 * about* is not counted as one that is imported.
 *
 * `dist/shapes-*.js` carries `{@link import('./notes.js')}` in a doc comment, and the entry
 * closure walked straight into it and demanded a `dist/notes.js` that no build emits. Doc text
 * is the only place this happens, which is why block comments are the case worth handling; the
 * emitted files are machine-written, so a `//` comment is a whole line.
 *
 * Deliberately narrow. Stripping *trailing* `//` would also eat the `//` inside any URL string
 * literal on that line, and a parser that quietly removes code is the failure mode this whole
 * file guards against — over-counting fails loudly and gets looked at, under-counting passes.
 * Newlines are preserved so the remaining offsets and line structure are unchanged.
 * @param {string} text file contents
 * @returns {string} the same text with comment bodies replaced by spaces
 */
export function stripComments(text) {
	return text
		.replace(BLOCK_COMMENT, (match) => match.replace(/[^\n]/g, ' '))
		.replace(/^[ \t]*\/\/[^\n]*/gm, (match) => ' '.repeat(match.length))
}
/**
 * Relative specifiers named by one emitted file, in source order.
 *
 * Split out from {@link closureOf} because it is the whole of the parsing risk and
 * needs no filesystem to exercise — see `test/scripts/bundle-size-ratchet.test.js`.
 * @param {string} text file contents
 * @returns {string[]} relative specifiers, duplicates included
 */
export function relativeImportsOf(text) {
	return [...stripComments(text).matchAll(RELATIVE_IMPORT)]
		.map(([, specifier]) => specifier)
		.filter((s) => s !== undefined)
}

/**
 * Every emitted file an entry pulls in, including itself.
 * @param {string} entry file name under `dir`
 * @param {string} [dir] directory holding the emitted files; defaults to `dist/`
 * @returns {string[]} file names, sorted
 */
export function closureOf(entry, dir = DIST) {
	const seen = new Set()
	const queue = [entry]
	while (queue.length) {
		const name = queue.shift()
		if (!name || seen.has(name)) continue
		const file = path.join(dir, name)
		if (!fs.existsSync(file)) throw new Error(`dist/${name} is missing — run \`pnpm run build\` first`)
		seen.add(name)
		for (const specifier of relativeImportsOf(fs.readFileSync(file, 'utf8'))) {
			queue.push(path.posix.normalize(path.posix.join(path.posix.dirname(name), specifier)))
		}
	}
	return [...seen].sort()
}

/**
 * One emitted file as a consumer's bundler would serve it: minified, then gzipped.
 *
 * `transformSync` rather than `build`, because the unit is the file, not the closure —
 * bundling here would tree-shake across the closure and turn the measurement into a claim
 * about one consumer's import graph rather than about what the package ships. Per-file
 * also keeps {@link measureEntries} synchronous and the shared chunks counted identically
 * for every entry that reaches them.
 *
 * `legalComments: 'none'` because the default keeps `@license` banners, which would put
 * the same prose sensitivity back in through the one comment class minifying preserves.
 * Nothing in `dist/` carries one today; this is so nothing has to notice when something
 * does.
 * @param {string} name file name under `dir`
 * @param {string} [dir] directory holding the emitted files; defaults to `dist/`
 * @returns {number} minified, gzipped bytes
 */
export function shippedBytes(name, dir = DIST) {
	const source = fs.readFileSync(path.join(dir, name), 'utf8')
	const { code } = esbuild.transformSync(source, {
		format: 'esm',
		legalComments: 'none',
		loader: 'js',
		minify: true,
		target: 'es2024',
	})
	return zlib.gzipSync(Buffer.from(code), { level: 9 }).byteLength
}

/** @param {number} bytes */
const kb = (bytes) => (bytes / 1024).toFixed(1) + ' kB'

/**
 * Measure every budgeted entry against `dist/`.
 * @returns {Map<string, {files: Array<{name: string, bytes: number}>, bytes: number}>}
 */
export function measureEntries() {
	const measured = new Map()
	for (const entry of ENTRIES) {
		const files = closureOf(entry).map((name) => ({ name, bytes: shippedBytes(name) }))
		measured.set(entry, { files, bytes: files.reduce((sum, file) => sum + file.bytes, 0) })
	}
	return measured
}

// ---------------------------------------------------------------- CLI

const USAGE = `Bundle-size budget — how big each published entry point has got.

  node scripts/bundle-size-ratchet.mjs            check (exit 1 over budget)
  node scripts/bundle-size-ratchet.mjs --freeze   rewrite the budget from dist/
  node scripts/bundle-size-ratchet.mjs --list     per-chunk breakdown

Options:
  --freeze    write ${HEADROOM_PCT}% above today's measurement into the budget file
  --list      print every file in each entry's closure, largest first
  -h, --help  show this message`

/** @param {string[]} argv @returns {number} process exit code */
export function main(argv) {
	const { values } = parseCli(argv, {
		usage: USAGE,
		options: {
			freeze: { type: 'boolean', default: false },
			list: { type: 'boolean', default: false },
		},
	})

	const measured = measureEntries()

	if (values.list) {
		for (const [entry, { files, bytes }] of measured) {
			console.log(`${entry}: ${kb(bytes)} minified+gzipped, ${files.length} file(s)`)
			for (const file of [...files].sort((a, b) => b.bytes - a.bytes))
				console.log(`  ${kb(file.bytes).padStart(9)}  ${file.name}`)
		}
		return 0
	}

	if (values.freeze) {
		/** @type {Record<string, number>} */
		const frozen = {}
		// Rounded up to a whole kB so the budget reads as a decision someone made rather than
		// as a build artifact copied into a file.
		for (const [entry, { bytes, files }] of measured) {
			frozen[entry] = Math.ceil((bytes * (1 + HEADROOM_PCT / 100)) / 1024) * 1024
			console.log(
				`bundle size: froze ${entry} at ${kb(frozen[entry])} (measured ${kb(bytes)}, ${files.length} file(s))`
			)
		}
		fs.writeFileSync(BUDGET, JSON.stringify(frozen, null, '\t') + '\n')
		return 0
	}

	/** @type {Record<string, number>} */
	const budgetFile = JSON.parse(fs.readFileSync(BUDGET, 'utf8'))
	const relBudget = path.relative(ROOT, BUDGET).split(path.sep).join('/')

	const missing = [...measured.keys()].filter((entry) => typeof budgetFile[entry] !== 'number')
	if (missing.length) {
		console.error(`bundle size FAILED — no budget for ${missing.join(', ')} in ${relBudget}.`)
		console.error('\n  pnpm run bundle-size:freeze')
		return 1
	}

	const checked = [...measured].map(([entry, { bytes, files }]) => ({
		entry,
		bytes,
		files,
		budget: budgetFile[entry] ?? 0,
	}))
	const over = checked.filter((row) => row.bytes > row.budget)
	const under = checked.filter(
		(row) => row.bytes < row.budget * (1 - SLACK_PCT / 100) && row.budget - row.bytes >= SLACK_MIN_BYTES
	)

	if (over.length) {
		console.error('bundle size FAILED — an entry grew past its budget:\n')
		for (const { entry, bytes, files, budget } of over) {
			console.error(`  ${entry}: ${kb(bytes)} minified+gzipped (budget ${kb(budget)}), ${files.length} file(s)`)
			for (const file of [...files].sort((a, b) => b.bytes - a.bytes).slice(0, 5))
				console.error(`    ${kb(file.bytes).padStart(9)}  ${file.name}`)
		}
		console.error('\nRun `pnpm run bundle-size:list` for the full breakdown. If the growth is')
		console.error(`intended, raise the number in ${relBudget} in the same commit and say what`)
		console.error('bought the bytes; if it is not, something reached the browser entry that should not.')
		return 1
	}

	if (under.length) {
		console.log(`bundle size: ${SLACK_PCT}%+ under budget — bank it by lowering ${relBudget}:\n`)
		for (const { entry, bytes, budget } of under) console.log(`  ${entry}: ${kb(budget)} -> ${kb(bytes)}`)
		console.log('\n  pnpm run bundle-size:freeze')
		return 1
	}

	for (const { entry, bytes, files, budget } of checked)
		console.log(
			`bundle size: ok — ${entry} ${kb(bytes)} minified+gzipped (budget ${kb(budget)}), ${files.length} file(s)`
		)
	return 0
}

if (isMain(import.meta.url)) await runCli(() => main(process.argv.slice(2)))
