#!/usr/bin/env node
/**
 * Round-trip harness — does a generated script rebuild the deck it was generated from?
 *
 * Why this exists. A printed script that typechecks proves nothing, and one that *runs*
 * proves only that no argument was malformed. Neither says whether the deck it produces is
 * the deck it was read from. This closes that: source → IR₁ → script → run it → output →
 * IR₂, then diff IR₁ against IR₂ with the printer's own fidelity notes as the exclusion
 * list. A note is a promise that a construct will not survive, so a difference no note
 * predicted is a defect and nothing else.
 *
 * The comparison is a *projection* diff, not byte identity. The output package cannot be
 * byte-identical — regenerated shape ids, fresh rel ids, a rebuilt shape tree — so comparing
 * packages would report a total mismatch for every deck and measure nothing.
 *
 * What it cannot see, stated so nobody reads a clean run as more than it is: both IRs come
 * from the same reader, so a construct the read model does not read is missing from both
 * and compares equal. `pnpm run read:census` measures that surface. This script certifies
 * "nothing the converter can see was lost", never "nothing was lost".
 *
 * Two tiers, one harness. `--tier b` (the default) prints the template-anchored script and
 * runs it beside the source deck; `--tier a` prints the standalone one and deliberately does
 * *not* lay a template down, so a script that still needed one fails here rather than passing
 * on the file it is meant to replace. The comparison is the same either way — only what the
 * fidelity notes excuse differs, which is the point of running both.
 *
 * Usage: `pnpm run script:roundtrip -- --help` (the flag list lives in USAGE below, so
 * there is one copy of it to keep true).
 */
import { execFile } from 'node:child_process'
import fs from 'node:fs/promises'
import path from 'node:path'
import { promisify } from 'node:util'
import { ROOT, corpusDecks, parseCliOrExit, resolveCorpusDir } from './script-utils.mjs'
import { Presentation } from '../dist/read.js'
import { canonicalDeckIr, diffDeckIr, printScript, printStandaloneScript, readModelToIr } from '../dist/script.js'

const run = promisify(execFile)

const USAGE = `Round-trip harness — does a generated script rebuild the deck it came from?

  pnpm run script:roundtrip
  pnpm run script:roundtrip -- --tier a
  pnpm run script:roundtrip -- --fixture mixed.pptx --verbose
  pnpm run script:roundtrip -- --dir ~/decks        # absolute paths welcome

Options:
  --tier <a|b>       a = standalone script, b = template-anchored (default b)
  --fixture <name>   restrict the run to one .pptx by file name
  --dir <path>       corpus directory (default test/read/fixtures)
  --json             machine-readable report on stdout
  --verbose          list each undeclared difference and the write-path defaults
  -h, --help         show this message`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: {
		tier: { type: 'string', default: 'b' },
		fixture: { type: 'string' },
		dir: { type: 'string' },
		json: { type: 'boolean', default: false },
		verbose: { type: 'boolean', default: false },
	},
})
const asJson = values.json
const verbose = values.verbose
const only = values.fixture ?? null
const tier = values.tier.toLowerCase()
if (tier !== 'a' && tier !== 'b') {
	console.error(`--tier must be a (standalone) or b (template-anchored), got ${JSON.stringify(values.tier)}`)
	console.error('\n' + USAGE)
	process.exit(2)
}
const DIR = resolveCorpusDir(values.dir)

// Inside the repo rather than the OS temp directory, and required to be: the emitted script
// imports this package by its published name, which Node resolves by the self-reference rule
// only from a path underneath the package root. `/.tmp/` is gitignored.
const SCRATCH = path.join(ROOT, '.tmp')

const names = await corpusDecks({ dir: DIR, only })

await fs.mkdir(SCRATCH, { recursive: true })

/**
 * Print a script for one deck, run it, and read back what it produced.
 * @param {string} file path to the source deck
 * @returns {Promise<import('../dist/script.js').RoundTripReport>}
 */
async function roundTrip(file) {
	const dir = await fs.mkdtemp(path.join(SCRATCH, 'roundtrip-'))
	try {
		const bytes = await fs.readFile(file)
		const ir = readModelToIr(await Presentation.load(bytes))
		const printed = tier === 'a' ? printStandaloneScript(ir) : printScript(ir)

		await fs.writeFile(path.join(dir, 'script.ts'), printed.code)
		// Tier B's template is the source deck unchanged: `fromTemplate` strips its slides
		// itself. Tier A gets nothing, on purpose — see the header.
		if (tier === 'b') await fs.writeFile(path.join(dir, 'template.pptx'), bytes)
		if (printed.assets.size > 0) {
			await fs.mkdir(path.join(dir, 'assets'), { recursive: true })
			for (const [name, data] of printed.assets) await fs.writeFile(path.join(dir, 'assets', name), data)
		}

		// Node's type stripping runs the emitted TypeScript directly, so the harness needs no
		// transpiler and the script under test is the exact text a user would be handed.
		await run(process.execPath, ['--no-warnings', path.join(dir, 'script.ts')])
		const output = readModelToIr(await Presentation.load(await fs.readFile(path.join(dir, 'output.pptx'))))
		return diffDeckIr(canonicalDeckIr(ir), canonicalDeckIr(output), printed.notes)
	} finally {
		await fs.rm(dir, { recursive: true, force: true })
	}
}

/**
 * @typedef {import('../dist/script.js').RoundTripReport} Report
 * @typedef {{fixture: string, failed: string | null, report: Report | null}} Result
 */

/** @type {Result[]} */
const results = []
for (const name of names) {
	try {
		results.push({ fixture: name, failed: null, report: await roundTrip(path.join(DIR, name)) })
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error)
		results.push({ fixture: name, failed: message.split('\n')[0] ?? message, report: null })
	}
}

/** @param {Result} result */
const ok = (result) => result.report !== null && result.report.undeclared.length === 0

if (asJson) {
	console.log(JSON.stringify({ dir: path.relative(ROOT, DIR), tier, results }, null, 2))
	process.exit(results.every(ok) ? 0 : 1)
}

const pad = Math.max(...names.map((name) => name.length))
for (const result of results) {
	const report = result.report
	if (!report) {
		console.log(`${result.fixture.padEnd(pad)}  FAILED  ${(result.failed ?? '').slice(0, 90)}`)
		continue
	}
	console.log(
		`${result.fixture.padEnd(pad)}  ${report.undeclared.length === 0 ? 'ok  ' : 'LOSS'}` +
			`  undeclared=${String(report.undeclared.length).padStart(4)}` +
			`  declared=${String(report.declared.length).padStart(4)}  added=${String(report.added.length).padStart(4)}`
	)
	if (!verbose) continue
	for (const difference of report.undeclared.slice(0, 40)) {
		console.log(
			`      ${difference.kind} slide ${difference.slideNumber} ${difference.shapeName ?? '—'} ` +
				`${difference.path}: ${difference.expected} → ${difference.actual}`
		)
	}
}

/**
 * Roll the differences up by field, which is the unit a fix or a note is written against.
 * @param {(report: Report) => import('../dist/script.js').IrDifference[]} pick
 * @returns {[string, number][]}
 */
function tally(pick) {
	/** @type {Map<string, number>} */
	const counts = new Map()
	for (const result of results) {
		if (!result.report) continue
		for (const difference of pick(result.report)) {
			const key = `${difference.kind} ${difference.field}`
			counts.set(key, (counts.get(key) ?? 0) + 1)
		}
	}
	return [...counts].sort((a, b) => b[1] - a[1])
}

/** @param {(report: Report) => import('../dist/script.js').IrDifference[]} pick */
const total = (pick) => results.reduce((sum, result) => sum + (result.report ? pick(result.report).length : 0), 0)
const undeclared = total(/** @param {Report} report */ (report) => report.undeclared)
const declared = total(/** @param {Report} report */ (report) => report.declared)
const added = total(/** @param {Report} report */ (report) => report.added)
const broken = results.filter((result) => result.report === null).length

console.log(
	`\ntier ${tier} — ${results.length} deck(s): ${undeclared} undeclared difference(s), ${declared} declared, ` +
		`${added} write-path default(s)${broken ? `, ${broken} failed to run` : ''}.`
)

if (undeclared > 0) {
	console.log('\nUndeclared, by field — each is either a converter defect or a missing fidelity note:')
	for (const [key, count] of tally(/** @param {Report} report */ (report) => report.undeclared).slice(0, 25)) {
		console.log(`  ${String(count).padStart(5)}  ${key}`)
	}
}
if (added > 0 && verbose) {
	console.log('\nWrite-path defaults, by field — benign only where the value matches what the source inherited:')
	for (const [key, count] of tally(/** @param {Report} report */ (report) => report.added).slice(0, 25)) {
		console.log(`  ${String(count).padStart(5)}  ${key}`)
	}
}

process.exit(undeclared === 0 && broken === 0 ? 0 : 1)
