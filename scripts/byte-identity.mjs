#!/usr/bin/env node
/**
 * Byte-identity gate for write-side refactors.
 *
 * AGENTS.md: "OOXML is fixture-gated; no changing emitted bytes as cleanup."
 * This proves a behavior-preserving refactor of the `src/gen/` emitters does not
 * change a single emitted byte, by generating the full demo deck, exploding it
 * (recursing into every embedded .xlsx, which is its own OPC package), and
 * diffing every part against a frozen baseline.
 *
 *   node scripts/byte-identity.mjs baseline   # freeze current output as the reference
 *   node scripts/byte-identity.mjs check      # rebuild, regenerate, diff vs baseline
 *
 * Workflow: freeze a baseline BEFORE the refactor, then `check` after each step.
 * `baseline` enforces that ordering by refusing to run on a dirty `src/gen/`
 * (override with `--allow-dirty`). Both subcommands run `pnpm run build` first —
 * the deck is generated from `dist/`, not `src/`, so a stale build would silently
 * gate the wrong code.
 */

import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { ROOT, run } from './script-utils.mjs'

const OUT_ROOT = path.join(ROOT, '.tmp', 'byte-identity')
const BASELINE = path.join(OUT_ROOT, 'baseline')
const CURRENT = path.join(OUT_ROOT, 'current')
const DECK = path.join(ROOT, 'demos', 'node', 'output', 'TsPptx_Demo_All.pptx')

const mode = process.argv[2]
if (mode !== 'baseline' && mode !== 'check') {
	console.error('usage: node scripts/byte-identity.mjs <baseline|check>')
	process.exit(2)
}

/**
 * Emitted values that legitimately differ between two identical runs.
 * Deliberately narrow: normalizing ONLY these keeps a changed *fixed* GUID
 * (e.g. a built-in table-style id) visible as a real diff.
 */
const NORMALIZERS = [
	// core.xml timestamps — the deck's and every embedded workbook's
	[
		/<dcterms:(created|modified) xsi:type="dcterms:W3CDTF">[^<]*<\/dcterms:\1>/g,
		'<dcterms:$1 xsi:type="dcterms:W3CDTF">NORMALIZED-TIMESTAMP</dcterms:$1>',
	],
	// presentation.xml section ids — random GUID per run
	[/(<p14:section[^>]*\bid=")\{[^}]*\}"/g, '$1{NORMALIZED-SECTION}"'],
	// chartN.xml uniqueId — random GUID per run
	[/(<c16:uniqueId[^>]*\bval=")\{[^}]*\}"/g, '$1{NORMALIZED-UNIQUEID}"'],
]

function normalize(text) {
	return NORMALIZERS.reduce((out, [re, sub]) => out.replace(re, sub), text)
}

/** Generate the demo deck with every nondeterministic source pinned. */
async function generateDeck() {
	// The chart demo builds its series from Math.random; pin it so the emitted
	// chart XML (and the embedded workbooks) are stable run-to-run.
	let seed = 0x2545f491
	Math.random = () => {
		seed = (seed * 1103515245 + 12345) & 0x7fffffff
		return seed / 0x80000000
	}

	fs.rmSync(DECK, { force: true })
	// Demo image paths are relative (`../common/images/*`), so cwd must be demos/node.
	const cwd = process.cwd()
	process.chdir(path.join(ROOT, 'demos', 'node'))
	try {
		const TsPptx = (await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)).default
		const { runEveryTest } = await import(pathToFileURL(path.join(ROOT, 'demos', 'modules', 'demos.mjs')).href)
		await runEveryTest(TsPptx)
	} finally {
		process.chdir(cwd)
	}
	if (!fs.existsSync(DECK)) throw new Error('demo deck was not written: ' + DECK)
}

/** Explode an OPC package into `destDir`, recursing into embedded .xlsx parts. */
async function explode(destDir) {
	const { unzipSync } = await import(pathToFileURL(path.join(ROOT, 'node_modules', 'fflate', 'esm', 'browser.js')).href)
	const decoder = new TextDecoder('utf-8')

	const dump = (zipBytes, dir) => {
		const entries = unzipSync(zipBytes)
		for (const name of Object.keys(entries).sort()) {
			const bytes = entries[name]
			// Each embedded workbook is its own OPC zip — recurse rather than
			// diffing opaque compressed bytes.
			if (/\.xlsx$/i.test(name)) {
				dump(bytes, path.join(dir, name + '!'))
				continue
			}
			const dest = path.join(dir, name)
			fs.mkdirSync(path.dirname(dest), { recursive: true })
			if (/\.(xml|rels)$/i.test(name)) fs.writeFileSync(dest, normalize(decoder.decode(bytes)), 'utf8')
			else fs.writeFileSync(dest, bytes)
		}
	}

	fs.rmSync(destDir, { recursive: true, force: true })
	fs.mkdirSync(destDir, { recursive: true })
	dump(new Uint8Array(fs.readFileSync(DECK)), destDir)
}

function listParts(dir) {
	const out = []
	const walk = (d, prefix) => {
		for (const entry of fs.readdirSync(d, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
			const rel = prefix ? prefix + '/' + entry.name : entry.name
			if (entry.isDirectory()) walk(path.join(d, entry.name), rel)
			else out.push(rel)
		}
	}
	walk(dir, '')
	return out
}

/** Compare two exploded packages. Returns a list of human-readable differences. */
function diffParts(baseDir, curDir) {
	const base = new Set(listParts(baseDir))
	const cur = new Set(listParts(curDir))
	const diffs = []
	for (const part of base) if (!cur.has(part)) diffs.push('REMOVED  ' + part)
	for (const part of cur) if (!base.has(part)) diffs.push('ADDED    ' + part)
	for (const part of base) {
		if (!cur.has(part)) continue
		const a = fs.readFileSync(path.join(baseDir, part))
		const b = fs.readFileSync(path.join(curDir, part))
		if (!a.equals(b)) diffs.push('CHANGED  ' + part)
	}
	return diffs.sort()
}

/**
 * A baseline is the *pre-refactor* reference. Frozen from an already-edited
 * `src/gen/`, it bakes in the very change it exists to detect, and every later
 * `check` passes trivially — a green gate that proves nothing. Refuse instead,
 * and say so, because the obvious workaround is `git stash`: stashing a dirty
 * tree to fake a retroactive baseline risks losing unrelated work to a pop
 * conflict (AGENTS.md: "Preserve unrelated dirty state").
 *
 * Scoped to `src/gen/` deliberately — dirt anywhere else cannot affect the
 * emitted bytes, so it is none of this gate's business.
 */
function assertGenTreeClean() {
	let dirty
	try {
		dirty = execFileSync('git', ['status', '--porcelain', '--', 'src/gen'], { cwd: ROOT, encoding: 'utf8' }).trim()
	} catch {
		return // not a git checkout, or git unavailable — nothing to assert
	}
	if (!dirty) return
	console.error('refusing to freeze a baseline: src/gen/ has uncommitted changes:')
	for (const line of dirty.split('\n')) console.error('  ' + line)
	console.error('')
	console.error('The baseline must be taken BEFORE the refactor. Taken now, it records the')
	console.error('post-edit bytes as the reference and `check` can no longer fail.')
	console.error('Commit or revert src/gen/ first, or freeze from the pre-refactor commit.')
	console.error('Do NOT `git stash` a dirty tree to work around this.')
	console.error('If the current state genuinely IS the intended reference: --allow-dirty')
	process.exit(2)
}

// ---------------------------------------------------------------- main

if (mode === 'baseline' && !process.argv.includes('--allow-dirty')) assertGenTreeClean()

// Run the bundler's JS entry directly rather than `pnpm run build`: on Windows
// the pnpm shim is a .cmd, and Node >=20 refuses to spawn one without a shell.
await run(process.execPath, [path.join(ROOT, 'node_modules', 'tsdown', 'dist', 'run.mjs')])

if (mode === 'baseline') {
	await generateDeck()
	await explode(BASELINE)
	console.log('baseline frozen: ' + listParts(BASELINE).length + ' parts -> ' + path.relative(ROOT, BASELINE))
	process.exit(0)
}

if (!fs.existsSync(BASELINE)) {
	console.error('no baseline to compare against; freeze one first:')
	console.error('  node scripts/byte-identity.mjs baseline')
	process.exit(2)
}

await generateDeck()
await explode(CURRENT)

const diffs = diffParts(BASELINE, CURRENT)
if (diffs.length === 0) {
	console.log('PASS - byte-identical (' + listParts(CURRENT).length + ' parts)')
	process.exit(0)
}

console.error('FAIL - emitted bytes changed (' + diffs.length + ' part(s)):')
for (const line of diffs.slice(0, 40)) console.error('  ' + line)
if (diffs.length > 40) console.error('  ... and ' + (diffs.length - 40) + ' more')
process.exit(1)
