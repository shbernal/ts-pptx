#!/usr/bin/env node
/**
 * Byte-identity gate for write-side refactors.
 *
 * AGENTS.md: "OOXML is fixture-gated; no changing emitted bytes as cleanup."
 * This proves a behavior-preserving refactor of the `src/gen/` emitters does not
 * change a single emitted byte, by generating every showcase deck, exploding each
 * one (recursing into its embedded .xlsx parts, which are their own OPC packages),
 * and diffing every part against a frozen baseline.
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
import { diffParts, explodePackage, listParts, loadShowcases } from './pptx-parts.mjs'
import { ROOT, parseCliOrExit, run } from './script-utils.mjs'

const OUT_ROOT = path.join(ROOT, '.tmp', 'byte-identity')
const BASELINE = path.join(OUT_ROOT, 'baseline')
const CURRENT = path.join(OUT_ROOT, 'current')
// Written here rather than into `demos/showcases/output/`: the gate builds decks on every
// run, and it has no business clobbering the artifacts `pnpm demos:build` leaves for a human.
const DECKS = path.join(OUT_ROOT, 'decks')

const USAGE = `Byte-identity gate for write-side refactors.

  node scripts/byte-identity.mjs baseline   freeze current output as the reference
  node scripts/byte-identity.mjs check      rebuild, regenerate, diff vs baseline

Freeze a baseline BEFORE the refactor, then \`check\` after each step.

Options:
  --allow-dirty   permit \`baseline\` on an uncommitted src/gen/ (read the refusal first)
  -h, --help      show this message`

const { values, positionals } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	allowPositionals: true,
	options: { 'allow-dirty': { type: 'boolean', default: false } },
})
const mode = positionals[0]
if (mode !== 'baseline' && mode !== 'check') {
	console.error(mode ? `unknown subcommand: ${mode}` : 'a subcommand is required')
	console.error('\n' + USAGE)
	process.exit(2)
}

/**
 * Build every corpus deck with each nondeterministic source pinned.
 *
 * The corpus has two halves, and they are different kinds of thing:
 *
 * - **Showcase decks** (`demos/showcases/`) are presentation decks that happen to drive the
 *   emitters end to end. They resolve their assets from their own URL, so — unlike the demo
 *   runner this replaced — no `process.chdir` is needed, and they import `@shbernal/ts-pptx`
 *   through the workspace link, which resolves to the `dist/` the build above just wrote.
 * - **Gate decks** (`scripts/gate-decks/`) are fixture matrices shaped like a `.pptx`. They
 *   exist because the showcases only reach what a plausible deck would reach: three chart
 *   types out of nine chart emitters, which left most of `src/gen/chart/` with no evidence
 *   at all. AGENTS.md is explicit that a PASS on an emitter no deck reaches is "unproven,
 *   not proven unchanged" — so the parts that no showcase would ever want get their own
 *   corpus rather than being bolted onto a deck that has a different job.
 *
 * Both are loaded by dynamic import rather than a static one: they pull in `dist/`, which
 * the build above writes moments earlier and which may not exist when this module is first
 * evaluated.
 *
 * Returns one `{ slug, file }` per deck.
 */
async function generateDecks() {
	const { GATE_DECKS } = await import('./gate-decks/index.mjs')
	const corpus = [...(await loadShowcases()), ...GATE_DECKS]

	fs.rmSync(DECKS, { recursive: true, force: true })
	fs.mkdirSync(DECKS, { recursive: true })

	const decks = []
	for (const showcase of corpus) {
		// `getUuid` (gen-utils) and the chart-colour fallback both draw on Math.random, so
		// section ids and `c16:uniqueId` vary per run. Reseed per deck rather than once for
		// the process: with a single stream, editing deck 1 shifts every GUID in deck 2 and
		// the gate reports a diff in a deck nobody touched.
		let seed = 0x2545f491
		Math.random = () => {
			seed = (seed * 1103515245 + 12345) & 0x7fffffff
			return seed / 0x80000000
		}

		const file = path.join(DECKS, showcase.fileName)
		await showcase.build(file)
		if (!fs.existsSync(file)) throw new Error('showcase deck was not written: ' + file)
		decks.push({ slug: showcase.slug, file })
	}
	return decks
}

/**
 * Explode every deck into `destDir/<slug>/`, recursing into embedded .xlsx parts.
 *
 * The per-slug prefix is what keeps a diff readable — two decks share part names
 * (`ppt/slides/slide1.xml` and friends), and flattening them into one tree would both
 * collide and hide which deck moved.
 * @param {readonly {file: string, slug: string}[]} decks
 * @param {string} destDir
 */
async function explodeDecks(decks, destDir) {
	fs.rmSync(destDir, { recursive: true, force: true })
	for (const deck of decks) {
		await explodePackage(new Uint8Array(fs.readFileSync(deck.file)), path.join(destDir, deck.slug))
	}
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

if (mode === 'baseline' && !values['allow-dirty']) assertGenTreeClean()

// Run the bundler's JS entry directly rather than `pnpm run build`: on Windows
// the pnpm shim is a .cmd, and Node >=20 refuses to spawn one without a shell.
await run(process.execPath, [path.join(ROOT, 'node_modules', 'tsdown', 'dist', 'run.mjs')])

if (mode === 'baseline') {
	await explodeDecks(await generateDecks(), BASELINE)
	console.log('baseline frozen: ' + listParts(BASELINE).length + ' parts -> ' + path.relative(ROOT, BASELINE))
	process.exit(0)
}

if (!fs.existsSync(BASELINE)) {
	console.error('no baseline to compare against; freeze one first:')
	console.error('  node scripts/byte-identity.mjs baseline')
	process.exit(2)
}

await explodeDecks(await generateDecks(), CURRENT)

const diffs = diffParts(BASELINE, CURRENT)
if (diffs.length === 0) {
	console.log('PASS - byte-identical (' + listParts(CURRENT).length + ' parts)')
	process.exit(0)
}

console.error('FAIL - emitted bytes changed (' + diffs.length + ' part(s)):')
for (const line of diffs.slice(0, 40)) console.error('  ' + line)
if (diffs.length > 40) console.error('  ... and ' + (diffs.length - 40) + ' more')
process.exit(1)
