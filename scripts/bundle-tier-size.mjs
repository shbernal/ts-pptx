#!/usr/bin/env node
/**
 * Tier-size budget — what a consumer program actually downloads.
 *
 *   node scripts/bundle-tier-size.mjs            # check (exit 1 over budget)
 *   node scripts/bundle-tier-size.mjs --freeze   # rewrite the budget from dist/
 *   node scripts/bundle-tier-size.mjs --list     # per-chunk breakdown
 *
 * **Why a second size gate.** `bundle-size-ratchet.mjs` measures, per published entry, the
 * transitive closure of that entry's relative imports, minified and gzipped but never
 * bundled — deliberately, so its figure is an upper bound on what the package *ships*. That
 * is the right answer to "what does importing this entry cost at worst", and it is
 * structurally unable to see the question this one asks. Making a construct family
 * unreachable for a program that never calls into it does not delete a byte from the
 * closure, so the ratchet's number does not move; only a measurement that bundles a real
 * program and lets a bundler shake it can tell whether the reachability actually changed.
 * The two numbers will not agree and are not meant to. Neither replaces the other.
 *
 * **Method.** `scripts/comparison/hygiene.mjs` already argues these conventions and this
 * does not re-derive them: bundle one plausible consumer program with esbuild, minify,
 * gzip at level 9, and record two figures rather than one. `initial` is what the program pays
 * to start: the entry chunk plus every chunk reachable from it by `import` statements, which is
 * what a browser fetches before the first line runs. `total` is every chunk it can reach,
 * because font metrics load `opentype.js` through a dynamic import that only runs on first
 * font registration: charging a program for a chunk it may never fetch is as wrong as
 * hiding bytes it might. There is no fair single number, so there is no single number.
 *
 * **The one convention inverted from hygiene.** Hygiene packs a tarball and installs it
 * into a clean prefix, because it compares two libraries and a development tree is not
 * what either consumer gets. This runs on every `verify:full`, where a pack plus an
 * install per run is not a gate but a chore, so the program imports `dist/browser.js`
 * directly and bare specifiers resolve against the repo's own `node_modules`. That costs
 * nothing here: the comparison being made is against this repo's own frozen numbers, not
 * against another library's install, so a hoisted tree cannot tilt it.
 *
 * **Resolved against `dist/`, not `src/`.** This is the number a consumer pays, and
 * measuring `dist/` also puts chunk granularity under the gate: a consumer's bundler can
 * only shake what the chunk boundary lets it, so a chunking change is a size lever that
 * moves nothing in `src/` and would be invisible to a `src/`-based measurement. It is not
 * costing anything: measured when this gate was written, the same text program came to
 * 139.2 kB gzip from `dist/` against 139.8 kB from `src/`, dependencies external on both
 * sides, so tsdown's chunking is if anything a shade cheaper than the flat graph. Only the
 * gap is the claim — the absolutes have moved since, and the rows the gate prints are the
 * current ones. It is here for the day the gap stops falling `dist/`'s way.
 *
 * **Two program shapes, not one.** The three `new TsPptx()` rows measure the default class,
 * which is composed with every family and always will be. The `composed-*` rows measure
 * `createPresentation`, which is composed with the core tier plus what it names — the surface
 * the whole split exists to make possible. Both shapes are real programs a consumer writes,
 * and the difference between the two `composed-*` rows is what one family costs, on `dist/`
 * bytes, with no stub anywhere.
 *
 * **What it is watching for.** Now that a family can be left out, the regression only this
 * gate can see is a static import from `slide.ts`, `gen/slide/object.ts` or
 * `package/assemble.ts` into a family module: that puts the family back in the graph of
 * every program, the ones composed without it included. Nothing else reports it. The types
 * check, the tests pass, the emitted bytes are identical, and `bundle-size-ratchet.mjs` does
 * not move, because `dist/` ships exactly what it shipped before. Here, the `composed-*` rows
 * climb toward the `full` one and the budget fails. `docs/bundle-size.md` states that rule
 * for a reader and carries the numbers a consumer cares about.
 *
 * **No per-family stub mode.** "What would removing this family save" is a different
 * measurement: it bundles `src/` with whole modules stubbed out, and it is only interesting
 * while the family is still unconditionally reachable. Once a family can be left out, the
 * tiers below report its cost as the difference between two rows, on `dist/` bytes, without
 * a stub anywhere. Folding a `src/`-side stub loop into a script whose claim is "what
 * `dist/` costs a consumer" would put two measurements behind one name.
 */

import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import zlib from 'node:zlib'
import esbuild from 'esbuild'
import { HEADROOM_PCT, SLACK_MIN_BYTES, SLACK_PCT } from './bundle-size-ratchet.mjs'
import { ROOT, isMain, parseCli, runCli } from './script-utils.mjs'

const DIST_ENTRY = path.join(ROOT, 'dist', 'browser.js')
const DIST_FAMILIES = path.join(ROOT, 'dist', 'families.js')
const BUDGET = path.join(ROOT, 'scripts', 'bundle-tier-budget.json')
const WORK = path.join(ROOT, '.tmp', 'bundle-tier')

/** The entry file name esbuild is given, and therefore the output chunk that is the entry. */
const PROGRAM_FILE = 'program.mjs'

/** Where each program parks its deck, so the export is neither printed nor dead. */
const SINK = '__bundleTierSink'

/** A 1x1 transparent PNG, so `addImage` gets a real base64 payload without a fixture read. */
const PNG_1X1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/**
 * The three consumer programs, each a superset of the one above it.
 *
 * Plausible programs, not synthetic minima: a tier that only measures the smallest call a
 * type checker will accept measures the type checker. `text` is the tier the family split
 * exists to make cheap. `full` is the control — it must *not* get much cheaper, because a
 * `full` that falls with `text` means something was dropped rather than deferred.
 *
 * Every program ends by assigning the export to a `globalThis` property. An export whose
 * value is discarded is dead code, and a minifier that proves it drops the call graph
 * behind it, which would turn each row into a measurement of side-effect annotations
 * instead of size. Hygiene uses `console.log` for this; a property assignment is equally
 * undroppable and stays quiet when {@link measureTier} runs the program for real.
 * @type {Record<string, string[]>}
 */
const TIERS = {
	text: ["slide.addText('hello', { x: 1, y: 1, w: 4, h: 1 })"],
	'text-shape-image': [
		"slide.addShape('rect', { x: 1, y: 2.5, w: 2, h: 1, fill: { color: '0088CC' } })",
		`slide.addImage({ data: ${JSON.stringify(PNG_1X1)}, x: 4, y: 2.5, w: 1, h: 1 })`,
	],
	full: [
		"slide.addChart([{ name: 'Series 1', labels: ['a', 'b', 'c'], values: [1, 2, 3] }], { type: 'bar', x: 1, y: 4, w: 4, h: 2 })",
		"slide.addTable([[{ text: 'a' }, { text: 'b' }]], { x: 5.5, y: 4, w: 4 })",
		"slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/aaaaaaaaaaa', x: 1, y: 6, w: 4, h: 2 })",
	],
}

/**
 * The composed programs: `createPresentation` with the core tier and the families each names.
 *
 * Not cumulative, and deliberately a pair. `composed` is the floor the split exists to reach;
 * `composed-charts` is the same program with one family added, so the difference between the two
 * rows is what that family costs a consumer who wants it. A single row could only say "the floor
 * moved" without saying what moved it.
 * @type {Record<string, {use: string[], calls: string[]}>}
 */
const COMPOSED = {
	composed: {
		use: [],
		calls: ["slide.addText('hello', { x: 1, y: 1, w: 4, h: 1 })"],
	},
	'composed-charts': {
		use: ['charts'],
		calls: [
			"slide.addText('hello', { x: 1, y: 1, w: 4, h: 1 })",
			"slide.addChart([{ name: 'Series 1', labels: ['a', 'b', 'c'], values: [1, 2, 3] }], { type: 'bar', x: 1, y: 4, w: 4, h: 2 })",
		],
	},
}

/** Every tier this gate measures, in report order: the composed programs, then the class ones. */
const TIER_NAMES = [...Object.keys(COMPOSED), ...Object.keys(TIERS)]

/** The two figures budgeted per tier, in report order. */
const FIGURES = /** @type {const} */ (['initial', 'total'])

/**
 * Where one tier's program is written, and bundled from.
 * @param {string} tier - a name in {@link TIER_NAMES}
 * @returns {string}
 */
const tierDir = (tier) => path.join(WORK, tier)

/**
 * How a tier's program names a built entry: relative, POSIX, explicitly so.
 *
 * A Windows absolute path would satisfy esbuild, which resolves a specifier as a path, and
 * would leave the program unrunnable by Node, which refuses `import('C:/...')` because `C:`
 * reads as a URL scheme. The program has to satisfy both, because {@link measureTier} runs
 * it before it measures it.
 * @param {string} tier - a tier name
 * @param {string} file - the built file to name (defaults to the browser entry)
 * @returns {string}
 */
function entrySpecifier(tier, file = DIST_ENTRY) {
	const relative = path.relative(tierDir(tier), file).split(path.sep).join('/')
	return relative.startsWith('.') ? relative : './' + relative
}

/**
 * One tier's program source.
 *
 * A {@link COMPOSED} name prints a `createPresentation` program that names its own families; a
 * {@link TIERS} name prints a `TsPptx` program accumulating every tier declared before it.
 * @param {string} tier - a name in {@link TIER_NAMES}
 * @returns {string}
 */
export function programFor(tier) {
	const composed = COMPOSED[tier]
	if (composed)
		return [
			`import { createPresentation } from ${JSON.stringify(entrySpecifier(tier))}`,
			...(composed.use.length > 0
				? [`import { ${composed.use.join(', ')} } from ${JSON.stringify(entrySpecifier(tier, DIST_FAMILIES))}`]
				: []),
			`const pres = createPresentation({ use: [${composed.use.join(', ')}] })`,
			'const slide = pres.addSlide()',
			...composed.calls,
			`globalThis.${SINK} = await pres.write({ outputType: 'arraybuffer' })`,
			'',
		].join('\n')
	const keys = Object.keys(TIERS)
	const upto = keys.indexOf(tier)
	if (upto < 0) throw new Error(`no such tier: ${tier}`)
	return [
		`import TsPptx from ${JSON.stringify(entrySpecifier(tier))}`,
		'const pres = new TsPptx()',
		'const slide = pres.addSlide()',
		...keys.slice(0, upto + 1).flatMap((key) => TIERS[key] ?? []),
		`globalThis.${SINK} = await pres.write({ outputType: 'arraybuffer' })`,
		'',
	].join('\n')
}

/**
 * Run one tier's program for real, and insist it produced a deck.
 *
 * This is what keeps the gate honest in the direction a ratchet cannot afford. esbuild
 * resolves modules; it does not care whether `slide.addChart` exists. So a phase that
 * renames or drops a method leaves every program still bundleable, the family that method
 * reached falls out of the graph, and the number goes **down** — indistinguishable from the
 * win this gate exists to measure. Running the program first turns that into a failure.
 * The decks are tiny and the whole pass costs a fraction of a second.
 * @param {string} tier - a name in {@link TIER_NAMES}
 * @param {string} program - the written program file
 * @returns {Promise<void>}
 */
async function runProgram(tier, program) {
	const global = /** @type {Record<string, ArrayBuffer | undefined>} */ (/** @type {unknown} */ (globalThis))
	// Cleared first, so a tier that runs without writing cannot inherit the previous tier's deck.
	delete global[SINK]
	try {
		await import(pathToFileURL(program).href)
	} catch (error) {
		throw new Error(
			`the \`${tier}\` program does not run against dist/, so its size would be a measurement of a ` +
				`program nobody can write: ${error instanceof Error ? error.message : String(error)}`,
			{ cause: error }
		)
	}
	if (!global[SINK]?.byteLength)
		throw new Error(`the \`${tier}\` program ran but wrote no deck, so it reached none of what it names`)
}

/**
 * Bundle one tier's program and weigh what a browser would fetch.
 *
 * Built in a directory of its own under `.tmp/` so the entry chunk is identifiable by name
 * and two tiers cannot read each other's output. `write: false` keeps the bundles in
 * memory — nothing reads them back, and files left behind from an earlier shape of the
 * graph would sit in `.tmp/` forever.
 * @param {string} tier - a name in {@link TIER_NAMES}
 * @returns {Promise<{initial: number, total: number, chunks: Array<{name: string, bytes: number}>}>}
 */
export async function measureTier(tier) {
	const dir = tierDir(tier)
	fs.mkdirSync(dir, { recursive: true })
	const program = path.join(dir, PROGRAM_FILE)
	fs.writeFileSync(program, programFor(tier))
	await runProgram(tier, program)
	const result = await esbuild.build({
		absWorkingDir: dir,
		bundle: true,
		entryPoints: [PROGRAM_FILE],
		format: 'esm',
		legalComments: 'none',
		metafile: true,
		minify: true,
		outdir: 'bundle',
		platform: 'browser',
		splitting: true,
		target: 'es2024',
		write: false,
	})
	const entryChunk = path.basename(PROGRAM_FILE, '.mjs') + '.js'
	const chunks = result.outputFiles.map((file) => ({
		name: path.basename(file.path),
		bytes: zlib.gzipSync(Buffer.from(file.contents), { level: 9 }).byteLength,
	}))
	const upFront = blockingChunks(result.metafile, entryChunk)
	// Defaulting to zero here would report the cheapest tier the gate can express, and pass.
	if (!chunks.some((chunk) => chunk.name === entryChunk))
		throw new Error(
			`the \`${tier}\` bundle has no ${entryChunk} to charge as the entry chunk; esbuild emitted ` +
				chunks.map((chunk) => chunk.name).join(', ')
		)
	return {
		initial: chunks.reduce((sum, chunk) => (upFront.has(chunk.name) ? sum + chunk.bytes : sum), 0),
		total: chunks.reduce((sum, chunk) => sum + chunk.bytes, 0),
		chunks,
	}
}

/**
 * Every chunk a browser must have before the program's first line runs: the entry chunk and
 * everything reachable from it by `import` statements, stopping at each `import()`.
 *
 * The entry chunk alone is not that figure, and reporting it as one is a trap the gate walked
 * into: the moment a module inside the library defers something with a dynamic import, esbuild
 * splits the shared code out of the entry chunk, and a row that had been reporting ~98 kB
 * reported 0.9 kB for the same program with the same download. Nothing got cheaper; the
 * measurement stopped counting the part that was still being fetched first.
 * @param {import('esbuild').Metafile} metafile
 * @param {string} entryChunk - the entry chunk's base name
 * @returns {Set<string>} chunk base names, entry included
 */
function blockingChunks(metafile, entryChunk) {
	/** @type {Map<string, string>} */
	const byName = new Map(Object.keys(metafile.outputs).map((out) => [path.basename(out), out]))
	const reached = new Set()
	const queue = [entryChunk]
	while (queue.length > 0) {
		const name = queue.pop()
		if (name === undefined || reached.has(name)) continue
		reached.add(name)
		const out = byName.get(name)
		for (const imported of (out && metafile.outputs[out]?.imports) || [])
			if (imported.kind === 'import-statement') queue.push(path.basename(imported.path))
	}
	return reached
}

/** @param {number} bytes */
const kb = (bytes) => (bytes / 1024).toFixed(1) + ' kB'

/**
 * Measure every tier against the built `dist/`.
 * @returns {Promise<Map<string, Awaited<ReturnType<typeof measureTier>>>>}
 */
export async function measureTiers() {
	if (!fs.existsSync(DIST_ENTRY))
		throw new Error(`${path.relative(ROOT, DIST_ENTRY)} is missing — run \`pnpm run build\` first`)
	/** @type {Map<string, Awaited<ReturnType<typeof measureTier>>>} */
	const measured = new Map()
	for (const tier of TIER_NAMES) measured.set(tier, await measureTier(tier))
	return measured
}

/**
 * Compare one measurement against one budget, on the ratchet's terms.
 *
 * Split out from {@link main} because it is the whole of the verdict logic and needs no
 * esbuild to exercise — see `test/scripts/bundle-tier-size.test.js`.
 * @param {number} bytes - what was measured
 * @param {number} budget - what is frozen
 * @returns {'over' | 'under' | 'ok'} `under` means far enough under to be worth banking
 */
export function verdictFor(bytes, budget) {
	if (bytes > budget) return 'over'
	if (bytes < budget * (1 - SLACK_PCT / 100) && budget - bytes >= SLACK_MIN_BYTES) return 'under'
	return 'ok'
}

/**
 * The budget a measurement freezes to: {@link HEADROOM_PCT} above it, rounded up to a whole
 * kB so the file reads as a decision someone made rather than a build artifact copied in.
 * @param {number} bytes
 * @returns {number}
 */
export function frozenBudget(bytes) {
	return Math.ceil((bytes * (1 + HEADROOM_PCT / 100)) / 1024) * 1024
}

// ---------------------------------------------------------------- CLI

const USAGE = `Tier-size budget — what a consumer program actually downloads.

  node scripts/bundle-tier-size.mjs            check (exit 1 over budget)
  node scripts/bundle-tier-size.mjs --freeze   rewrite the budget from dist/
  node scripts/bundle-tier-size.mjs --list     per-chunk breakdown

Tiers: ${TIER_NAMES.join(', ')}. Each bundles a consumer program against dist/browser.js
and records the entry chunk (initial) and every chunk it can reach (total).

Options:
  --freeze    write ${HEADROOM_PCT}% above today's measurement into the budget file
  --list      print every chunk of each tier's bundle, largest first
  -h, --help  show this message`

/** @param {string[]} argv @returns {Promise<number>} process exit code */
export async function main(argv) {
	const { values } = parseCli(argv, {
		usage: USAGE,
		options: {
			freeze: { type: 'boolean', default: false },
			list: { type: 'boolean', default: false },
		},
	})

	const measured = await measureTiers()

	if (values.list) {
		for (const [tier, { initial, total, chunks }] of measured) {
			console.log(`${tier}: ${kb(initial)} initial, ${kb(total)} total, ${chunks.length} chunk(s)`)
			for (const chunk of [...chunks].sort((a, b) => b.bytes - a.bytes))
				console.log(`  ${kb(chunk.bytes).padStart(9)}  ${chunk.name}`)
		}
		return 0
	}

	if (values.freeze) {
		/** @type {Record<string, Record<string, number>>} */
		const frozen = {}
		for (const [tier, measurement] of measured) {
			frozen[tier] = Object.fromEntries(FIGURES.map((figure) => [figure, frozenBudget(measurement[figure])]))
			console.log(
				`bundle tier: froze ${tier} at ${kb(frozen[tier].initial ?? 0)} initial / ${kb(frozen[tier].total ?? 0)} total` +
					` (measured ${kb(measurement.initial)} / ${kb(measurement.total)})`
			)
		}
		fs.writeFileSync(BUDGET, JSON.stringify(frozen, null, '\t') + '\n')
		return 0
	}

	/** @type {Record<string, Record<string, number> | undefined>} */
	const budgetFile = JSON.parse(fs.readFileSync(BUDGET, 'utf8'))
	const relBudget = path.relative(ROOT, BUDGET).split(path.sep).join('/')

	const missing = [...measured.keys()].filter((tier) =>
		FIGURES.some((figure) => typeof budgetFile[tier]?.[figure] !== 'number')
	)
	if (missing.length) {
		console.error(`bundle tier FAILED — no budget for ${missing.join(', ')} in ${relBudget}.`)
		console.error('\n  pnpm run bundle-tier:freeze')
		return 1
	}

	const checked = [...measured].flatMap(([tier, measurement]) =>
		FIGURES.map((figure) => ({
			label: `${tier} ${figure}`,
			tier,
			bytes: measurement[figure],
			budget: budgetFile[tier]?.[figure] ?? 0,
			chunks: measurement.chunks,
		}))
	)
	const over = checked.filter((row) => verdictFor(row.bytes, row.budget) === 'over')
	const under = checked.filter((row) => verdictFor(row.bytes, row.budget) === 'under')

	if (over.length) {
		console.error('bundle tier FAILED — a consumer program grew past its budget:\n')
		for (const { label, bytes, budget, chunks } of over) {
			console.error(`  ${label}: ${kb(bytes)} (budget ${kb(budget)}), ${chunks.length} chunk(s)`)
			for (const chunk of [...chunks].sort((a, b) => b.bytes - a.bytes).slice(0, 5))
				console.error(`    ${kb(chunk.bytes).padStart(9)}  ${chunk.name}`)
		}
		console.error('\nRun `pnpm run bundle-tier:list` for the full breakdown. If the growth is')
		console.error(`intended, raise the number in ${relBudget} in the same commit and say what`)
		console.error('bought the bytes; if it is not, a construct family reached a tier that should')
		console.error('not be able to see it.')
		return 1
	}

	if (under.length) {
		console.log(`bundle tier: ${SLACK_PCT}%+ under budget — bank it by lowering ${relBudget}:\n`)
		for (const { label, bytes, budget } of under) console.log(`  ${label}: ${kb(budget)} -> ${kb(bytes)}`)
		console.log('\n  pnpm run bundle-tier:freeze')
		return 1
	}

	for (const [tier, { initial, total, chunks }] of measured)
		console.log(
			`bundle tier: ok — ${tier} ${kb(initial)} initial / ${kb(total)} total` +
				` (budget ${kb(budgetFile[tier]?.initial ?? 0)} / ${kb(budgetFile[tier]?.total ?? 0)}), ${chunks.length} chunk(s)`
		)
	return 0
}

if (isMain(import.meta.url)) await runCli(() => main(process.argv.slice(2)))
