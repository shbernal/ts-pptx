#!/usr/bin/env node
/**
 * Export-reachability sweep — which `export`s in `src/` nothing else in the repo names, and
 * which of those reach a consumer anyway through the published `.d.ts`.
 *
 * Why the second half matters. The run this was written against reported 48 exports that no
 * other tracked file mentions, and 40 of them were the public type surface (`types/chart.ts`,
 * `inspect.ts`, `clip.ts`, `types/core.ts`) — correct as they are, since a published type has
 * no in-repo caller by construction. A report that is 83% noise gets run once and never again,
 * so the bucketing is not a refinement of this tool: it is the tool. Those two numbers are
 * that run and not a promise — they move with every export added, dropped or wired up, which
 * is the whole reason this reports rather than freezes a count.
 *
 * Method, in three steps:
 *
 *   1. Collect every `export <kind> NAME` declaration in a `.ts` file under `src/`.
 *   2. Test `\bNAME\b` against every tracked file except the one that declares it. A name
 *      that hits is referenced and is not reported.
 *   3. Bucket each survivor by where it appears in the built `dist/`:
 *        PUBLISHED — named in an entry `.d.ts` (one of `package.json`'s `exports` types).
 *                    A consumer can import it; nothing to do.
 *        CHUNK     — only in a shared `*-<hash>.d.ts`. Reachable structurally (it is the
 *                    type of something published) but not importable by name.
 *        INTERNAL  — nowhere in `dist/`. The bucket worth reading: an export with no
 *                    in-repo caller and no published surface.
 *
 * Two traps this encodes, both of which produced a wrong answer by hand:
 *
 *   - **Same-file uses count.** `OutputType` looks unreferenced until you notice
 *     `ZIP_OUTPUT_TYPE` is derived from it two lines below. A survivor that its own file
 *     uses is marked `self-used`, and that is the difference between "delete it" and
 *     "drop the `export` keyword".
 *   - **`export` is sometimes load-bearing against a lint**, not for a consumer. An
 *     INTERNAL survivor that is self-used is very often this. The sweep reports; it does
 *     not conclude.
 *
 * Diagnostic, not a gate: the last step is bucketing a survivor by intent, which is a
 * human's call. It exits 0 on any finding.
 */
import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { ROOT, parseCliOrExit } from './script-utils.mjs'

/** A newline, so a template literal never has to carry a bare one and lose its indentation. */
const LF = '\n'

const USAGE = `Which exports in src/ nothing else names, bucketed by published surface.

  node scripts/export-reach.mjs            report the unreferenced exports, by bucket
  node scripts/export-reach.mjs --all      also report how many references each export has

Options:
  --bucket <name>  report one bucket only: published, chunk or internal
  --all            list every export with its reference count, not only the survivors
  -h, --help       show this message`

const { values: flags } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: { all: { type: 'boolean', default: false }, bucket: { type: 'string' } },
})

/**
 * `export <kind> NAME`, for the kinds that introduce a binding of their own.
 *
 * Re-export lists (`export { a, b } from './x.js'`) and `export *` are deliberately not
 * matched: they name no new binding, and the name they forward is already reported at the
 * file that declares it.
 */
const EXPORT_DECL =
	/^export\s+(?:declare\s+)?(?:abstract\s+)?(?:async\s+)?(function|const|let|var|class|interface|type|enum|namespace)\s+([A-Za-z_$][\w$]*)/gm

/** Extensions worth scanning for a reference. Everything else in the tree is a binary or a lockfile. */
const TEXT_EXT = new Set([
	'.ts',
	'.tsx',
	'.mts',
	'.cts',
	'.js',
	'.mjs',
	'.cjs',
	'.jsx',
	'.json',
	'.jsonc',
	'.md',
	'.mts',
	'.vue',
	'.yml',
	'.yaml',
	'.html',
	'.ps1',
	'.py',
	'.txt',
	'.tsv',
])

/** Every tracked file, repo-relative and slash-separated. @returns {string[]} */
function trackedFiles() {
	return execFileSync('git', ['ls-files'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 })
		.split('\n')
		.map((line) => line.trim())
		.filter(Boolean)
}

/**
 * The identifiers a file contains, as a set.
 *
 * A token set rather than a regex per name: the sweep asks ~1,500 questions of ~900 files,
 * and `\bNAME\b` over each pair is a minute of work for an answer that one pass over each
 * file already holds. Splitting on the identifier boundary gives `\b` semantics exactly.
 * @param {string} text
 * @returns {Set<string>}
 */
function identifiersOf(text) {
	return new Set(text.split(/[^A-Za-z0-9_$]+/u).filter(Boolean))
}

/**
 * How many times a name appears in one file's identifier stream.
 * @param {string} text
 * @param {string} name
 * @returns {number}
 */
function countIn(text, name) {
	let hits = 0
	for (const token of text.split(/[^A-Za-z0-9_$]+/u)) if (token === name) hits++
	return hits
}

/** The entry `.d.ts` files `package.json` publishes, plus every shared chunk beside them. */
function distSurfaces() {
	const distDir = path.join(ROOT, 'dist')
	if (!fs.existsSync(distDir)) {
		console.error('Missing dist/. Run `pnpm run build` first — the buckets are read off the built types.')
		process.exit(1)
	}
	const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
	/** @type {Set<string>} */
	const entries = new Set()
	/** @param {unknown} node */
	const walkExports = (node) => {
		if (typeof node === 'string') {
			if (node.endsWith('.d.ts')) entries.add(path.basename(node))
			return
		}
		if (node && typeof node === 'object') for (const value of Object.values(node)) walkExports(value)
	}
	walkExports(manifest.exports)

	const types = fs.readdirSync(distDir).filter((name) => name.endsWith('.d.ts'))
	const read = (/** @type {string} */ name) => fs.readFileSync(path.join(distDir, name), 'utf8')
	return {
		published: new Set(types.filter((name) => entries.has(name)).flatMap((name) => [...identifiersOf(read(name))])),
		chunk: new Set(types.filter((name) => !entries.has(name)).flatMap((name) => [...identifiersOf(read(name))])),
	}
}

/** Every `.ts` under `src/`, repo-relative. @param {string[]} tracked @returns {string[]} */
function sourceFiles(tracked) {
	return tracked.filter((file) => file.startsWith('src/') && file.endsWith('.ts'))
}

function main() {
	const tracked = trackedFiles()
	const sources = sourceFiles(tracked)
	const surfaces = distSurfaces()

	// One read per file, reused for every name. `tracked` includes files this process cannot
	// usefully scan (images, fonts); those are dropped by extension rather than by sniffing.
	/** @type {Map<string, string>} */
	const contents = new Map()
	for (const file of tracked) {
		if (!TEXT_EXT.has(path.extname(file))) continue
		try {
			contents.set(file, fs.readFileSync(path.join(ROOT, file), 'utf8'))
		} catch {
			// A tracked path that is not readable as text tells us nothing about a reference.
		}
	}
	/** @type {Map<string, Set<string>>} */
	const tokens = new Map()
	for (const [file, text] of contents) tokens.set(file, identifiersOf(text))

	/** @type {{name: string, kind: string, file: string, refs: number, selfUses: number}[]} */
	const exports = []
	for (const file of sources) {
		const text = contents.get(file)
		if (!text) continue
		for (const match of text.matchAll(EXPORT_DECL)) {
			const [, kind, name] = match
			if (!name || !kind) continue
			let refs = 0
			for (const [other, set] of tokens) if (other !== file && set.has(name)) refs++
			// The declaration itself is one occurrence; anything beyond it is a use.
			exports.push({ name, kind, file, refs, selfUses: countIn(text, name) - 1 })
		}
	}
	exports.sort((a, b) => a.name.localeCompare(b.name))

	/** @param {string} name */
	const bucketOf = (name) =>
		surfaces.published.has(name) ? 'published' : surfaces.chunk.has(name) ? 'chunk' : 'internal'

	const survivors = exports.filter((entry) => entry.refs === 0)
	/** @type {Record<string, typeof survivors>} */
	const buckets = { published: [], chunk: [], internal: [] }
	for (const entry of survivors) buckets[bucketOf(entry.name)]?.push(entry)

	if (flags.all) {
		console.log(`# every export in src/ (${exports.length})\n`)
		for (const entry of exports)
			console.log(
				`${String(entry.refs).padStart(4)} file(s)  ${bucketOf(entry.name).padEnd(9)} ${entry.kind.padEnd(9)} ${entry.name}  ${entry.file}`
			)
		console.log('')
	}

	/** @type {Record<string, string>} */
	const titles = {
		published: 'PUBLISHED — named in an entry .d.ts, so a consumer can import it. Correct as they are.',
		chunk: 'CHUNK — only in a shared chunk .d.ts: reachable structurally, not importable by name.',
		internal: 'INTERNAL — nowhere in dist/. No in-repo caller and no published surface.',
	}
	const wanted = flags.bucket ? [String(flags.bucket)] : ['internal', 'chunk', 'published']
	for (const key of wanted) {
		const rows = buckets[key]
		if (!rows) {
			console.error(`unknown bucket "${key}"; expected published, chunk or internal`)
			return 2
		}
		console.log(`## ${titles[key]}`)
		console.log(`   ${rows.length} of ${survivors.length} unreferenced export(s)` + LF)
		for (const entry of rows)
			console.log(
				`  ${entry.kind.padEnd(9)} ${entry.name}${entry.selfUses > 0 ? ` (self-used ${entry.selfUses}x)` : ''}` +
					LF +
					`             ${entry.file}`
			)
		console.log('')
	}

	console.log(
		`${exports.length} export(s) in ${sources.length} source file(s); ${survivors.length} named by no other tracked file.`
	)
	return 0
}

process.exitCode = main()
