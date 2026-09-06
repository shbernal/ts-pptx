/**
 * What each library costs to install and to ship, from clean per-library installs.
 *
 * Every number here comes out of a temp directory holding one `npm install` and nothing
 * else: upstream from the registry (the install `measure.mjs` already made), ours from a
 * `pnpm pack` of the working tree. Measuring our own side out of the repo instead would
 * measure the development tree — `node_modules` hoisted flat, `dist/` sitting beside
 * sources that never ship — and none of that is what a consumer gets.
 *
 * ## The bundled programs, and why they are not the ratchet's numbers
 *
 * `scripts/bundle-size-ratchet.mjs` documents the measurement conventions this follows:
 * minify before measuring, because half of `dist/` by weight is doc comments that no
 * consumer's build keeps, and gzip at level 9. Read that header for why, rather than
 * re-deriving it here.
 *
 * One convention is deliberately inverted. The ratchet measures per file and never
 * bundles, so it cannot tree-shake across the closure and its figure is an upper bound on
 * what the package ships. This bundles — `bundle: true`, one real consumer program as the
 * entry — because a consumer's bundle is precisely the thing being compared, and a
 * comparison that forbade tree-shaking would credit whichever library happens to export
 * less per module rather than whichever costs less to use. **The two numbers will not
 * match, and the page has to say so**, because someone will hold them side by side.
 *
 * `splitting: true` is on and both figures are recorded, but **the page prints only
 * `initialBytes`** — what the program pays before its first line runs. `totalBytes` is
 * every chunk it can reach, and it is recorded here for anyone auditing the split rather
 * than published as a comparison row, because as a row it compares nothing. Our two
 * deferred chunks are `opentype.js`, behind the dynamic import in
 * `measure/font-metrics.ts` that only runs on first font registration, and the default
 * video poster in `media/playbtn.ts`. Upstream has no counterpart to the first and ships
 * the second inside its entry chunk, so a `totalBytes` row would set our sum of both
 * against a number that answers neither.
 *
 * ## Why more than one program
 *
 * A hello world is one point on a curve, and it is the point where two tree-shaken bundles
 * look most alike: almost everything either library holds has been shaken out of it. The
 * corpus in `./programs.mjs` climbs from there to a deck using every construct the shared
 * baseline shows both libraries emitting, so the page can say what the second slide costs
 * as well as the first. Every program is identical in intent on both sides and written in
 * each library's own idiom, on the rule the probe corpus follows.
 *
 * Each program is **run before it is bundled**. esbuild compiles a call that does not
 * exist, and a misspelled method comes out as a *smaller* bundle rather than as an error,
 * because the tree-shaker keeps less. Executing first is what stops a typo being published
 * as a saving.
 */
import fs from 'node:fs'
import path from 'node:path'
import zlib from 'node:zlib'
import esbuild from 'esbuild'
import { packPackage } from '../pack-utils.mjs'
import { run } from '../script-utils.mjs'
import { PROGRAMS, programArm, programFrame, programModule, programSource, resetProgramData } from './programs.mjs'
import { unavailable } from './unavailable.mjs'

/** Our package's name on npm, and the directory it installs into. */
const SELF_PACKAGE = 'pptx-ts'

/**
 * Total bytes under a directory, following no symlinks.
 *
 * `du`'s apparent size rather than its on-disk size: block rounding is a property of the
 * filesystem the measurement happened to run on, and two libraries measured on different
 * machines have to be comparable.
 * @param {string} dir
 * @returns {number}
 */
function treeBytes(dir) {
	let total = 0
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		const full = path.join(dir, entry.name)
		if (entry.isSymbolicLink()) continue
		if (entry.isDirectory()) total += treeBytes(full)
		else if (entry.isFile()) total += fs.statSync(full).size
	}
	return total
}

/**
 * Every package name present in an install tree, nested `node_modules` included.
 *
 * npm hoists what it can, so most of these sit flat, but a version conflict puts a copy
 * under its dependent and a flat `readdir` would miss it. Scoped directories are not
 * packages — `@types` holds them — so they are descended into rather than counted.
 * @param {string} nodeModules
 * @returns {Set<string>}
 */
function installedPackages(nodeModules) {
	/** @type {Set<string>} */
	const names = new Set()
	if (!fs.existsSync(nodeModules)) return names
	for (const entry of fs.readdirSync(nodeModules, { withFileTypes: true })) {
		if (!entry.isDirectory() || entry.name === '.bin') continue
		const full = path.join(nodeModules, entry.name)
		if (entry.name.startsWith('@')) {
			for (const scoped of fs.readdirSync(full, { withFileTypes: true })) {
				if (!scoped.isDirectory()) continue
				names.add(entry.name + '/' + scoped.name)
				for (const nested of installedPackages(path.join(full, scoped.name, 'node_modules'))) names.add(nested)
			}
			continue
		}
		names.add(entry.name)
		for (const nested of installedPackages(path.join(full, 'node_modules'))) names.add(nested)
	}
	return names
}

/**
 * The subpaths a manifest's `exports` publishes.
 *
 * Two shapes are legal and both are in play here: a map of subpath keys (ours, eleven of
 * them) and a bare conditions object with no subpaths at all (upstream's, which publishes
 * `.` and only `.`). Telling them apart is a `.` prefix on the first key, which is what the
 * resolution algorithm itself keys on.
 *
 * `./package.json` is dropped. It is a declared subpath, and tooling does reach for it, but
 * it is not an entry point in the sense the row is asking about, and leaving it in makes a
 * one-entry package look like a two-entry one.
 * @param {Record<string, unknown> | undefined} exports
 * @returns {string[]}
 */
function entrySubpaths(exports) {
	if (!exports || typeof exports !== 'object') return []
	const keys = Object.keys(exports)
	if (!keys.some((key) => key.startsWith('.'))) return ['.']
	return keys.filter((key) => key !== './package.json')
}

/**
 * Which module formats a consumer can import, read off the manifest.
 *
 * From the conditions actually present rather than from `type`: `type` says how bare `.js`
 * is parsed, which is not the same question as what the package offers. A package with a
 * `require` condition ships CJS whatever its `type` field says, and one whose only entry is
 * an unconditional path ships whatever `type` makes it — which is the one case `type` does
 * settle, and the only one it is consulted for.
 * @param {Record<string, unknown>} manifest
 * @returns {string[]} sorted, e.g. `['cjs', 'esm']`
 */
function moduleFormats(manifest) {
	/** @type {Set<string>} */
	const formats = new Set()
	/** @param {unknown} node */
	const walk = (node) => {
		if (!node || typeof node !== 'object') return
		for (const [key, value] of Object.entries(node)) {
			if (key === 'import' || key === 'module') formats.add('esm')
			if (key === 'require') formats.add('cjs')
			walk(value)
		}
	}
	walk(manifest.exports)
	if (typeof manifest.module === 'string') formats.add('esm')
	if (typeof manifest.main === 'string' && !formats.has('cjs')) formats.add(manifest.type === 'module' ? 'esm' : 'cjs')
	if (formats.size === 0) formats.add(manifest.type === 'module' ? 'esm' : 'cjs')
	return [...formats].sort()
}

/**
 * Bundle one consumer program and weigh what a browser would fetch.
 *
 * Built inside the install prefix so bare specifiers resolve to that install and nothing
 * else — the repo's own `node_modules` is a hoisted development tree and must never be on
 * this resolution path. `platform: 'browser'` because that is where a bundle size is a cost
 * a user pays, and because it is the setting that honours each manifest's `browser` field.
 * @param {string} prefix - the install directory
 * @param {import('./programs.mjs').Program} program
 * @param {string} subject
 * @returns {Promise<{initialBytes: number, totalBytes: number, chunks: number} | import('./unavailable.mjs').Unavailable>}
 */
async function bundleProgram(prefix, program, subject) {
	const entry = program.id + '.mjs'
	fs.writeFileSync(path.join(prefix, entry), programModule(program, subject))
	try {
		const result = await esbuild.build({
			absWorkingDir: prefix,
			bundle: true,
			entryPoints: [entry],
			format: 'esm',
			legalComments: 'none',
			minify: true,
			outdir: program.id + '-bundle',
			platform: 'browser',
			splitting: true,
			target: 'es2024',
			write: false,
		})
		let totalBytes = 0
		let initialBytes = 0
		for (const file of result.outputFiles) {
			const bytes = zlib.gzipSync(Buffer.from(file.contents), { level: 9 }).byteLength
			totalBytes += bytes
			if (path.basename(file.path) === program.id + '.js') initialBytes = bytes
		}
		return { initialBytes, totalBytes, chunks: result.outputFiles.length }
	} catch (error) {
		return unavailable('esbuild could not bundle the ' + program.id + ' program: ' + messageOf(error))
	}
}

/**
 * Build every program with the real library before any of them is bundled.
 *
 * A throw rather than a recorded outcome, because this is not a measurement: a program the
 * corpus cannot run is an error in the corpus, and the sizes it would produce are sizes for
 * a deck nobody can build. The check is what makes the bundle figures mean what the page
 * says they mean, since esbuild reports nothing about a call that does not exist.
 * @param {Record<string, () => any>} subjects - a fresh presentation per library
 * @returns {Promise<void>}
 */
async function checkPrograms(subjects) {
	for (const program of PROGRAMS)
		for (const [subject, construct] of Object.entries(subjects)) {
			// Every arm sees the data the corpus declares, not what the previous library left in
			// it. See `./corpus-data.mjs`.
			resetProgramData()
			const pres = construct()
			try {
				await programArm(program, subject)(pres)
				await pres.write({ outputType: 'arraybuffer' })
			} catch (error) {
				throw new Error(
					'the bundle program "' + program.id + '" does not build with ' + subject + ': ' + messageOf(error),
					{ cause: error }
				)
			}
		}
}

/** @param {unknown} error @returns {string} */
function messageOf(error) {
	return error instanceof Error ? error.message : String(error)
}

/**
 * Install our own packed tarball into a clean prefix, as a consumer would get it.
 * @param {string} workDir
 * @param {boolean} reuse - use an existing install instead of packing and reinstalling
 * @returns {Promise<string>} the installed package's root
 */
async function installSelf(workDir, reuse) {
	const prefix = path.join(workDir, 'self')
	const root = path.join(prefix, 'node_modules', SELF_PACKAGE)
	if (!reuse) {
		const { tarball } = await packPackage(path.join(workDir, 'pack'))
		fs.mkdirSync(prefix, { recursive: true })
		await run('npm', ['install', tarball, '--prefix', prefix, '--no-audit', '--no-fund'])
	}
	if (!fs.existsSync(root)) throw new Error(SELF_PACKAGE + ' is not installed at ' + root)
	return root
}

/**
 * Everything one clean install has to say about itself.
 *
 * Bundles are keyed by program id rather than listed, so the page pairs the two columns by
 * id: a corpus reordered between two measurements must not silently pair a chart deck with
 * a table one.
 * @param {string} prefix - the install directory, holding one `node_modules`
 * @param {string} root - the measured package's own directory inside it
 * @param {string} subject - which library this install holds
 * @returns {Promise<Record<string, unknown>>}
 */
async function measureInstall(prefix, root, subject) {
	const nodeModules = path.join(prefix, 'node_modules')
	const manifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'))
	const installed = installedPackages(nodeModules)
	installed.delete(manifest.name)

	/** @type {Record<string, unknown>} */
	const bundles = {}
	for (const program of PROGRAMS) bundles[program.id] = await bundleProgram(prefix, program, subject)

	return {
		install: { bytes: treeBytes(nodeModules), packageBytes: treeBytes(root) },
		dependencies: {
			direct: Object.keys(manifest.dependencies ?? {}).sort(),
			transitive: installed.size,
		},
		entryPoints: entrySubpaths(manifest.exports),
		moduleFormats: moduleFormats(manifest),
		engines: manifest.engines?.node ?? null,
		bundles,
	}
}

/**
 * One entry from a per-subject map, or a throw naming the subject.
 * @template T
 * @param {Record<string, T>} map
 * @param {string} subject
 * @returns {T}
 */
function forSubject(map, subject) {
	const value = map[subject]
	if (value === undefined) throw new Error('nothing registered for subject "' + subject + '"')
	return value
}

/**
 * Install both libraries clean and measure what each costs.
 *
 * `programs` sits beside the two subject keys rather than inside them: the corpus is one
 * set of decks, and duplicating each program's description under both libraries would let
 * the two copies disagree about what was measured.
 * @param {object} opts
 * @param {string} opts.workDir
 * @param {string} opts.upstreamRoot - the installed `pptxgenjs` directory `measure.mjs` made
 * @param {Record<string, () => any>} opts.subjects - a fresh presentation per library
 * @param {boolean} [opts.reuse] - skip packing and reinstalling our own side
 * @returns {Promise<Record<string, unknown>>}
 */
export async function measureHygiene({ workDir, upstreamRoot, subjects, reuse = false }) {
	await checkPrograms(subjects)
	const selfRoot = await installSelf(workDir, reuse)
	/** @type {Record<string, {prefix: string, root: string}>} */
	const installs = {
		'ts-pptx': { prefix: path.join(workDir, 'self'), root: selfRoot },
		pptxgenjs: { prefix: path.join(workDir, 'upstream'), root: upstreamRoot },
	}

	/** @type {Record<string, unknown>} */
	const hygiene = {
		frame: Object.fromEntries(Object.keys(subjects).map((subject) => [subject, programFrame(subject)])),
		programs: PROGRAMS.map((program) => ({
			id: program.id,
			label: program.label,
			what: program.what,
			source: Object.fromEntries(Object.keys(subjects).map((subject) => [subject, programSource(program, subject)])),
		})),
	}
	for (const subject of Object.keys(subjects)) {
		const { prefix, root } = forSubject(installs, subject)
		hygiene[subject] = await measureInstall(prefix, root, subject)
	}
	return hygiene
}
