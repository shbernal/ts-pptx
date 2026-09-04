/**
 * What each library costs to install and to ship, from clean per-library installs.
 *
 * Every number here comes out of a temp directory holding one `npm install` and nothing
 * else: upstream from the registry (the install `measure.mjs` already made), ours from a
 * `pnpm pack` of the working tree. Measuring our own side out of the repo instead would
 * measure the development tree — `node_modules` hoisted flat, `dist/` sitting beside
 * sources that never ship — and none of that is what a consumer gets.
 *
 * ## The hello-world bundle, and why it is not the ratchet's number
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
 * `splitting: true` is on and both figures are recorded. `initialBytes` is what the
 * program pays to start; `totalBytes` is every chunk it can reach. They differ on our side
 * because font metrics load `opentype.js` through a dynamic import that only runs on first
 * font registration, and a bundler that can defer that will. Reporting only the total
 * would charge the program for a chunk it may never fetch; reporting only the initial
 * would hide bytes it might. Neither is a fair single number, so there is no single
 * number.
 *
 * The program is identical in intent on both sides — one slide, one text box, export — and
 * expressed in each library's own idiom, on the same rule the probe corpus follows.
 */
import fs from 'node:fs'
import path from 'node:path'
import zlib from 'node:zlib'
import esbuild from 'esbuild'
import { packPackage } from '../pack-utils.mjs'
import { run } from '../script-utils.mjs'
import { unavailable } from './unavailable.mjs'

/** Our package's name on npm, and the directory it installs into. */
const SELF_PACKAGE = 'pptx-ts'

/**
 * The consumer program, per library. One slide, one text box, export.
 *
 * The result is passed to `console.log` on purpose: an export whose value is discarded is
 * dead code, and a minifier that proves it can drop the whole call graph behind it. That
 * turns the row into a measurement of how well each library annotates side effects, which
 * is not the question being asked.
 * @type {Record<string, string>}
 */
const PROGRAMS = {
	'ts-pptx': [
		"import TsPptx from 'pptx-ts'",
		'const pres = new TsPptx()',
		"pres.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })",
		"console.log(await pres.write({ outputType: 'arraybuffer' }))",
	].join('\n'),
	pptxgenjs: [
		"import PptxGenJS from 'pptxgenjs'",
		'const pres = new PptxGenJS()',
		"pres.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })",
		"console.log(await pres.write({ outputType: 'arraybuffer' }))",
	].join('\n'),
}

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
 * @param {string} program - the consumer source
 * @returns {Promise<{initialBytes: number, totalBytes: number, chunks: number} | import('./unavailable.mjs').Unavailable>}
 */
async function bundleHelloWorld(prefix, program) {
	const entry = 'hello-world.mjs'
	fs.writeFileSync(path.join(prefix, entry), program + '\n')
	try {
		const result = await esbuild.build({
			absWorkingDir: prefix,
			bundle: true,
			entryPoints: [entry],
			format: 'esm',
			legalComments: 'none',
			minify: true,
			outdir: 'hello-world-bundle',
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
			if (path.basename(file.path) === 'hello-world.js') initialBytes = bytes
		}
		return { initialBytes, totalBytes, chunks: result.outputFiles.length }
	} catch (error) {
		return unavailable('esbuild could not bundle the consumer program: ' + messageOf(error))
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
 * @param {string} prefix - the install directory, holding one `node_modules`
 * @param {string} root - the measured package's own directory inside it
 * @param {string} program - the consumer program for this library
 * @returns {Promise<Record<string, unknown>>}
 */
async function measureInstall(prefix, root, program) {
	const nodeModules = path.join(prefix, 'node_modules')
	const manifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'))
	const installed = installedPackages(nodeModules)
	installed.delete(manifest.name)

	return {
		install: { bytes: treeBytes(nodeModules), packageBytes: treeBytes(root) },
		dependencies: {
			direct: Object.keys(manifest.dependencies ?? {}).sort(),
			transitive: installed.size,
		},
		entryPoints: entrySubpaths(manifest.exports),
		moduleFormats: moduleFormats(manifest),
		engines: manifest.engines?.node ?? null,
		helloWorld: await bundleHelloWorld(prefix, program),
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
 * @param {object} opts
 * @param {string} opts.workDir
 * @param {string} opts.upstreamRoot - the installed `pptxgenjs` directory `measure.mjs` made
 * @param {boolean} [opts.reuse] - skip packing and reinstalling our own side
 * @returns {Promise<Record<string, unknown>>}
 */
export async function measureHygiene({ workDir, upstreamRoot, reuse = false }) {
	const selfRoot = await installSelf(workDir, reuse)
	/** @type {Record<string, {prefix: string, root: string}>} */
	const installs = {
		'ts-pptx': { prefix: path.join(workDir, 'self'), root: selfRoot },
		pptxgenjs: { prefix: path.join(workDir, 'upstream'), root: upstreamRoot },
	}

	/** @type {Record<string, unknown>} */
	const hygiene = {}
	for (const subject of Object.keys(PROGRAMS)) {
		const { prefix, root } = forSubject(installs, subject)
		hygiene[subject] = await measureInstall(prefix, root, forSubject(PROGRAMS, subject))
	}
	return hygiene
}
