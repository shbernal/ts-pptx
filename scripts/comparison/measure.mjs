#!/usr/bin/env node
/**
 * Measure construct coverage by running both libraries and reading the bytes they emit.
 *
 * For every probe in `./probes.mjs`, this builds a deck with each library, opens the
 * package, and looks for the probe's construct in the probe's part. The result is one of
 * four outcomes per library, and `scripts/comparison/snapshot.json` is the committed
 * record of them. No row on the finished page comes from anywhere else.
 *
 * ## Why upstream is installed rather than depended on
 *
 * `npm install pptxgenjs@latest --prefix <tmpdir>` at measure time, deliberately not a
 * devDependency: it stays out of `pnpm-lock.yaml` and out of Renovate's way, and the
 * snapshot then records whichever version was actually measured rather than whichever one
 * a lockfile happens to pin. The resolved version and its publish date go in the snapshot
 * header, so a reader can tell how old the measurement is without trusting this file.
 *
 * ## What fails the run
 *
 * Three things, all of them cases where a quiet pass would be worse than a loud stop:
 *
 *   - **ts-pptx builds a deck and the construct is not in it.** A regression in our own
 *     output must fail here rather than become a comparison row that says we do not
 *     support tables. This is wider than the shared baseline on purpose: the baseline is
 *     where it matters most, but there is no probe where we would rather find out from
 *     the published page.
 *   - **A `no-api` claim the bundle contradicts.** Every `no-api` is checked by grepping
 *     that library's shipped bundle for the construct token. A hit means either the claim
 *     is wrong or the token is there for some other reason, and the second case has to be
 *     written down in the probe's `sightings` to pass.
 *   - **A `sightings` entry nothing sighted.** An acknowledgement that has outlived the
 *     thing it acknowledged is a comment that has become false, and the next reader will
 *     believe it.
 *
 * The built decks stay on disk under the deck directory and their paths come back from
 * {@link measure}; the validity and install-size measurements read exactly these rather
 * than rebuilding a second corpus that would drift from this one.
 */
import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { isMain, parseCli, ROOT, run, runCli } from '../script-utils.mjs'
import { PROBES, SUBJECTS } from './probes.mjs'

const SNAPSHOT = path.join(ROOT, 'scripts', 'comparison', 'snapshot.json')
const DEFAULT_WORK_DIR = path.join(ROOT, '.tmp', 'comparison')

const USAGE = `Usage: node scripts/comparison/measure.mjs [options]

Builds every probe with both libraries and writes scripts/comparison/snapshot.json.

Options:
  --probe <id>        measure one probe only; does not write the snapshot
  --work-dir <dir>    where the upstream install and the built decks go
                      (default: .tmp/comparison)
  --reuse-upstream    skip the npm install and use whatever is already in the work dir
  --out <file>        snapshot path (default: scripts/comparison/snapshot.json)
  -h, --help          show this message`

/**
 * Where each library's shipped bundle lives, for the `no-api` verification grep.
 *
 * Both are the *published* artifact rather than the source tree: what a consumer can reach
 * is what the claim is about, and ts-pptx's `dist/` is many chunks while upstream's is one
 * file per module format.
 * @param {string} upstreamRoot
 * @returns {Record<string, string[]>}
 */
function shippedBundles(upstreamRoot) {
	const dist = path.join(ROOT, 'dist')
	return {
		'ts-pptx': fs
			.readdirSync(dist)
			.filter((name) => name.endsWith('.js'))
			.map((name) => path.join(dist, name)),
		pptxgenjs: ['pptxgen.es.js', 'pptxgen.cjs.js'].map((name) => path.join(upstreamRoot, 'dist', name)),
	}
}

/** fflate's `unzipSync`, loaded by URL out of the repo's own node_modules, as the other package tooling does. */
async function unzipSync() {
	const fflate = await import(pathToFileURL(path.join(ROOT, 'node_modules', 'fflate', 'esm', 'browser.js')).href)
	return fflate.unzipSync
}

/**
 * Install upstream into `workDir` and report what npm resolved.
 * @param {string} workDir
 * @param {boolean} reuse - use an existing install instead of reinstalling
 * @returns {Promise<{root: string, version: string, published: string | null}>}
 */
async function installUpstream(workDir, reuse) {
	const root = path.join(workDir, 'upstream', 'node_modules', 'pptxgenjs')
	if (!reuse) {
		fs.mkdirSync(path.join(workDir, 'upstream'), { recursive: true })
		await run('npm', [
			'install',
			'pptxgenjs@latest',
			'--prefix',
			path.join(workDir, 'upstream'),
			'--no-audit',
			'--no-fund',
		])
	}
	const manifest = path.join(root, 'package.json')
	if (!fs.existsSync(manifest)) throw new Error('pptxgenjs is not installed at ' + manifest)
	const { version } = JSON.parse(fs.readFileSync(manifest, 'utf8'))
	return { root, version, published: await publishDate(version) }
}

/**
 * When the registry says a version was published. `null` rather than a throw when the
 * registry is unreachable: an offline `--reuse-upstream` run should still produce a
 * snapshot, with the field visibly empty rather than invented.
 * @param {string} version
 * @returns {Promise<string | null>}
 */
async function publishDate(version) {
	try {
		const { stdout } = await run('npm', ['view', 'pptxgenjs@' + version, 'time', '--json'], { capture: true })
		const times = JSON.parse(stdout)
		const stamp = typeof times === 'string' ? times : times?.[version]
		return typeof stamp === 'string' ? stamp.slice(0, 10) : null
	} catch {
		return null
	}
}

/**
 * A fresh presentation from each library, as a consumer would construct one.
 * @param {string} upstreamRoot
 * @returns {Promise<Record<string, () => any>>}
 */
async function loadSubjects(upstreamRoot) {
	const tsPptx = await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)
	const upstream = await import(pathToFileURL(path.join(upstreamRoot, 'dist', 'pptxgen.es.js')).href)
	const TsPptx = tsPptx.default
	const PptxGenJS = upstream.default
	return {
		'ts-pptx': () => new TsPptx(),
		pptxgenjs: () => new PptxGenJS(),
	}
}

/**
 * Build one probe with one library and read the result.
 *
 * A build that throws is an outcome, not a crash: `error` carries the message so the page
 * can print what the library actually said instead of paraphrasing it.
 * @param {import('./probes.mjs').Probe} probe
 * @param {string} subject
 * @param {() => any} construct - makes an empty presentation
 * @param {string} deckDir
 * @returns {Promise<{outcome: string, message?: string, deck?: string}>}
 */
async function measureProbe(probe, subject, construct, deckDir) {
	const build = probe.build[subject]
	if (!build) return { outcome: 'no-api' }

	const deck = path.join(deckDir, subject, probe.id + '.pptx')
	fs.mkdirSync(path.dirname(deck), { recursive: true })
	try {
		const pres = construct()
		await build(pres)
		await pres.writeFile({ fileName: deck })
	} catch (error) {
		return { outcome: 'error', message: error instanceof Error ? error.message : String(error) }
	}

	const unzip = await unzipSync()
	const entries = unzip(new Uint8Array(fs.readFileSync(deck)))
	const part = entries[probe.part]
	if (!part) return { outcome: 'absent', message: 'no ' + probe.part + ' in the package', deck }
	const xml = new TextDecoder('utf-8').decode(part)
	return { outcome: xml.includes(probe.construct) ? 'emitted' : 'absent', deck }
}

/**
 * One subject's entry from a per-subject map, or a throw naming the subject.
 *
 * `SUBJECTS` and these maps are built independently, so a subject added to the corpus with
 * no bundle path or no constructor behind it has to stop the run: silently measuring the
 * remaining subjects would produce a snapshot with a column quietly missing from it.
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
 * Does a library's shipped bundle contain the construct token anywhere?
 * @param {string[]} files
 * @param {string} token
 * @returns {boolean}
 */
function bundleMentions(files, token) {
	return files.some((file) => fs.readFileSync(file, 'utf8').includes(token))
}

/**
 * Run the whole corpus.
 * @param {object} [opts]
 * @param {string} [opts.workDir] - upstream install and built decks
 * @param {boolean} [opts.reuseUpstream]
 * @param {string | null} [opts.only] - a single probe id
 * @returns {Promise<{snapshot: any, decks: Record<string, Record<string, string>>, failures: string[]}>}
 */
export async function measure({ workDir = DEFAULT_WORK_DIR, reuseUpstream = false, only = null } = {}) {
	const probes = only ? PROBES.filter((probe) => probe.id === only) : PROBES
	if (probes.length === 0) throw new Error('no probe with id "' + only + '"')

	fs.mkdirSync(workDir, { recursive: true })
	const upstream = await installUpstream(workDir, reuseUpstream)
	const subjects = await loadSubjects(upstream.root)
	const bundles = shippedBundles(upstream.root)

	const deckDir = path.join(workDir, 'decks')
	fs.rmSync(deckDir, { recursive: true, force: true })

	/** @type {string[]} */
	const failures = []
	/** @type {Record<string, Record<string, string>>} */
	const decks = {}
	/** @type {any[]} */
	const coverage = []

	for (const probe of probes) {
		/** @type {Record<string, string>} */
		const results = {}
		/** @type {Record<string, string>} */
		const notes = {}
		for (const subject of SUBJECTS) {
			const { outcome, message, deck } = await measureProbe(probe, subject, forSubject(subjects, subject), deckDir)
			results[subject] = outcome
			if (message) notes[subject] = message
			if (deck) (decks[probe.id] ??= {})[subject] = path.relative(ROOT, deck)

			if (subject === 'ts-pptx' && (outcome === 'absent' || outcome === 'error'))
				failures.push(
					`${probe.id}: ts-pptx built a deck and ${outcome === 'error' ? 'threw' : 'emitted no'} ` +
						`${probe.construct}${message ? ' -- ' + message : ''}`
				)

			if (outcome === 'no-api') {
				const sighted = bundleMentions(forSubject(bundles, subject), probe.construct)
				const acknowledged = probe.sightings?.[subject]
				if (sighted && !acknowledged)
					failures.push(
						`${probe.id}: ${subject} is marked no-api but ${probe.construct} is in its shipped bundle; ` +
							'either the claim is wrong or the probe needs a `sightings` entry saying why'
					)
				if (!sighted && acknowledged)
					failures.push(
						`${probe.id}: ${subject} carries a \`sightings\` note but ${probe.construct} is not in its ` +
							'shipped bundle; the note has outlived what it acknowledged'
					)
				if (sighted && acknowledged) notes[subject] = acknowledged
			}
		}
		coverage.push({
			id: probe.id,
			label: probe.label,
			group: probe.group,
			construct: probe.construct,
			part: probe.part,
			results,
			...(Object.keys(notes).length > 0 ? { notes } : {}),
		})
	}

	const emitted = (/** @type {any} */ row, /** @type {string} */ subject) => row.results[subject] === 'emitted'
	const snapshot = {
		generatedAt: new Date().toISOString().slice(0, 10),
		subjects: {
			'ts-pptx': { version: readVersion(), source: 'workspace dist/' },
			pptxgenjs: { version: upstream.version, published: upstream.published, source: 'npm' },
		},
		coverage,
		// Recorded rather than left to be read off the table, because the whole point of these
		// two lists is that they are checked for emptiness. A page that never prints them cannot
		// be distinguished from one whose corpus was picked so they would come out empty.
		upstreamAhead: coverage.filter((row) => emitted(row, 'pptxgenjs') && !emitted(row, 'ts-pptx')).map((r) => r.id),
		sharedGaps: coverage.filter((row) => !SUBJECTS.some((s) => emitted(row, s))).map((r) => r.id),
	}
	return { snapshot, decks, failures }
}

/** This package's own version, for the snapshot header. */
function readVersion() {
	return JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version
}

/**
 * The coverage table, as a human reads it after a run.
 *
 * The two summary lines are printed only for a whole-corpus run. Over a `--probe` subset they
 * would still be true and still be useless: "upstream ahead: nothing" across one row invites
 * exactly the reading the lists exist to prevent.
 * @param {any} snapshot
 * @param {boolean} whole - was the whole corpus measured
 * @returns {void}
 */
function report(snapshot, whole) {
	const width = Math.max(...snapshot.coverage.map((/** @type {any} */ r) => r.label.length))
	let group = ''
	for (const row of snapshot.coverage) {
		if (row.group !== group) {
			group = row.group
			console.log('\n' + group)
		}
		const cells = SUBJECTS.map((subject) => `${subject} ${row.results[subject]}`.padEnd(24))
		console.log('  ' + row.label.padEnd(width) + '  ' + cells.join(' '))
	}
	console.log('')
	if (!whole) return
	console.log('upstream ahead: ' + (snapshot.upstreamAhead.join(', ') || 'nothing'))
	console.log('shared gaps:    ' + (snapshot.sharedGaps.join(', ') || 'nothing'))
}

async function main() {
	const { values } = parseCli(process.argv.slice(2), {
		usage: USAGE,
		options: {
			probe: { type: 'string' },
			'work-dir': { type: 'string' },
			'reuse-upstream': { type: 'boolean', default: false },
			out: { type: 'string' },
		},
	})

	const workDir = values['work-dir'] ? path.resolve(ROOT, values['work-dir']) : DEFAULT_WORK_DIR
	const { snapshot, decks, failures } = await measure({
		workDir,
		reuseUpstream: Boolean(values['reuse-upstream']),
		only: values.probe ?? null,
	})

	const built = Object.keys(decks).length
	report(snapshot, !values.probe)
	console.log('decks: ' + built + (built === 1 ? ' probe' : ' probes') + ' built under ' + path.relative(ROOT, workDir))

	if (failures.length > 0) {
		console.error('\n' + failures.length + ' failure(s):')
		for (const failure of failures) console.error('  ' + failure)
		return 1
	}

	if (values.probe) {
		console.log('\nsingle probe: snapshot not written')
		return 0
	}
	const out = values.out ? path.resolve(ROOT, values.out) : SNAPSHOT
	fs.writeFileSync(out, JSON.stringify(snapshot, null, '\t') + '\n')
	console.log('wrote ' + path.relative(ROOT, out))
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
