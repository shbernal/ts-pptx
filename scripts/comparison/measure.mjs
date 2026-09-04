#!/usr/bin/env node
/**
 * Measure both libraries by running them and reading what comes out.
 *
 * For every probe in `./probes.mjs`, this builds a deck with each library, opens the
 * package, and looks for the probe's construct in the probe's part. The result is one of
 * four outcomes per library, and `scripts/comparison/snapshot.json` is the committed
 * record of them. No row on the finished page comes from anywhere else.
 *
 * Three further families hang off that corpus, each in its own module and its own snapshot
 * key:
 *
 *   - **`validity`** (`./validity.mjs`) runs the decks just built through the same schema
 *     oracle `test:schema` uses.
 *   - **`hygiene`** (`./hygiene.mjs`) installs each library clean into a temp directory and
 *     weighs what a consumer gets.
 *   - **`health`** (`./health.mjs`) is deliberately apart from the other three: it measures
 *     *projects* rather than output, and belongs in its own section of the page for the
 *     same reason it has its own key here.
 *
 * A `--probe` run measures coverage only. The other three say nothing about one probe, and
 * two of them cost a pack, two installs and a clone to find that out.
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
 *   - **An unavailable measurement, unless `--allow-unavailable` says so.** Every fetch
 *     records why it failed instead of aborting, so a rate limit cannot block a release —
 *     and this is what stops a snapshot full of those holes being committed by reflex.
 *
 * The built decks stay on disk under the deck directory and their paths come back from
 * {@link measure}; the validity measurement reads exactly these rather than rebuilding a
 * second corpus that would drift from this one.
 */
import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { isMain, parseCli, ROOT, run, runCli } from '../script-utils.mjs'
import { measureHealth } from './health.mjs'
import { measureHygiene } from './hygiene.mjs'
import { PROBES, SUBJECTS } from './probes.mjs'
import { findUnavailable, isUnavailable, unavailable } from './unavailable.mjs'
import { measureValidity } from './validity.mjs'

const SNAPSHOT = path.join(ROOT, 'scripts', 'comparison', 'snapshot.json')
const DEFAULT_WORK_DIR = path.join(ROOT, '.tmp', 'comparison')

const USAGE = `Usage: node scripts/comparison/measure.mjs [options]

Builds every probe with both libraries, measures schema validity, package hygiene and
project health, and writes scripts/comparison/snapshot.json.

Options:
  --probe <id>          measure one probe's coverage only; does not write the snapshot
  --work-dir <dir>      where the installs, the clone and the built decks go
                        (default: .tmp/comparison)
  --reuse-installs      skip the installs and the clone, and use what is already in the
                        work dir
  --allow-unavailable   write the snapshot even when a measurement could not be taken
  --out <file>          snapshot path (default: scripts/comparison/snapshot.json)
  -h, --help            show this message`

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
 * @returns {Promise<{root: string, manifest: any, version: string, published: string | import('./unavailable.mjs').Unavailable}>}
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
	const manifestPath = path.join(root, 'package.json')
	if (!fs.existsSync(manifestPath)) throw new Error('pptxgenjs is not installed at ' + manifestPath)
	const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
	return { root, manifest, version: manifest.version, published: await publishDate(manifest.version) }
}

/**
 * When the registry says the measured version was published.
 *
 * Not the same fact as `health.pptxgenjs.npm.lastPublish`, which is the *latest* version's
 * date: this one dates the artifact every coverage row was read off, so a reader can tell
 * how old the measurement is without trusting the file it is written in.
 *
 * An {@link unavailable} marker rather than a throw when the registry cannot be reached,
 * on the same rule the rest of the snapshot follows.
 * @param {string} version
 * @returns {Promise<string | import('./unavailable.mjs').Unavailable>}
 */
async function publishDate(version) {
	try {
		const { stdout } = await run('npm', ['view', 'pptxgenjs@' + version, 'time', '--json'], { capture: true })
		const times = JSON.parse(stdout)
		const stamp = typeof times === 'string' ? times : times?.[version]
		if (typeof stamp === 'string') return stamp.slice(0, 10)
		return unavailable('npm view returned no publish time for pptxgenjs@' + version)
	} catch (error) {
		return unavailable('npm view failed: ' + (error instanceof Error ? error.message.split('\n')[0] : String(error)))
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
 * @param {string} [opts.workDir] - installs, clone and built decks
 * @param {boolean} [opts.reuseInstalls]
 * @param {string | null} [opts.only] - a single probe id; measures coverage alone
 * @returns {Promise<{snapshot: any, decks: Record<string, Record<string, string>>, failures: string[]}>}
 */
export async function measure({ workDir = DEFAULT_WORK_DIR, reuseInstalls = false, only = null } = {}) {
	const probes = only ? PROBES.filter((probe) => probe.id === only) : PROBES
	if (probes.length === 0) throw new Error('no probe with id "' + only + '"')

	fs.mkdirSync(workDir, { recursive: true })
	const upstream = await installUpstream(workDir, reuseInstalls)
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
		...(only
			? {}
			: {
					validity: await measureValidity(coverage, decks),
					hygiene: await measureHygiene({ workDir, upstreamRoot: upstream.root, reuse: reuseInstalls }),
					health: await measureHealth({
						workDir,
						upstreamManifest: upstream.manifest,
						reuse: reuseInstalls,
					}),
				}),
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
	reportFamilies(snapshot)
}

/** One value as a console cell: a number, a string, or why it is missing. */
function cell(/** @type {any} */ value) {
	if (isUnavailable(value)) return 'unavailable'
	if (typeof value === 'number') return value.toLocaleString('en-US')
	return String(value)
}

/** @param {number} bytes */
const kb = (bytes) => (bytes / 1024).toFixed(0) + ' kB'

/**
 * The three families that are not the coverage table, one line each.
 *
 * A digest, not the page: this is what a human reads to decide whether the run is worth
 * committing, and the rendered page in `docs/` is where the numbers get their framing.
 * @param {any} snapshot
 * @returns {void}
 */
function reportFamilies(snapshot) {
	for (const subject of SUBJECTS) {
		const validity = snapshot.validity?.[subject]
		const hygiene = snapshot.hygiene?.[subject]
		const health = snapshot.health?.[subject]
		if (!validity && !hygiene && !health) continue
		console.log('\n' + subject)
		if (validity)
			console.log(
				'  validity   ' + `${validity.cleanDecks ?? 0}/${validity.decks} decks clean, ${validity.errors ?? 0} error(s)`
			)
		if (hygiene)
			console.log(
				'  hygiene    ' +
					`${kb(hygiene.install.bytes)} installed, ${hygiene.dependencies.transitive} transitive dep(s), ` +
					`${hygiene.entryPoints.length} entry point(s), hello-world ` +
					(isUnavailable(hygiene.helloWorld)
						? 'unavailable'
						: `${kb(hygiene.helloWorld.initialBytes)} initial / ${kb(hygiene.helloWorld.totalBytes)} total`)
			)
		if (health) {
			console.log(
				'  health     ' +
					`${cell(health.stars)} stars, ${cell(health.npm.downloadsLastMonth)} downloads/month, ` +
					`last ${cell(health.defaultBranch)} commit ${cell(health.lastDefaultBranchCommit)}, ` +
					`last publish ${cell(health.npm.lastPublish)}`
			)
			const coverage = health.source.statementCoverage
			console.log(
				'  source     ' +
					`${cell(health.source.lines)} src lines, ${cell(health.source.testLines)} test lines, ` +
					'statements ' +
					(typeof coverage?.pct === 'number' ? `${coverage.pct}% (${coverage.lane} lane)` : cell(coverage))
			)
		}
	}
}

async function main() {
	const { values } = parseCli(process.argv.slice(2), {
		usage: USAGE,
		options: {
			probe: { type: 'string' },
			'work-dir': { type: 'string' },
			'reuse-installs': { type: 'boolean', default: false },
			'allow-unavailable': { type: 'boolean', default: false },
			out: { type: 'string' },
		},
	})

	const workDir = values['work-dir'] ? path.resolve(ROOT, values['work-dir']) : DEFAULT_WORK_DIR
	const { snapshot, decks, failures } = await measure({
		workDir,
		reuseInstalls: Boolean(values['reuse-installs']),
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
		console.log('\nsingle probe: coverage only, snapshot not written')
		return 0
	}

	const holes = findUnavailable(snapshot)
	if (holes.length > 0) {
		const stream = values['allow-unavailable'] ? console.log : console.error
		stream('\n' + holes.length + ' measurement(s) unavailable:')
		for (const hole of holes) stream('  ' + hole.path + ': ' + hole.reason)
		if (!values['allow-unavailable']) {
			console.error('\nsnapshot not written. Pass --allow-unavailable to commit it with these holes in it.')
			return 1
		}
	}

	const out = values.out ? path.resolve(ROOT, values.out) : SNAPSHOT
	fs.writeFileSync(out, JSON.stringify(snapshot, null, '\t') + '\n')
	console.log('wrote ' + path.relative(ROOT, out))
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
