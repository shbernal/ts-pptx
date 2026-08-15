#!/usr/bin/env node
/**
 * Coverage merge — one report, two collectors.
 *
 * The Node suite runs `dist/node.js` under Vitest and reports what it executed. The
 * browser lane runs `dist/browser.js` unbundled in Chromium and, until this script
 * existed, reported nothing into that number at all. So the repo's headline coverage
 * counted `src/runtime/browser.ts` in its denominator (correctly — dropping the exclusion
 * stopped hiding it) while counting almost none of it in its numerator (also correctly —
 * Node cannot execute an adapter that needs `fetch`, `FileReader` and a canvas). An
 * honest denominator with a blind collector is a number that understates the truth, and
 * the standing fix for it was written into vitest.config.ts before it was built: merge
 * the lane's coverage in.
 *
 *   node scripts/coverage-merge.mjs        # merge, write coverage/merged/, print the delta
 *
 * Inputs, both of which must already exist:
 *
 *   - `coverage/coverage-final.json` — `pnpm run test:coverage`
 *   - `.tmp/browser-coverage/*.json` — `pnpm run test:browser` (test/browser/fixtures.mjs)
 *
 * ## Why this can be merged at all
 *
 * Two things make it sound rather than a plausible-looking average of two different
 * measurements:
 *
 * 1. **The same remapper.** Vitest 4's v8 provider converts V8 coverage to Istanbul with
 *    `ast-v8-to-istanbul`, and so does this script — same package, same version, pinned as
 *    a devDependency for exactly that reason. Feed it the same original sources and it
 *    produces the same statement, function and branch structures. A merge across two
 *    *different* remappers would union two disagreeing views of one file and inflate the
 *    denominator, which is the one failure mode worse than the number being too low.
 * 2. **Location-keyed merging.** `istanbul-lib-coverage` merges file coverage by source
 *    location, not by array index, so nothing depends on the two sides having discovered
 *    the same nodes in the same order.
 *
 * (1) holds only approximately, and the gap is real: the Node side's input is the
 * SSR-transformed `dist/node.js` Vite handed Vitest, this side's is the raw `dist/`
 * chunk the browser fetched — a *different bundle of the same source*. Both remap to
 * `src/**` coordinates, but a handful of mappings resolve one end of a span to a
 * different column depending on which bundle they came through. Measured at 137 of
 * 13,248 locations, ~1%.
 *
 * ## What that means for the merge, and the rule that follows
 *
 * A location the two sides spell differently would not merge — it would be added
 * *alongside* its twin, so one statement would occupy two slots and the denominator would
 * grow by the amount the two bundlers happened to disagree. A coverage report must never
 * do that, so this script does not let it:
 *
 *   **The Node report defines the shape. The browser lane contributes hits.**
 *
 * Every browser file coverage is projected onto the Node report's own statement, function
 * and branch maps, matched by source location. Consequences, all of them intended:
 *
 *   - the merged denominator is *identical* to the Node report's, so the merged
 *     percentage is directly comparable to the one it replaces rather than being a
 *     different measurement wearing the same name;
 *   - the merged numerator can only rise, never fall;
 *   - a browser hit on one of those ~1% of jittered locations is dropped rather than
 *     invented somewhere else. The count is printed on every run. It errs downward, which
 *     is the direction a coverage number is allowed to err.
 *
 * The share is also a gate: past `MAX_ORPHAN_SHARE` the two lanes are not looking at the
 * same build at all — a stale `dist/`, a half-rebuilt chunk — and the run fails instead of
 * reporting a number assembled from two different packages.
 */

import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import astV8ToIstanbul from 'ast-v8-to-istanbul'
import libCoverage from 'istanbul-lib-coverage'
import libReport from 'istanbul-lib-report'
import reports from 'istanbul-reports'
import { parseAstAsync } from 'vitest/node'
import { ROOT } from './script-utils.mjs'

const NODE_REPORT = path.join(ROOT, 'coverage', 'coverage-final.json')
const BROWSER_DIR = path.join(ROOT, '.tmp', 'browser-coverage')
const OUT_DIR = path.join(ROOT, 'coverage', 'merged')

/** @type {(keyof import('istanbul-lib-coverage').CoverageSummaryData)[]} */
const AXES = ['statements', 'branches', 'functions', 'lines']

/**
 * @param {string} message
 * @returns {never}
 */
function fail(message) {
	console.error('coverage-merge: ' + message)
	process.exit(1)
}

/** @param {string} file */
function relative(file) {
	return path.relative(ROOT, file).replace(/\\/g, '/')
}

// --- inputs ---

function readNodeMap() {
	if (!fs.existsSync(NODE_REPORT)) {
		fail(`no Node report at ${relative(NODE_REPORT)}.\n  Run: pnpm run test:coverage`)
	}
	return libCoverage.createCoverageMap(JSON.parse(fs.readFileSync(NODE_REPORT, 'utf8')))
}

function readBrowserRecords() {
	if (!fs.existsSync(BROWSER_DIR)) {
		fail(`no browser coverage at ${relative(BROWSER_DIR)}/.\n  Run: pnpm run test:browser`)
	}
	const files = fs.readdirSync(BROWSER_DIR).filter((name) => name.endsWith('.json'))
	if (!files.length) {
		fail(`${relative(BROWSER_DIR)}/ is empty.\n  Run: pnpm run test:browser`)
	}
	return files.map((name) => JSON.parse(fs.readFileSync(path.join(BROWSER_DIR, name), 'utf8')))
}

// --- V8 -> Istanbul, the way @vitest/coverage-v8 does it ---

/** Parsed once per `dist/` file and reused across every record that touched it. */
const sources = new Map()

/** @param {string} distPath a server path like `/dist/browser.js` */
async function sourcesFor(distPath) {
	const cached = sources.get(distPath)
	if (cached) return cached

	const file = path.join(ROOT, distPath.replace(/^\//, ''))
	if (!fs.existsSync(file))
		fail(`${distPath} was served to the browser but is not in dist/. Rebuild and re-run both lanes.`)

	const code = fs.readFileSync(file, 'utf8')
	const url = pathToFileURL(file).href

	// Sourcemap `sources` are relative to the map; resolving them to absolute `file://`
	// URLs is what makes `ast-v8-to-istanbul` key its output on the same absolute `src/**`
	// paths the Node report uses. This mirrors `getSources` in @vitest/coverage-v8.
	const mapFile = file + '.map'
	if (!fs.existsSync(mapFile)) fail(`${distPath} has no sourcemap; coverage cannot be remapped to src/.`)
	const map = JSON.parse(fs.readFileSync(mapFile, 'utf8'))
	map.sources = (map.sources ?? [])
		.filter(Boolean)
		.map(/** @param {string} source */ (source) => new URL(source, url).href)

	const entry = { code, url, map, ast: await parseAstAsync(code) }
	sources.set(distPath, entry)
	return entry
}

/**
 * @typedef {import('node:inspector').Profiler.FunctionCoverage} FunctionCoverage
 * @param {{entries: {path: string, sourceLength: number|null, functions: FunctionCoverage[]}[]}[]} records
 */
async function convertBrowserCoverage(records) {
	const map = libCoverage.createCoverageMap({})
	for (const record of records) {
		for (const entry of record.entries) {
			const { code, url, map: sourceMap, ast } = await sourcesFor(entry.path)

			// V8's ranges are byte offsets into what the browser parsed. If that is not the
			// file on disk, every offset below is being applied to the wrong text, and the
			// result would be wrong rather than absent — so this is a hard stop.
			if (entry.sourceLength != null && entry.sourceLength !== code.length) {
				fail(
					`${entry.path} is ${code.length} chars on disk but was ${entry.sourceLength} in the browser.\n` +
						`  dist/ changed between the two lanes; rebuild and re-run both.`
				)
			}

			map.merge(
				await astV8ToIstanbul({
					code,
					sourceMap,
					ast,
					coverage: { functions: entry.functions, url },
					// Browser scripts are not wrapped the way node:vm wraps a module.
					wrapperLength: 0,
				})
			)
		}
	}

	// The Node report drops files that no longer exist on disk; match it, and drop the
	// bundled dependencies' own sources while we are here — they are not in the Node
	// report's file set, so adding them would grow the denominator with code this repo
	// does not own.
	map.filter(/** @param {string} file */ (file) => fs.existsSync(file) && !file.includes('node_modules'))

	// Round-trip through JSON, which is not a no-op and is not cosmetic. Istanbul writes
	// an open-ended location's end column as `Infinity` in memory, and `JSON.stringify`
	// turns that into `null` — so the Node side, which this script reads back from
	// `coverage-final.json`, spells those locations differently from this side, which never
	// left memory. `istanbul-lib-coverage` merges *by* location key, so the two spellings
	// would not merge: every such statement would be counted twice, once covered and once
	// not, quietly inflating the denominator. Sending this side through the same
	// serialization the other side already went through makes them the same data.
	return libCoverage.createCoverageMap(JSON.parse(JSON.stringify(map.toJSON())))
}

// --- projecting the browser lane onto the Node report's shape ---

/**
 * Past this share of browser locations having no slot in the Node report, the two lanes
 * are not describing the same build and nothing they say together is worth reporting.
 * Measured ~1% on a matched pair, which is the bundler jitter documented in the header;
 * a stale `dist/` on one side puts it in a different league entirely.
 */
const MAX_ORPHAN_SHARE = 5

/**
 * Both sides have been through JSON by the time this runs, so an open-ended location's end
 * column is `null` on both — but the normalization is kept anyway, because a comparison
 * that reports a false disagreement is worse than useless: it would condemn a merge that
 * is in fact exact. See the round-trip in `convertBrowserCoverage` for where `Infinity`
 * becomes `null` and why it has to.
 */
/** @param {import('istanbul-lib-coverage').Location} p */
const pos = (p) => `${p.line}:${Number.isFinite(p.column) ? p.column : -1}`
/** @param {import('istanbul-lib-coverage').Range} loc */
const locKey = (loc) => `${pos(loc.start)}-${pos(loc.end)}`

/**
 * Every hit the browser lane recorded for one file, keyed by source location.
 *
 * Counts are summed rather than replaced: the same location can appear more than once
 * across the lane's eight scenarios, and what matters downstream is only whether the
 * total is zero.
 * @param {import('istanbul-lib-coverage').FileCoverageData} data
 * @returns {Map<string, number>}
 */
function hitsByLocation(data) {
	/** @type {Map<string, number>} */
	const hits = new Map()
	/**
	 * @param {string} key
	 * @param {number} count
	 */
	const add = (key, count) => hits.set(key, (hits.get(key) ?? 0) + count)

	for (const [index, loc] of Object.entries(data.statementMap)) add('s:' + locKey(loc), data.s[index] ?? 0)
	for (const [index, fn] of Object.entries(data.fnMap)) add('f:' + locKey(fn.loc), data.f[index] ?? 0)
	for (const [index, branch] of Object.entries(data.branchMap)) {
		branch.locations.forEach(
			/**
			 * @param {import('istanbul-lib-coverage').Range} loc
			 * @param {number} arm
			 */
			(loc, arm) => add('b:' + locKey(loc), data.b[index]?.[arm] ?? 0)
		)
	}
	return hits
}

/**
 * Re-express the browser lane's coverage of one file in the Node report's own maps.
 *
 * The returned file coverage has the Node report's exact `statementMap`, `fnMap` and
 * `branchMap` — so merging it can only move hit counts, never the denominator — with each
 * count taken from whatever the browser recorded at that source location, or zero if it
 * recorded nothing there.
 * @param {import('istanbul-lib-coverage').FileCoverageData} nodeData
 * @param {import('istanbul-lib-coverage').FileCoverageData} browserData
 */
function project(nodeData, browserData) {
	const hits = hitsByLocation(browserData)
	/** @type {Set<string>} */
	const used = new Set()
	/**
	 * @param {string} key
	 * @returns {number}
	 */
	const take = (key) => {
		const hit = hits.get(key)
		if (hit === undefined) return 0
		used.add(key)
		return hit
	}

	/** @type {Record<string, number>} */
	const s = {}
	for (const [index, loc] of Object.entries(nodeData.statementMap)) s[index] = take('s:' + locKey(loc))

	/** @type {Record<string, number>} */
	const f = {}
	for (const [index, fn] of Object.entries(nodeData.fnMap)) f[index] = take('f:' + locKey(fn.loc))

	/** @type {Record<string, number[]>} */
	const b = {}
	for (const [index, branch] of Object.entries(nodeData.branchMap)) {
		b[index] = branch.locations.map(
			/** @param {import('istanbul-lib-coverage').Range} loc */ (loc) => take('b:' + locKey(loc))
		)
	}

	return {
		coverage: {
			path: nodeData.path,
			statementMap: nodeData.statementMap,
			fnMap: nodeData.fnMap,
			branchMap: nodeData.branchMap,
			s,
			f,
			b,
		},
		// Locations the browser lane measured that the Node report has no slot for. Their
		// hits are dropped; see the header for why that is the safe direction.
		orphans: hits.size - used.size,
		measured: hits.size,
	}
}

// --- reporting ---

/** @param {import('istanbul-lib-coverage').CoverageMap} map */
function summarize(map) {
	const summary = libCoverage.createCoverageSummary()
	for (const file of map.files()) summary.merge(map.fileCoverageFor(file).toSummary())
	return summary
}

/**
 * @param {import('istanbul-lib-coverage').CoverageSummaryData} before
 * @param {import('istanbul-lib-coverage').CoverageSummaryData} after
 */
function printDelta(before, after) {
	console.log('')
	console.log('  merged coverage (Node suite + browser lane)')
	console.log('')
	console.log('  axis         node only     merged      delta')
	for (const axis of AXES) {
		const from = before[axis].pct
		const to = after[axis].pct
		const delta = to - from
		console.log(
			'  ' +
				axis.padEnd(12) +
				from.toFixed(2).padStart(9) +
				to.toFixed(2).padStart(11) +
				(delta >= 0 ? '+' : '') +
				delta.toFixed(2).padStart(10)
		)
	}
	console.log('')
}

// --- main ---

const nodeMap = readNodeMap()
const records = readBrowserRecords()
const browserMap = await convertBrowserCoverage(records)

// A deep copy, not `merged.merge(nodeMap)`. Istanbul's `FileCoverage` wraps the data
// object it is handed rather than copying it, so a merged map seeded that way shares
// every file's data with `nodeMap` — and merging into it silently rewrites the very
// report being compared against. The "node only" column would then read back as the
// merged one and every delta would print 0.00.
const merged = libCoverage.createCoverageMap(JSON.parse(JSON.stringify(nodeMap.toJSON())))

let orphans = 0
let measured = 0
const unknownFiles = []
for (const file of browserMap.files()) {
	const nodeData = nodeMap.data[file]
	if (!nodeData) {
		// The Node report includes every file reachable from `dist/**`, covered or not, so
		// there should be nothing here. If there is, it is a finding about the Node report's
		// `include` rather than a file to quietly add on different terms from the rest.
		unknownFiles.push(file)
		continue
	}
	const projected = project(nodeMap.fileCoverageFor(file).data, browserMap.fileCoverageFor(file).data)
	orphans += projected.orphans
	measured += projected.measured
	merged.merge({ [file]: projected.coverage })
}

const orphanShare = measured ? (orphans / measured) * 100 : 0
if (orphanShare > MAX_ORPHAN_SHARE) {
	fail(
		`${orphans} of ${measured} browser locations (${orphanShare.toFixed(1)}%) have no slot in the Node report.\n` +
			`  That is past the ${MAX_ORPHAN_SHARE}% bundler-jitter budget, which means the two lanes ran against\n` +
			`  different builds. Rebuild dist/ and re-run both lanes:\n` +
			`    pnpm run test:coverage && pnpm run test:browser`
	)
}

fs.rmSync(OUT_DIR, { recursive: true, force: true })
const context = libReport.createContext({ dir: OUT_DIR, coverageMap: merged })
for (const reporter of /** @type {(keyof import('istanbul-reports').ReportOptions)[]} */ ([
	'json',
	'json-summary',
	'html',
])) {
	reports.create(reporter, { projectRoot: ROOT }).execute(context)
}

printDelta(summarize(nodeMap).data, summarize(merged).data)
console.log(
	`  ${records.length} browser record(s), ${sources.size} dist script(s), ` +
		`${browserMap.files().length} source file(s) reached`
)
console.log(
	`  ${orphans} of ${measured} browser locations (${orphanShare.toFixed(2)}%) had no slot in the Node ` +
		`report's shape and were dropped`
)
for (const file of unknownFiles) {
	console.log(`  NOTE: ${relative(file)} is covered by the browser lane but absent from the Node report; not merged`)
}
console.log(`  report: ${relative(OUT_DIR)}/`)
console.log('')
