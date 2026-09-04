/**
 * Project health: activity, adoption and the size of each source tree.
 *
 * Kept in its own snapshot key, and it belongs in its own page section, because it
 * measures *projects* rather than output. Coverage, validity and hygiene are all readings
 * of bytes a library emitted; nothing below is. Mixing them would let a strong row in one
 * family stand in for a weak one in another, which is the failure mode a comparison
 * published by one of the two libraries has to work hardest to avoid.
 *
 * ## Two traps, both of them load-bearing
 *
 * **`pushed_at` is not activity on the default branch.** GitHub sets it from a push to
 * *any* ref, so upstream's `2025-11-28` sits against a `master` head of `2025-06-25`:
 * something landed on a side branch. This records the default branch's head commit and the
 * last npm publish, and never `pushed_at`. The supportable sentence is "no release or
 * default-branch commit since June 2025"; the word "abandoned" is not, and does not appear
 * here or on the page.
 *
 * **The adoption gap gets printed.** Roughly eleven million downloads a month against two
 * thousand. A comparison that hides its worst row is not read as balanced, it is read as
 * dishonest, and it takes the credible rows down with it. Ours is the sum over both names
 * it ships under, and the per-name split is recorded too — most of that traffic is still
 * arriving at the scoped alias, and a single total would quietly imply otherwise.
 *
 * ## Upstream's test story is re-derived, not inherited
 *
 * A shallow clone at measure time, then: does the manifest declare a test script, are there
 * spec files, is there a test directory. Today the answer is none of the three and
 * `TESTING.md` describes a manual checklist, so coverage records as `no automated suite` —
 * a fact with a date on it, not a permanent one. If a suite lands, this notices without
 * anyone editing a string.
 */
import fs from 'node:fs'
import path from 'node:path'
import { ALIAS_NAME } from '../alias-package.mjs'
import { ROOT, run } from '../script-utils.mjs'
import { isUnavailable, unavailable } from './unavailable.mjs'

/** The merged coverage report — the number `coverage:gate` holds this repo to. */
const MERGED_COVERAGE = path.join(ROOT, 'coverage', 'merged', 'coverage-summary.json')

/** The Node lane alone, which `test:coverage` leaves behind. See {@link ownCoverage}. */
const NODE_COVERAGE = path.join(ROOT, 'coverage', 'coverage-summary.json')

/** Source extensions that count as lines of code, on both sides. */
const CODE_EXTENSIONS = ['.ts', '.mts', '.js', '.mjs', '.cjs']

/**
 * Directory names a source count never descends into.
 *
 * Build output and scratch space, not source. `.tmp` matters most: this repo's holds the
 * comparison work directory, which by the time this runs contains two whole installs and a
 * clone of the library being compared against.
 */
const SKIP_DIRS = new Set(['node_modules', '.git', '.tmp', 'dist', 'coverage'])

/**
 * A JSON endpoint, or an {@link unavailable} marker carrying why it was not reached.
 *
 * Never throws. A rate limit, a proxy or an offline afternoon must not be able to block a
 * release, and every caller here treats a hole as data — see `--allow-unavailable`, which
 * is what stops a snapshot full of them being committed by reflex.
 *
 * `GITHUB_TOKEN` is used when the environment offers one, purely for the rate limit. The
 * endpoints below are all public and none of them needs it.
 * @param {string} url
 * @returns {Promise<any>}
 */
async function fetchJson(url) {
	/** @type {Record<string, string>} */
	const headers = { accept: 'application/json', 'user-agent': 'ts-pptx-comparison' }
	const token = process.env.GITHUB_TOKEN
	if (token && url.startsWith('https://api.github.com/')) headers.authorization = 'Bearer ' + token
	try {
		const response = await fetch(url, { headers })
		if (!response.ok) return unavailable(url + ' returned HTTP ' + response.status)
		return await response.json()
	} catch (error) {
		return unavailable(url + ' could not be reached: ' + (error instanceof Error ? error.message : String(error)))
	}
}

/**
 * A failed fetch's own reason, or a new one — never the response body.
 *
 * {@link fetchJson} answers with either parsed JSON or a marker, and a caller that reaches
 * into that JSON for a field can find it missing. Passing the object straight through as
 * the field's value would put a whole GitHub repository object in the snapshot where a star
 * count belongs, so anything that is not already a marker becomes one.
 * @param {unknown} value - what the fetch returned
 * @param {string} reason - why the field is missing, when the fetch itself succeeded
 * @returns {import('./unavailable.mjs').Unavailable}
 */
function orUnavailable(value, reason) {
	return isUnavailable(value) ? value : unavailable(reason)
}

/**
 * `owner/repo` from a manifest's `repository` field, or `null` when it does not name a
 * GitHub repository.
 * @param {Record<string, any>} manifest
 * @returns {string | null}
 */
export function githubSlug(manifest) {
	const url = typeof manifest.repository === 'string' ? manifest.repository : manifest.repository?.url
	if (typeof url !== 'string') return null
	const match = /github\.com[/:]([^/]+)\/([^/]+?)(?:\.git)?$/.exec(url)
	return match ? match[1] + '/' + match[2] : null
}

/**
 * How many open issues, and how many open pull requests.
 *
 * Two search queries rather than the repository object's `open_issues_count`, which counts
 * pull requests as issues. That field is the one most likely to be quoted at this table, so
 * being wrong in the direction it is wrong would be the expensive mistake.
 * @param {string} slug
 * @returns {Promise<{openIssues: number | import('./unavailable.mjs').Unavailable, openPullRequests: number | import('./unavailable.mjs').Unavailable}>}
 */
async function openWork(slug) {
	/** @param {string} type @returns {Promise<any>} */
	const count = async (type) => {
		const query = encodeURIComponent('repo:' + slug + ' type:' + type + ' state:open')
		const result = await fetchJson('https://api.github.com/search/issues?per_page=1&q=' + query)
		return typeof result?.total_count === 'number'
			? result.total_count
			: orUnavailable(result, 'the search result carries no total')
	}
	return { openIssues: await count('issue'), openPullRequests: await count('pr') }
}

/**
 * Activity and attention on one GitHub repository.
 * @param {string | null} slug
 * @returns {Promise<Record<string, unknown>>}
 */
async function githubHealth(slug) {
	if (slug === null) return { repo: unavailable('the manifest names no GitHub repository') }

	const repo = await fetchJson('https://api.github.com/repos/' + slug)
	const branch = typeof repo?.default_branch === 'string' ? repo.default_branch : null
	const head = branch ? await fetchJson('https://api.github.com/repos/' + slug + '/commits/' + branch) : repo
	const committed = head?.commit?.committer?.date

	// A failure here cascades: no repository object means no default branch, which means no
	// head commit to ask for. Each field carries the *original* reason rather than a
	// paraphrase of the field above it, so a run reports one cause four times instead of
	// three consequences and a cause.
	return {
		repo: slug,
		defaultBranch: branch ?? orUnavailable(repo, 'the repository object carries no default branch'),
		// The head commit of the default branch, deliberately not `pushed_at` — see the header.
		lastDefaultBranchCommit:
			typeof committed === 'string'
				? committed.slice(0, 10)
				: orUnavailable(head, 'the head commit carries no committer date'),
		stars:
			typeof repo?.stargazers_count === 'number'
				? repo.stargazers_count
				: orUnavailable(repo, 'the repository object carries no star count'),
		...(await openWork(slug)),
	}
}

/**
 * Publish date and monthly downloads, summed over every name a library ships under.
 *
 * The per-name split stays in the record. Ours is spread over the canonical name and the
 * scoped alias, and a bare total would read as one package's adoption when it is two names
 * for the same bytes — see `scripts/alias-package.mjs` for why the alias exists.
 * @param {string[]} names - npm package names, canonical first
 * @returns {Promise<Record<string, unknown>>}
 */
async function npmHealth(names) {
	/** @type {Record<string, unknown>} */
	const perName = {}
	let downloads = 0
	let downloadsKnown = false
	/** @type {string | null} */
	let lastPublish = null

	for (const name of names) {
		const point = await fetchJson('https://api.npmjs.org/downloads/point/last-month/' + name)
		if (typeof point?.downloads === 'number') {
			perName[name] = point.downloads
			downloads += point.downloads
			downloadsKnown = true
		} else perName[name] = orUnavailable(point, 'the downloads endpoint returned no count for ' + name)

		const registry = await fetchJson('https://registry.npmjs.org/' + encodeURIComponent(name))
		const latest = registry?.['dist-tags']?.latest
		const stamp = typeof latest === 'string' ? registry?.time?.[latest] : undefined
		// The newest publish across the names, since the alias is published after the
		// canonical package and either could be the more recent one.
		if (typeof stamp === 'string' && (lastPublish === null || stamp > lastPublish)) lastPublish = stamp.slice(0, 10)
	}

	return {
		names,
		downloadsLastMonth: downloadsKnown ? downloads : unavailable('no download figure could be fetched'),
		downloadsByName: perName,
		lastPublish: lastPublish ?? unavailable('no publish date could be fetched'),
	}
}

/**
 * Lines in one file, counting a trailing newline as a terminator rather than a line.
 * @param {string} file
 * @returns {number}
 */
function lineCount(file) {
	const text = fs.readFileSync(file, 'utf8')
	if (text.length === 0) return 0
	return text.split('\n').length - (text.endsWith('\n') ? 1 : 0)
}

/**
 * Lines in every source file under a directory.
 *
 * A line count, not a statement count: the two libraries are formatted to different rules
 * and no normalisation makes them the same measurement, so this is reported as what it is —
 * an order-of-magnitude figure for how much code there is to maintain.
 * @param {string} dir
 * @returns {number} 0 when the directory does not exist
 */
export function sourceLines(dir) {
	if (!fs.existsSync(dir)) return 0
	let lines = 0
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		if (entry.isDirectory()) {
			if (!SKIP_DIRS.has(entry.name)) lines += sourceLines(path.join(dir, entry.name))
			continue
		}
		if (!entry.isFile() || !CODE_EXTENSIONS.includes(path.extname(entry.name))) continue
		lines += lineCount(path.join(dir, entry.name))
	}
	return lines
}

/**
 * Statement coverage, measured — never the gate's notch.
 *
 * Not from `scripts/coverage-gates.json`: that file holds the thresholds, and a threshold is
 * a floor somebody chose, not a measurement. Publishing it would be publishing a number that
 * is true by construction.
 *
 * Two real sources, and which one was read is recorded beside the number:
 *
 *   - **`merged`**, from `coverage/merged/`, is the honest one — every line it counts has a
 *     collector, because the browser lane is folded in. It is what `coverage:gate` holds
 *     this repo to.
 *   - **`node`**, from `coverage/coverage-summary.json`, is the fallback, and it is a
 *     *lower* number: the Node lane's denominator includes browser-only code it structurally
 *     cannot execute. Falling back therefore errs against us, which is the only direction a
 *     fallback in a comparison we publish about ourselves may err in.
 *
 * The fallback exists because the merged report is a CI artifact — the two lanes run in
 * separate jobs and are stitched together in a third — so demanding it here would mean the
 * snapshot could only be regenerated on a machine that had just run both. That is a
 * coupling, not a standard.
 * @returns {{pct: number, lane: string} | import('./unavailable.mjs').Unavailable}
 */
function ownCoverage() {
	/** @type {Array<[string, string]>} */
	const sources = [
		['merged', MERGED_COVERAGE],
		['node', NODE_COVERAGE],
	]
	for (const [lane, file] of sources) {
		if (!fs.existsSync(file)) continue
		const pct = JSON.parse(fs.readFileSync(file, 'utf8'))?.total?.statements?.pct
		if (typeof pct === 'number') return { pct, lane }
	}
	return unavailable('no coverage report on disk; run `pnpm run coverage:gate` or `pnpm run test:coverage` first')
}

/** Directory names that make everything under them test code. */
const TEST_DIRS = ['test', 'tests', '__tests__']

/**
 * Whether a source tree has an automated test suite, what says so, and how big it is.
 *
 * Three independent signals, all recorded: a `test`-prefixed script in the manifest, files
 * named `*.test.*` or `*.spec.*`, and a directory called `test`, `tests` or `__tests__`.
 * Recording the evidence rather than a verdict is what lets the answer change on its own
 * when one of the three appears — nobody has to remember to edit a string.
 *
 * `testLines` comes out of the same walk rather than from a second pass over `testDirs`,
 * because the two conventions are not nested: a project can hold its specs beside the code
 * they cover with no test directory at all, and summing directories would report zero test
 * lines for a project that is entirely tested. Counted once per file, so a spec inside a
 * test directory is not counted twice.
 * @param {string} tree - repository root
 * @returns {{testScripts: string[], specFiles: number, testDirs: string[], testLines: number}}
 */
export function testEvidence(tree) {
	const manifestPath = path.join(tree, 'package.json')
	const scripts = fs.existsSync(manifestPath) ? (JSON.parse(fs.readFileSync(manifestPath, 'utf8')).scripts ?? {}) : {}
	const testScripts = Object.keys(scripts).filter((name) => name === 'test' || name.startsWith('test:'))

	let specFiles = 0
	let testLines = 0
	/** @type {string[]} */
	const testDirs = []
	/** @param {string} dir @param {string} rel @param {boolean} underTestDir */
	const walk = (dir, rel, underTestDir) => {
		for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
			const here = rel ? rel + '/' + entry.name : entry.name
			if (entry.isDirectory()) {
				if (SKIP_DIRS.has(entry.name)) continue
				const isTestDir = underTestDir || TEST_DIRS.includes(entry.name)
				if (isTestDir && !underTestDir) testDirs.push(here)
				walk(path.join(dir, entry.name), here, isTestDir)
				continue
			}
			if (!entry.isFile()) continue
			const isSpec = /\.(test|spec)\.[cm]?[jt]sx?$/.test(entry.name)
			if (isSpec) specFiles += 1
			if ((isSpec || underTestDir) && CODE_EXTENSIONS.includes(path.extname(entry.name)))
				testLines += lineCount(path.join(dir, entry.name))
		}
	}
	walk(tree, '', false)
	return { testScripts, specFiles, testDirs, testLines }
}

/**
 * A shallow clone of one repository, or an {@link unavailable} marker.
 *
 * Shallow because only the working tree is wanted; the history is measured through the API
 * above, where one request answers what a full clone would cost megabytes to.
 * @param {string} slug
 * @param {string} into
 * @param {boolean} reuse
 * @returns {Promise<string | import('./unavailable.mjs').Unavailable>}
 */
async function shallowClone(slug, into, reuse) {
	if (reuse && fs.existsSync(path.join(into, 'package.json'))) return into
	fs.rmSync(into, { recursive: true, force: true })
	try {
		await run('git', ['clone', '--depth', '1', 'https://github.com/' + slug + '.git', into])
		return into
	} catch (error) {
		return unavailable(
			'could not clone ' + slug + ': ' + (error instanceof Error ? error.message.split('\n')[0] : String(error))
		)
	}
}

/**
 * The size and test story of one source tree.
 *
 * Both sides go through here, ours included: `src/` lines, then whatever the same walk
 * finds. Measuring our own tree by one rule and upstream's by another is how a size
 * comparison stops being one, and `test/` happens to be a directory name this recognises,
 * so nothing about our layout needs special-casing to be counted.
 * @param {string} tree - repository root
 * @param {(evidence: {testScripts: string[], specFiles: number, testDirs: string[]}) => unknown} coverage - how this side's statement coverage is obtained
 * @returns {Record<string, unknown>}
 */
function sourceHealth(tree, coverage) {
	const { testLines, ...evidence } = testEvidence(tree)
	return {
		lines: sourceLines(path.join(tree, 'src')),
		testLines,
		statementCoverage: coverage(evidence),
		testEvidence: evidence,
	}
}

/**
 * Statement coverage for a tree that is not this repo: only ever the fact that it was not
 * measured, and which of the two reasons applies.
 *
 * Running someone else's suite would mean installing their whole devDependency tree and
 * trusting their config, and this is not the harness for that. So where a suite exists the
 * honest record is that we did not measure it — and where none does, that is the finding.
 * @param {{testScripts: string[], specFiles: number, testDirs: string[]}} evidence
 * @returns {string | import('./unavailable.mjs').Unavailable}
 */
function coverageFromEvidence(evidence) {
	const automated = evidence.testScripts.length > 0 || evidence.specFiles > 0 || evidence.testDirs.length > 0
	return automated ? unavailable('a suite exists but this harness does not run it') : 'no automated suite'
}

/**
 * The size and test story of a source tree that has to be cloned first.
 * @param {string | null} slug
 * @param {string} into
 * @param {boolean} reuse
 * @returns {Promise<Record<string, unknown>>}
 */
async function clonedSource(slug, into, reuse) {
	if (slug === null) return { lines: unavailable('the manifest names no GitHub repository') }
	const tree = await shallowClone(slug, into, reuse)
	if (typeof tree !== 'string') return { lines: tree, testLines: tree, statementCoverage: tree }
	return sourceHealth(tree, coverageFromEvidence)
}

/**
 * Both projects' health, from the GitHub and npm APIs and from the two source trees.
 * @param {object} opts
 * @param {string} opts.workDir
 * @param {Record<string, any>} opts.upstreamManifest - the installed `pptxgenjs` manifest
 * @param {boolean} [opts.reuse] - keep an existing upstream clone instead of refetching
 * @returns {Promise<Record<string, unknown>>}
 */
export async function measureHealth({ workDir, upstreamManifest, reuse = false }) {
	const ownManifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
	const upstreamSlug = githubSlug(upstreamManifest)

	return {
		'ts-pptx': {
			...(await githubHealth(githubSlug(ownManifest))),
			npm: await npmHealth([ownManifest.name, ALIAS_NAME]),
			source: sourceHealth(ROOT, ownCoverage),
		},
		pptxgenjs: {
			...(await githubHealth(upstreamSlug)),
			npm: await npmHealth([upstreamManifest.name]),
			source: await clonedSource(upstreamSlug, path.join(workDir, 'upstream-src'), reuse),
		},
	}
}
