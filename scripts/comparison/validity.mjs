/**
 * Schema validity of the decks the coverage corpus built, per library.
 *
 * The oracle is `ooxml-validate` at `Microsoft365` — the same binary and the same
 * conformance target `test:schema` runs, so a diagnostic here means exactly what a
 * diagnostic there means. It is pointed at the decks `measure()` already left on disk
 * rather than at a second corpus, because a second corpus would drift from the one every
 * other row of the comparison is about.
 *
 * ## What this is evidence about
 *
 * **These decks, not "the library".** A library is not conformant or non-conformant; a
 * file either matches the schema or does not. Twenty-one clean decks say twenty-one deck
 * intents came out conformant, and that is all they say — the corpus was chosen to measure
 * construct coverage, not to hunt for schema faults, so it exercises whatever those probes
 * happen to reach and nothing else. The page must carry that sentence, and the number gets
 * printed whichever way it comes out.
 *
 * ## Errors, and the absent warning column
 *
 * The SDK validator has one severity. `ValidationResult` carries `errors` and nothing
 * else, so there is no warning count to record, and inventing a zero for one would be a
 * column that reads as a measurement and is not. What replaces it is a breakdown: errors
 * by `DiagnosticType` and by distinct diagnostic, because twelve errors that are one
 * ordering fault repeated across ten decks is a different fact from twelve unrelated ones,
 * and a bare total cannot tell them apart.
 */
import path from 'node:path'
import { validate, validatorAvailable } from 'ooxml-validate'
import { ROOT } from '../script-utils.mjs'
import { SUBJECTS } from './probes.mjs'
import { unavailable } from './unavailable.mjs'

/** How many distinct diagnostics reach the snapshot per library. Enough to characterise, not a log. */
const DIAGNOSTIC_SAMPLE = 5

/**
 * @typedef {object} DiagnosticSample
 * @property {string} id - the SDK's stable code, e.g. `Sch_UndeclaredAttribute`
 * @property {string} type - Schema, Semantic, MarkupCompatibility or Package
 * @property {string | null} partUri - the part it points at, null when unattributable
 * @property {string} description - upstream prose; reworded between SDK releases, so never asserted on
 * @property {number} decks - how many of this library's decks carry it
 */

/**
 * Which decks each library built, from `measure()`'s `decks` map.
 * @param {Record<string, Record<string, string>>} decks - probe id to subject to repo-relative path
 * @param {string} subject
 * @returns {string[]} absolute paths, in probe order
 */
function decksOf(decks, subject) {
	return Object.values(decks)
		.map((bySubject) => bySubject[subject])
		.filter((deck) => deck !== undefined)
		.map((deck) => path.join(ROOT, deck))
}

/**
 * Why the probes that produced no deck produced none, counted by outcome.
 *
 * A validity row reading "10 decks" beside another reading "21" is not a result until it
 * says where the missing eleven went, and the answer is already in the coverage table: a
 * probe with no API on that side never got as far as a file. Counting them here means the
 * page can print the denominator instead of leaving a reader to subtract.
 * @param {Array<{results: Record<string, string>}>} coverage
 * @param {string} subject
 * @returns {Record<string, number>}
 */
function notBuilt(coverage, subject) {
	/** @type {Record<string, number>} */
	const counts = {}
	for (const row of coverage) {
		const outcome = row.results[subject]
		if (outcome === 'no-api' || outcome === 'error') counts[outcome] = (counts[outcome] ?? 0) + 1
	}
	return counts
}

/**
 * Collapse one library's diagnostics into counts and a sample.
 *
 * Identity for "the same diagnostic" is `id` plus `partUri`, never `description`: the prose
 * belongs to the SDK and can be reworded in any release, while the code and the part are
 * stable. `decks` counts files, so a diagnostic repeated inside one deck counts that deck
 * once — otherwise the field reads as a deck count and is an occurrence total.
 * @param {readonly import('ooxml-validate').ValidationResult[]} results
 * @returns {{errors: number, cleanDecks: number, byType: Record<string, number>, diagnostics: DiagnosticSample[]}}
 */
function summarise(results) {
	/** @type {Record<string, number>} */
	const byType = {}
	/** @type {Map<string, DiagnosticSample>} */
	const distinct = new Map()
	let errors = 0
	let cleanDecks = 0

	for (const result of results) {
		if (result.valid) cleanDecks += 1
		const seenHere = new Set()
		for (const error of result.errors) {
			errors += 1
			byType[error.type] = (byType[error.type] ?? 0) + 1
			const key = error.id + ' ' + (error.partUri ?? '')
			const existing = distinct.get(key)
			if (existing === undefined)
				distinct.set(key, {
					id: error.id,
					type: error.type,
					partUri: error.partUri,
					description: error.description,
					decks: 1,
				})
			else if (!seenHere.has(key)) existing.decks += 1
			seenHere.add(key)
		}
	}

	const diagnostics = [...distinct.values()].sort((a, b) => b.decks - a.decks).slice(0, DIAGNOSTIC_SAMPLE)
	return { errors, cleanDecks, byType, diagnostics }
}

/**
 * Validate every deck the corpus built, per library.
 *
 * One `validate()` call per library rather than one per deck: the oracle is a .NET
 * single-file app whose startup dominates a batch this small.
 * @param {Array<{results: Record<string, string>}>} coverage - the coverage rows `measure()` produced
 * @param {Record<string, Record<string, string>>} decks - probe id to subject to repo-relative path
 * @returns {Promise<Record<string, unknown>>}
 */
export async function measureValidity(coverage, decks) {
	if (!(await validatorAvailable())) return { oracle: unavailable('the ooxml-validate binary could not be obtained') }

	/** @type {Record<string, unknown>} */
	const validity = {}
	/** @type {Record<string, string>} */
	let oracle = {}

	for (const subject of SUBJECTS) {
		const files = decksOf(decks, subject)
		if (files.length === 0) {
			validity[subject] = { decks: 0, notBuilt: notBuilt(coverage, subject) }
			continue
		}
		const report = await validate(files)
		oracle = { format: report.format, sdkVersion: report.sdkVersion }
		validity[subject] = { decks: files.length, ...summarise(report.results), notBuilt: notBuilt(coverage, subject) }
	}

	return { oracle, ...validity }
}
