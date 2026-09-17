#!/usr/bin/env node
/**
 * Turn the committed snapshot into the comparison pages and the README summary.
 *
 * This is the cheap half of a deliberate split. `./measure.mjs` is expensive and needs the
 * network: it builds the corpus with both libraries, installs upstream, packs ours, clones
 * a tree and calls two public APIs, and it runs on release cadence. This reads the JSON
 * that run left behind and writes markdown. It touches no network, builds nothing, and
 * finishes in milliseconds, which is what lets `--check` sit in a per-push gate.
 *
 * The consequence is the point: a push can never silently move a published number, because
 * numbers only move when someone re-measures, and a push never pays for measuring one
 * either. The pages print the date they were measured on, which is the honest way to present
 * a figure refreshed on a slower cadence than the file around it.
 *
 * ## Four outputs, one source
 *
 * `docs/comparison.md` is the short page a reader choosing a library reads,
 * `docs/comparison-method.md` carries the method and every full table, and
 * `docs/comparison-syntax.md` prints the calls behind each row. All three are written whole.
 * `README.md` is not: it is hand-written prose with one generated region between
 * {@link REGION_START} and {@link REGION_END}, spliced in place so the rest of the file
 * survives. Every output comes from the same snapshot, so the short versions cannot date while
 * the long version moves. That failure mode is the entire reason the README block is generated
 * rather than typed once and forgotten.
 *
 * The pages are **committed**, unlike `docs/doc-index.md`, which is generated and gitignored.
 * They have to be: a fresh checkout must build the docs site without network access and
 * without installing another library.
 *
 * ## What the prose may say
 *
 * Only what a measurement backs. "10 of 22 probes" is a reading; "more complete" is a
 * judgement a reader can reasonably make the other way, and a comparison written by one of
 * the two subjects has to leave that judgement to them. The hand-written passages here are
 * the concessions and the framing, both of which are the kind of statement that gets less
 * trustworthy, not more, when a generator writes it.
 *
 * A row whose measurement is missing is dropped rather than blanked. An empty cell in a
 * comparison table reads as a measured zero, and on these pages a zero is never neutral. The
 * page says how many rows it dropped, so an omission is visible as an omission.
 *
 * Prose is reflowed by {@link wrap} rather than hard-wrapped in the source. Half of these
 * paragraphs interpolate a number whose width changes with the measurement, and a
 * hand-wrapped line that was tidy at four digits is ragged at eight.
 */
import fs from 'node:fs'
import path from 'node:path'
import { isMain, parseCli, ROOT, runCli } from '../script-utils.mjs'
import { isUnavailable } from './unavailable.mjs'

const SNAPSHOT = path.join(ROOT, 'scripts', 'comparison', 'snapshot.json')
const PAGE = path.join(ROOT, 'docs', 'comparison.md')
const METHOD_PAGE = path.join(ROOT, 'docs', 'comparison-method.md')
const SYNTAX_PAGE = path.join(ROOT, 'docs', 'comparison-syntax.md')
const README = path.join(ROOT, 'README.md')

/** The generated region inside `README.md`. Everything between these lines is rewritten. */
const REGION_START = '<!-- comparison:start -->'
const REGION_END = '<!-- comparison:end -->'

/** Column at which generated prose is reflowed. Tables and frontmatter are exempt. */
const WIDTH = 90

/**
 * The two subjects, named rather than derived from the snapshot.
 *
 * The snapshot could carry a third and this file would still only render two, because the
 * prose is written about these two specifically: what ts-pptx gives up is ours to state,
 * and the adoption sentence is about a gap with a direction. Generalising the tables to N
 * subjects while the prose stayed bilateral would be fake generality, and it would hide the
 * moment a third subject actually arrives behind a page that renders it half-right.
 */
const OURS = 'ts-pptx'
const UPSTREAM = 'pptxgenjs'
const COLUMNS = [OURS, UPSTREAM]

const USAGE = `Usage: node scripts/comparison/render.mjs [--check]

Renders scripts/comparison/snapshot.json into docs/comparison.md,
docs/comparison-method.md, docs/comparison-syntax.md and the generated region
of README.md.

Options:
  --check   report drift and exit 1; write nothing
  -h, --help`

/**
 * @typedef {{version: string, published?: string, source: string}} Subject
 * @typedef {{id: string, label: string, group: string, construct: string, part: string,
 *   results: Record<string, string>, source?: Record<string, string>,
 *   notes?: Record<string, string>}} CoverageRow
 * @typedef {{id: string, type: string, partUri?: string, description: string, decks: number}} Diagnostic
 * @typedef {{decks: number, errors: number, cleanDecks: number, byType: Record<string, number>,
 *   diagnostics: Diagnostic[], notBuilt: Record<string, number>}} SubjectValidity
 */

/**
 * @typedef {object} Snapshot
 * @property {string} generatedAt
 * @property {Record<string, Subject>} subjects
 * @property {CoverageRow[]} coverage
 * @property {string[]} upstreamAhead
 * @property {string[]} sharedGaps
 * @property {Record<string, any>} validity
 * @property {Record<string, any>} hygiene
 * @property {Record<string, any>} timing
 * @property {Record<string, any>} health
 */

/**
 * How each construct family is titled.
 *
 * Kept in `groups.json` beside the snapshot rather than in this file, because the site's
 * coverage chart titles the same groups, and two copies of a label list are how a table and a
 * chart come to disagree. A group the corpus grows without a label still renders, capitalised,
 * rather than failing the build: a missing heading is a cosmetic problem, and stopping the docs
 * build over one would be out of proportion.
 * @type {Record<string, string>}
 */
const GROUP_LABELS = JSON.parse(fs.readFileSync(path.join(ROOT, 'scripts', 'comparison', 'groups.json'), 'utf8'))

/**
 * The four outcomes, as the tables print them.
 * @type {Record<string, string>}
 */
const OUTCOME_LABELS = {
	emitted: 'emitted',
	absent: 'absent',
	'no-api': 'no API',
	error: 'error',
}

/**
 * Reflow one paragraph to {@link WIDTH}.
 *
 * Long words are never broken, so a table-free paragraph carrying a long URL simply runs
 * over rather than being corrupted into two half-links.
 * @param {string} text
 * @param {string} [hang] indent for continuation lines
 * @returns {string[]}
 */
function wrap(text, hang = '') {
	const words = text.trim().split(/\s+/)
	/** @type {string[]} */
	const lines = []
	let line = words.shift() ?? ''
	for (const word of words) {
		if (line.length + 1 + word.length > WIDTH) {
			lines.push(line)
			line = hang + word
		} else {
			line += ' ' + word
		}
	}
	lines.push(line)
	return lines
}

/**
 * One paragraph, reflowed, followed by the blank line that ends it.
 * @param {string} text
 * @returns {string[]}
 */
const para = (text) => [...wrap(text), '']

/**
 * One list item, reflowed with a hanging indent.
 * @param {string} text
 * @returns {string[]}
 */
const bullet = (text) => wrap('- ' + text, '  ')

/**
 * A fixed locale, because a gate that re-renders the page has to get the same bytes on
 * every machine that runs it.
 * @param {number} value
 * @returns {string}
 */
const num = (value) => value.toLocaleString('en-US')

/**
 * Kibibytes, to one decimal, labelled as what they are.
 *
 * The divisor is 1024 because `scripts/bundle-size-ratchet.mjs` divides by 1024 too, and every
 * size this project has published sits on that footing, so a figure here stays comparable with
 * the gates' own. The label is `KiB` because that is what dividing by 1024 gives.
 *
 * The decimal is for the bundle table. At whole kibibytes that column prints one figure five
 * times over, and five identical rows read as a broken table rather than as the finding they
 * are: neither library splits along feature lines, so what a program calls barely moves what it
 * bundles. The digit is what shows the rows were measured separately.
 * @param {number} bytes
 * @returns {string}
 */
const kib = (bytes) => (bytes / 1024).toFixed(1) + ' KiB'

/**
 * Mebibytes, to one decimal, labelled for the reason {@link kib} is.
 * @param {number} bytes
 * @returns {string}
 */
const mib = (bytes) => (bytes / 1024 / 1024).toFixed(1) + ' MiB'

/** @param {string} value @returns {string} */
const code = (value) => '`' + value + '`'

/** @param {string} group @returns {string} */
const groupLabel = (group) => GROUP_LABELS[group] ?? group.charAt(0).toUpperCase() + group.slice(1)

/** @param {Snapshot} snapshot @param {string} subject @returns {SubjectValidity | undefined} */
const validityOf = (snapshot, subject) => snapshot.validity?.[subject]

/** @param {Snapshot} snapshot @param {string} subject @returns {any} */
const hygieneOf = (snapshot, subject) => snapshot.hygiene?.[subject]

/** @param {Snapshot} snapshot @param {string} subject @returns {any} */
const healthOf = (snapshot, subject) => snapshot.health?.[subject]

/** @param {CoverageRow} row @param {string} subject @returns {boolean} */
const emitted = (row, subject) => row.results[subject] === 'emitted'

/**
 * One comparison row, or `null` when either side is missing.
 *
 * Both sides or neither: a row with one cell filled is not a comparison, and it invites the
 * reader to read the hole as a result. The caller counts the nulls and says so.
 * @param {string} label
 * @param {unknown[]} cells one per column, in {@link COLUMNS} order
 * @param {(value: any) => string} format
 * @returns {string | null}
 */
function comparedRow(label, cells, format) {
	if (cells.some((cell) => cell === undefined || isUnavailable(cell))) return null
	return '| ' + label + ' | ' + cells.map((cell) => format(cell)).join(' | ') + ' |'
}

/**
 * ts-pptx against pptxgenjs, as a percentage of pptxgenjs.
 *
 * A percentage rather than a ratio because the question a reader brings is "how much more
 * or less than the one I already use", and because a ratio under 1 is the shape people
 * misread. The sign is the whole point: positive means ts-pptx costs more.
 *
 * `null` where the arithmetic says nothing: a list, a version string, or a denominator of
 * zero. A near-zero difference is spelled out rather than rounded to `0%`, which would
 * claim two measurements met exactly.
 * @param {unknown[]} cells one per column, in {@link COLUMNS} order
 * @returns {string | null}
 */
function percentDelta(cells) {
	const [ours, theirs] = cells
	if (typeof ours !== 'number' || typeof theirs !== 'number' || theirs === 0) return null
	const change = ((ours - theirs) / theirs) * 100
	if (Math.abs(change) < 0.5) return 'within 1%'
	return (change > 0 ? '+' : '') + change.toFixed(0) + '%'
}

/**
 * A compared row carrying the percentage as a fourth cell.
 *
 * Same both-or-neither rule as {@link comparedRow}, since it is the same row with one more
 * column on it.
 * @param {string} label
 * @param {unknown[]} cells one per column, in {@link COLUMNS} order
 * @param {(value: any) => string} format
 * @returns {string | null}
 */
function deltaRow(label, cells, format) {
	if (cells.some((cell) => cell === undefined || isUnavailable(cell))) return null
	const delta = percentDelta(cells)
	return '| ' + label + ' | ' + cells.map((cell) => format(cell)).join(' | ') + ' | ' + (delta ?? 'not a ratio') + ' |'
}

/**
 * Assemble a table, dropping the rows that could not be measured and saying how many.
 * @param {string[]} headers
 * @param {(string | null)[]} rows
 * @returns {string[]}
 */
function table(headers, rows) {
	const kept = /** @type {string[]} */ (rows.filter((row) => row !== null))
	const dropped = rows.length - kept.length
	const lines = ['| ' + headers.join(' | ') + ' |', '|' + headers.map(() => '---').join('|') + '|', ...kept, '']
	if (dropped > 0)
		lines.push(
			...para(
				dropped === 1
					? 'One row is not shown: that measurement could not be taken when the snapshot was written.'
					: `${dropped} rows are not shown: those measurements could not be taken when the snapshot was written.`
			)
		)
	return lines
}

/**
 * The banner every generated surface carries, worded as `scripts/docs-index.mjs` words it.
 * @param {string} kind what the banner sits on top of
 * @returns {string[]}
 */
function banner(kind) {
	return [
		`<!-- GENERATED ${kind}. Do not edit by hand.`,
		'     Regenerate with `pnpm run comparison:render`.',
		'     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->',
	]
}

/**
 * A page's lines as the file is written: runs of blank lines collapsed, one trailing newline.
 * @param {string[]} lines
 * @returns {string}
 */
function finish(lines) {
	return (
		lines
			.join('\n')
			.replace(/\n{3,}/g, '\n\n')
			.trimEnd() + '\n'
	)
}

/**
 * A list as English writes one, with the last item joined by "and".
 * @param {string[]} items
 * @returns {string}
 */
function listOf(items) {
	if (items.length <= 1) return items.join('')
	return items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1]
}

/** @param {Snapshot} snapshot @returns {string} */
function measuredOn(snapshot) {
	const ours = snapshot.subjects[OURS]
	const upstream = snapshot.subjects[UPSTREAM]
	return (
		`Measured on ${snapshot.generatedAt}: ts-pptx ${ours?.version ?? 'unknown'} built from this ` +
		`repository, against pptxgenjs ${upstream?.version ?? 'unknown'} installed from npm` +
		(upstream?.published ? ` (published ${upstream.published}).` : '.')
	)
}

/* ────────────────────────────────────────────────────────────────────────────
 * The short page
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function frontmatter(snapshot) {
	const ours = snapshot.subjects[OURS]?.version ?? ''
	const upstream = snapshot.subjects[UPSTREAM]?.version ?? ''
	return [
		'---',
		'doc-schema-version: 1',
		'title: "ts-pptx vs PptxGenJS"',
		`summary: "What ts-pptx ${ours} and pptxgenjs ${upstream} each emit, how many of their decks validate, what each costs to bundle and install, and how the two projects are run."`,
		'read_when:',
		'  - Choosing between ts-pptx and pptxgenjs',
		'  - Checking whether a construct is emitted by one library or by both',
		'  - Weighing what ts-pptx gives up against what it adds',
		'doc_type: "overview"',
		'---',
	]
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionIntro(snapshot) {
	const ours = snapshot.subjects[OURS]
	const upstream = snapshot.subjects[UPSTREAM]
	return [
		'# ts-pptx vs PptxGenJS',
		'',
		...para(
			`ts-pptx ${ours?.version ?? ''} and pptxgenjs ${upstream?.version ?? ''} were measured on ` +
				`${snapshot.generatedAt} by building the same ${snapshot.coverage.length} deck intents with each ` +
				'library and reading the bytes that came out. ts-pptx descends from pptxgenjs, detached at its ' +
				'v4.0.1 ([lineage](getting-started/introduction.md#lineage)), so every difference below comes ' +
				'from running both rather than from either one describing itself.'
		),
	]
}

/**
 * The concessions, hand-written and placed before any number.
 *
 * Only the adoption figures are interpolated. The rest is a standing statement of what this
 * package does not do, and it belongs where a reader meets it before the tables rather than
 * in a footnote under them: a page that puts its own weaknesses first is the only kind whose
 * strengths are worth reading.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionBeforeYouChoose(snapshot) {
	const ourNpm = healthOf(snapshot, OURS)?.npm
	const upstreamNpm = healthOf(snapshot, UPSTREAM)?.npm
	const lines = [
		'## Before you choose',
		'',
		...bullet('**Node.js 24 or later.** pptxgenjs declares no engine floor and runs on much older releases.'),
		...bullet(
			'**One ESM build.** `require("pptx-ts")` works through the ESM interop Node has had since 22.12, ' +
				'and a browser loads the package as a module. pptxgenjs also ships CommonJS and a ' +
				'classic-script bundle that defines a global. See [where it runs](getting-started/runtime.md).'
		),
		...bullet(
			'**Not a drop-in continuation of the upstream release line.** The API is close by descent, not ' +
				'by contract, and it has moved since. Moving code across is a port, not an upgrade: ' +
				'[porting from PptxGenJS](comparison-syntax.md) lists the calls that change.'
		),
		...bullet(
			'**No SmartArt on the write side, in either library.** It is not a difference between them, ' +
				'but it is a real gap in both.'
		),
	]
	if (
		typeof upstreamNpm?.downloadsLastMonth === 'number' &&
		typeof ourNpm?.downloadsLastMonth === 'number' &&
		upstreamNpm.downloadsLastMonth > ourNpm.downloadsLastMonth
	)
		lines.push(
			...bullet(
				`**Adoption is not close.** pptxgenjs was downloaded ${num(upstreamNpm.downloadsLastMonth)} times in ` +
					`the last month, against ${num(ourNpm.downloadsLastMonth)} for ts-pptx. That gap buys answers ` +
					'that already exist, examples written by people other than the maintainer, and good odds that ' +
					'a bug on a common path was hit by someone else first. If that outweighs the differences ' +
					'below, use pptxgenjs.'
			)
		)
	lines.push('')
	return lines
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionScorecard(snapshot) {
	const total = snapshot.coverage.length
	/** @type {any[]} */
	const programs = snapshot.hygiene?.programs ?? []
	const first = programs[0]
	return [
		'## Scorecard',
		'',
		...table(
			['', 'ts-pptx', 'pptxgenjs'],
			[
				comparedRow(
					'Intents emitted',
					COLUMNS.map((subject) => snapshot.coverage.filter((row) => emitted(row, subject)).length),
					(count) => `${num(count)} of ${num(total)}`
				),
				comparedRow(
					'Decks with no schema error',
					COLUMNS.map((subject) => validityOf(snapshot, subject)),
					(validity) => `${num(validity.cleanDecks)} of ${num(validity.decks)}`
				),
				first
					? comparedRow(
							`${first.label}, bundled and gzipped`,
							COLUMNS.map((subject) => hygieneOf(snapshot, subject)?.bundles?.[first.id]?.initialBytes),
							kib
						)
					: null,
				comparedRow(
					'Runtime dependencies, transitive',
					COLUMNS.map((subject) => hygieneOf(snapshot, subject)?.dependencies?.transitive),
					num
				),
				comparedRow(
					'Installed size, with dependencies',
					COLUMNS.map((subject) => hygieneOf(snapshot, subject)?.install?.bytes),
					mib
				),
			]
		),
		...para(
			'The bundled size is what a browser program fetches before its first line runs. ' +
				'[How the comparison was measured](comparison-method.md#package-hygiene) has every install ' +
				'and bundle figure, and the programs behind them.'
		),
	]
}

/**
 * Coverage as sentences: the text a reader without the chart, and `llms.txt`, still gets.
 *
 * Every intent lands in exactly one of the four sets, so the four sentences carry the whole
 * matrix. The one about pptxgenjs is stated even when it is empty, because an empty set there
 * is the result most worth a reader's checking.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionCoverageSummary(snapshot) {
	const rows = snapshot.coverage
	const total = rows.length
	const counts = COLUMNS.map((subject) => rows.filter((row) => emitted(row, subject)).length)
	const labels = (/** @type {CoverageRow[]} */ set) => listOf(set.map((row) => row.label))
	const both = rows.filter((row) => COLUMNS.every((subject) => emitted(row, subject)))
	const oursOnly = rows.filter((row) => emitted(row, OURS) && !emitted(row, UPSTREAM))
	const upstreamOnly = rows.filter((row) => !emitted(row, OURS) && emitted(row, UPSTREAM))
	const neither = rows.filter((row) => COLUMNS.every((subject) => !emitted(row, subject)))

	const lines = [
		'## Construct coverage',
		'',
		'<CoverageMatrix />',
		'',
		...para(`Of ${total} intents, ts-pptx emits ${counts[0]} and pptxgenjs emits ${counts[1]}.`),
		...para(both.length > 0 ? `Emitted by both: ${labels(both)}.` : 'No intent is emitted by both libraries.'),
		...para(
			oursOnly.length > 0
				? `Emitted by ts-pptx only: ${labels(oursOnly)}.`
				: 'No intent is emitted by ts-pptx and not by pptxgenjs.'
		),
		...para(
			upstreamOnly.length > 0
				? `Emitted by pptxgenjs only: ${labels(upstreamOnly)}.`
				: 'No intent is emitted by pptxgenjs and not by ts-pptx.'
		),
		...para(
			neither.length > 0
				? `Emitted by neither: ${labels(neither)}.`
				: 'Every intent is emitted by at least one of the two libraries.'
		),
	]
	const missed = rows.flatMap((row) =>
		COLUMNS.filter((subject) => !emitted(row, subject)).map((subject) => row.results[subject])
	)
	if (missed.length > 0 && missed.every((outcome) => outcome === 'no-api'))
		lines.push(
			...para(
				'Every intent a library does not emit is one it has no API for, rather than one it tried and ' + 'got wrong.'
			)
		)
	lines.push(
		...para(
			'[How the comparison was measured](comparison-method.md#construct-coverage) has the token each ' +
				'intent is read for, the part it is read from, and the notes on individual rows.'
		)
	)
	return lines
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionValiditySummary(snapshot) {
	const oracle = snapshot.validity?.oracle
	const lines = ['## Schema validity', '']
	if (!oracle || isUnavailable(oracle)) return [...lines, ...para('Not measured when this snapshot was written.')]
	lines.push('<ValidityBars />', '')
	const total = snapshot.coverage.length

	for (const subject of COLUMNS) {
		const validity = validityOf(snapshot, subject)
		if (!validity) continue
		const notBuilt = Object.values(validity.notBuilt ?? {}).reduce((sum, count) => sum + count, 0)
		const clean =
			validity.cleanDecks === validity.decks
				? `all ${num(validity.decks)}`
				: validity.cleanDecks === 0
					? 'none'
					: `${num(validity.cleanDecks)}`
		lines.push(
			...para(
				`${subject} built ${num(validity.decks)} of the ${total} decks, and ${clean} validated cleanly` +
					(validity.errors > 0 ? ` (${num(validity.errors)} errors in all)` : '') +
					'.' +
					(notBuilt > 0
						? ` It had no API, and so no deck, for the remaining ${notBuilt === 1 ? 'one' : num(notBuilt)}.`
						: '')
			)
		)
	}
	lines.push(
		...para(
			`Every deck went through the Open XML SDK validator (${oracle.sdkVersion}) at the ` +
				`${code(oracle.format)} conformance target. The decks a library could not build are counted ` +
				'because a library with fewer decks has fewer decks to be wrong in. ' +
				'[How the comparison was measured](comparison-method.md#schema-validity) lists each distinct error.'
		)
	)
	return lines
}

/**
 * The signed mean difference per mode, as a percentage of pptxgenjs, but only while every deck
 * points the same way in both modes: faster compressed and slower stored. `null` otherwise.
 *
 * Both pages state that reading, and both have to stop stating it the release it stops being
 * true, which is why it is computed once, here, rather than written into either.
 * @param {any} timing
 * @returns {{compressed: number, stored: number} | null}
 */
function modeDeltas(timing) {
	/** @type {any[]} */
	const cases = timing.cases ?? []
	/** @param {string} mode @returns {number[]} */
	const deltas = (mode) =>
		cases
			.map((row) => {
				const measured = row.modes?.[mode]
				if (!measured || isUnavailable(measured)) return null
				const ours = measured[OURS]?.median
				const theirs = measured[UPSTREAM]?.median
				return typeof ours === 'number' && typeof theirs === 'number' && theirs > 0
					? ((ours - theirs) / theirs) * 100
					: null
			})
			.filter((value) => value !== null)

	const compressed = deltas('deflate')
	const stored = deltas('store')
	if (cases.length === 0 || compressed.length !== cases.length || stored.length !== cases.length) return null
	if (!(compressed.every((value) => value < 0) && stored.every((value) => value > 0))) return null

	// Both are ratios of a ratio, so the mean of the row-wise percentages is the honest
	// summary: no row is weighted by how big its deck happened to be.
	const mean = (/** @type {number[]} */ values) => values.reduce((sum, value) => sum + value, 0) / values.length
	return { compressed: mean(compressed), stored: mean(stored) }
}

/**
 * A ratio of two medians, to two decimals below ten.
 * @param {number} value
 * @returns {string}
 */
const times = (value) => (value < 10 ? value.toFixed(2) : value.toFixed(1)) + '×'

/**
 * Every timing case as ts-pptx's median over pptxgenjs's, one column per mode.
 * @param {any} timing
 * @returns {string[]}
 */
function ratioTable(timing) {
	/** @type {any[]} */
	const cases = timing.cases ?? []
	return table(
		['Deck', 'Compressed', 'Stored'],
		cases.map((row) => {
			const ratios = ['deflate', 'store'].map((mode) => {
				const measured = row.modes?.[mode]
				if (!measured || isUnavailable(measured)) return undefined
				const ours = measured[OURS]?.median
				const theirs = measured[UPSTREAM]?.median
				return typeof ours === 'number' && typeof theirs === 'number' && theirs > 0 ? ours / theirs : undefined
			})
			if (ratios.some((ratio) => ratio === undefined)) return null
			return `| ${row.label} | ${ratios.map((ratio) => times(/** @type {number} */ (ratio))).join(' | ')} |`
		})
	)
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionTimingSummary(snapshot) {
	/** @type {any} */
	const timing = snapshot.timing
	if (!timing?.cases?.length) return []
	const flip = modeDeltas(timing)
	const lines = ['## Generation time', '', '<TimingRatio />', '']
	if (flip)
		lines.push(
			...para(
				'Compressed, which is what a file you intend to keep gets, ts-pptx is faster on every deck, by ' +
					`${Math.abs(flip.compressed).toFixed(0)}% on average. Stored, with compression turned off, it is ` +
					`slower on every deck, by ${flip.stored.toFixed(0)}% on average: its XML generation and package ` +
					'assembly cost more than pptxgenjs, and its compressor more than makes that back.'
			)
		)
	lines.push(
		...para(
			'Each cell is the ts-pptx median divided by the pptxgenjs median for the same deck, so a figure ' +
				'below 1× means ts-pptx took less time.'
		),
		...ratioTable(timing),
		...para(
			'The milliseconds, the machine they were taken on, and why the two settings point in opposite ' +
				'directions are in [how the comparison was measured](comparison-method.md#generation-time).'
		)
	)
	return lines
}

/**
 * The read side, deliberately not a table.
 *
 * There is nothing to compare: one library has the capability and the other does not. A
 * table would invite a score, and "wins the read side 3 to 0" is a sentence about a contest
 * nobody entered.
 * @returns {string[]}
 */
function sectionReadingDecks() {
	return [
		'## Reading decks',
		'',
		...para('pptxgenjs generates decks and does not read them. ts-pptx also reads:'),
		...bullet(
			'[Inspection](reference/pptx-inspection.md) reports what a package contains without parsing it into ' + 'a model.'
		),
		...bullet(
			'[Reading](reading/read-and-edit.md) loads a deck into an object model, edits it in place, and writes ' +
				'the package back out.'
		),
		...bullet(
			'[Deck to script](reference/pptx-to-script.md) turns a deck into the TypeScript that rebuilds it, ' +
				'reporting what it could not express rather than dropping it.'
		),
		'',
	]
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function healthTable(snapshot) {
	const rows = COLUMNS.map((subject) => healthOf(snapshot, subject))
	return table(
		['', 'ts-pptx', 'pptxgenjs'],
		[
			comparedRow(
				'Repository',
				rows.map((row) => row?.repo),
				(value) => `[${value}](https://github.com/${value})`
			),
			comparedRow(
				'Default branch',
				rows.map((row) => row?.defaultBranch),
				code
			),
			comparedRow(
				'Last commit on the default branch',
				rows.map((row) => row?.lastDefaultBranchCommit),
				String
			),
			comparedRow(
				'Last npm publish',
				rows.map((row) => row?.npm?.lastPublish),
				String
			),
			comparedRow(
				'Downloads, last month',
				rows.map((row) => row?.npm?.downloadsLastMonth),
				num
			),
			comparedRow(
				'Stars',
				rows.map((row) => row?.stars),
				num
			),
			comparedRow(
				'Open issues',
				rows.map((row) => row?.openIssues),
				num
			),
			comparedRow(
				'Open pull requests',
				rows.map((row) => row?.openPullRequests),
				num
			),
			comparedRow(
				'Source lines',
				rows.map((row) => row?.source?.lines),
				num
			),
			comparedRow(
				'Test lines',
				rows.map((row) => row?.source?.testLines),
				num
			),
			comparedRow(
				'Test suite',
				rows.map((row) => row?.source?.testEvidence),
				formatTestEvidence
			),
			comparedRow(
				'Statement coverage',
				rows.map((row) => row?.source?.statementCoverage),
				formatCoverage
			),
		]
	)
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionHealthSummary(snapshot) {
	return [
		'## Adoption and project health',
		'',
		...para(
			'How the two projects are run, kept apart from everything above because stars and downloads ' +
				'measure history as much as merit.'
		),
		...healthTable(snapshot),
		...para(
			'[How the comparison was measured](comparison-method.md#project-health) says how each of these ' +
				'figures is taken.'
		),
	]
}

/** @returns {string[]} */
function sectionMore() {
	return [
		'## More',
		'',
		...bullet(
			'[How the comparison was measured](comparison-method.md): the corpus, every full table, and how each ' +
				'figure was taken.'
		),
		...bullet(
			'[Porting from PptxGenJS](comparison-syntax.md): the calls that change between the libraries, and the code behind every row.'
		),
		'',
	]
}

/**
 * The short page: what a reader choosing a library needs, with a link to the rest.
 * @param {Snapshot} snapshot
 * @returns {string}
 */
export function renderPage(snapshot) {
	return finish([
		...frontmatter(snapshot),
		'',
		...banner('FILE'),
		'',
		...sectionIntro(snapshot),
		...sectionBeforeYouChoose(snapshot),
		...sectionScorecard(snapshot),
		...sectionCoverageSummary(snapshot),
		...sectionValiditySummary(snapshot),
		...sectionTimingSummary(snapshot),
		...sectionReadingDecks(),
		...sectionHealthSummary(snapshot),
		...sectionMore(),
	])
}

/* ────────────────────────────────────────────────────────────────────────────
 * The method page
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function methodFrontmatter(snapshot) {
	const ours = snapshot.subjects[OURS]?.version ?? ''
	const upstream = snapshot.subjects[UPSTREAM]?.version ?? ''
	return [
		'---',
		'doc-schema-version: 1',
		'title: "How the comparison was measured"',
		`summary: "The corpus, the four outcomes and every full table behind the comparison of ts-pptx ${ours} with pptxgenjs ${upstream}: validation, installs, bundles, generation time and project health, and how each was taken."`,
		'read_when:',
		'  - Checking a number on the comparison page',
		'  - Adding a probe or a program to the comparison corpus',
		'  - Re-measuring the comparison for a release',
		'doc_type: "reference"',
		'---',
	]
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionCorpus(snapshot) {
	return [
		'# How the comparison was measured',
		'',
		...para(
			'Every figure on [ts-pptx vs PptxGenJS](comparison.md) comes from `scripts/comparison/snapshot.json`, ' +
				'which is refreshed on release cadence, and nothing on either page is edited by hand. This page ' +
				'is the method behind those figures and the full tables they summarise.'
		),
		...para(measuredOn(snapshot)),
		'## The corpus',
		'',
		...para(
			`The corpus is ${snapshot.coverage.length} deck intents. Each one states an intent ("a slide that ` +
				'enters with a push transition"), and each library expresses that intent in its own ' +
				"idiom. Transcribing one library's calls into the other is how a comparison gets " +
				'rigged, so the two arms of a probe deliberately do not have to look alike. Both decks ' +
				'are then opened, and the part the probe names is read for the token it names.'
		),
		...para('Four outcomes are possible, per probe per library:'),
		'| Outcome | Meaning |',
		'|---|---|',
		'| emitted | the token is present in the named part |',
		'| absent | an API exists, and the output does not carry the token |',
		'| no API | nothing in the public surface expresses the intent |',
		'| error | the build threw |',
		'',
		...para(
			'`no API` is the only one of the four that is a claim rather than a reading, so it is ' +
				"checked rather than trusted: that library's shipped bundle is searched for the token, " +
				'and a hit fails the measurement run unless the corpus carries a written reason for it. ' +
				'Those reasons are printed under the table they belong to.'
		),
		...para(
			'Two things a reader should price in. The corpus is ours, so it was chosen by an ' +
				'interested party. It is kept honest in two specific ways: it carries a probe neither ' +
				'library can satisfy, and the set of probes upstream emits and ts-pptx does not is ' +
				'reported below even when it is empty, so an empty set is a stated result rather than ' +
				'something a reader has to infer from a gap. A pull request that adds a probe is ' +
				'welcome, including one ts-pptx fails.'
		),
	]
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionCoverage(snapshot) {
	const counts = COLUMNS.map((subject) => snapshot.coverage.filter((row) => emitted(row, subject)).length)
	const total = snapshot.coverage.length
	const labelOf = (/** @type {string} */ id) => snapshot.coverage.find((row) => row.id === id)?.label ?? id
	const baseline = snapshot.coverage.filter((row) => row.group === 'shared').length

	const lines = [
		'## Construct coverage',
		'',
		...para(`Of ${total} probes, ts-pptx emitted ${counts[0]} and pptxgenjs emitted ${counts[1]}.`),
		...para(
			'"Looked for" is the token the harness reads for, in the part named beside it. It is the OOXML ' +
				'element in every case but one, where the intent is speaker notes and the token is the note ' +
				'text itself. [Porting from PptxGenJS](comparison-syntax.md) prints the calls behind every row.'
		),
	]

	/** @type {string[]} */
	const groups = []
	for (const row of snapshot.coverage) if (!groups.includes(row.group)) groups.push(row.group)

	for (const group of groups) {
		const rows = snapshot.coverage.filter((row) => row.group === group)
		lines.push(
			`### ${groupLabel(group)}`,
			'',
			'| Intent | Looked for | Part | ts-pptx | pptxgenjs |',
			'|---|---|---|---|---|'
		)
		for (const row of rows) {
			const outcomes = COLUMNS.map((subject) => OUTCOME_LABELS[row.results[subject] ?? ''] ?? row.results[subject])
			lines.push(`| ${row.label} | ${code(row.construct)} | ${code(row.part)} | ${outcomes[0]} | ${outcomes[1]} |`)
		}
		lines.push('')
		const notes = rows.flatMap((row) =>
			Object.entries(row.notes ?? {}).flatMap(([subject, note]) => bullet(`${row.label}, ${subject}: ${note}.`))
		)
		if (notes.length > 0) lines.push(...notes, '')
	}

	lines.push(
		...para(
			'The shared baseline is the control group. A corpus holding only constructs one side ' +
				'cannot produce would prove that the corpus was chosen, not that the libraries differ, ' +
				`so ${baseline} of the probes are ones both libraries are expected to pass. A failure there ` +
				'fails the measurement run instead of becoming a row on this page.'
		),
		...para(
			snapshot.upstreamAhead.length === 0
				? 'No probe in this corpus is emitted by pptxgenjs and not by ts-pptx.'
				: 'Emitted by pptxgenjs and not by ts-pptx: ' + snapshot.upstreamAhead.map(labelOf).join(', ') + '.'
		),
		...para(
			snapshot.sharedGaps.length === 0
				? 'Every probe in this corpus is emitted by at least one of the two libraries.'
				: 'Emitted by neither library: ' + snapshot.sharedGaps.map(labelOf).join(', ') + '.'
		)
	)
	return lines
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionValidity(snapshot) {
	const oracle = snapshot.validity?.oracle
	const lines = [
		'## Schema validity',
		'',
		...para(
			'**This validates the decks this corpus builds, not either library in general.** A deck ' +
				'no probe builds is not covered by any of it, and a library can be perfectly correct on ' +
				'everything these probes never touch.'
		),
	]
	if (!oracle || isUnavailable(oracle)) {
		lines.push(...para('Not measured when this snapshot was written.'))
		return lines
	}

	const rows = COLUMNS.map((subject) => validityOf(snapshot, subject))
	lines.push(
		...para(
			`Every deck the corpus built was passed through the Open XML SDK validator (${oracle.sdkVersion}) at ` +
				`the ${code(oracle.format)} conformance target: the same oracle, and the same target, that this ` +
				"project's own `test:schema` suite uses."
		),
		...table(
			['', 'ts-pptx', 'pptxgenjs'],
			[
				comparedRow(
					'Decks validated',
					rows.map((row) => row?.decks),
					num
				),
				comparedRow(
					'Decks with no error',
					rows.map((row) => row?.cleanDecks),
					num
				),
				comparedRow(
					'Errors',
					rows.map((row) => row?.errors),
					num
				),
				comparedRow(
					'Intents with no deck to validate',
					rows.map((row) =>
						row ? Object.values(row.notBuilt ?? {}).reduce((sum, count) => sum + count, 0) : undefined
					),
					num
				),
			]
		),
		...para(
			'The last row is the denominator a validity count needs. A library that builds fewer ' +
				'decks has fewer decks to be wrong in, and reading the error counts without it would ' +
				'reward not having an API.'
		),
		...para(
			'There is no warning column. This validator reports a single severity, so a zero in a ' +
				'second column would be a number nobody measured.'
		)
	)

	for (const subject of COLUMNS) {
		const validity = validityOf(snapshot, subject)
		if (!validity || validity.diagnostics.length === 0) continue
		lines.push(
			`### What failed in the ${subject} decks`,
			'',
			...para(
				'Distinct diagnostics rather than a raw error total. One fault repeated across every ' +
					'deck, and that many unrelated faults, are different facts about a library, and a ' +
					'total on its own cannot tell them apart.'
			),
			'| Diagnostic | Part | Decks |',
			'|---|---|---|'
		)
		for (const diagnostic of validity.diagnostics)
			lines.push(
				`| ${code(diagnostic.id)} | ${code(diagnostic.partUri ?? 'unknown part')} | ${num(diagnostic.decks)} |`
			)
		lines.push('')
		// The validator's own wording, in a code span: these messages carry angle brackets and
		// brace patterns that markdown would otherwise read as an autolink or as emphasis.
		for (const diagnostic of validity.diagnostics)
			lines.push(...bullet(`${code(diagnostic.partUri ?? 'unknown part')}: ${code(diagnostic.description.trim())}`))
		lines.push('')
	}
	return lines
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionHygiene(snapshot) {
	const rows = COLUMNS.map((subject) => hygieneOf(snapshot, subject))
	const lines = [
		'## Package hygiene',
		'',
		...para(
			'What a consumer gets. Each library was installed on its own into an empty directory, ' +
				'upstream from the registry and ts-pptx from a pack of this working tree, so nothing ' +
				'here is measured against a development checkout with its dependencies hoisted flat.'
		),
		...table(
			['', 'ts-pptx', 'pptxgenjs', 'Difference'],
			[
				deltaRow(
					'Installed size, with dependencies',
					rows.map((row) => row?.install?.bytes),
					mib
				),
				deltaRow(
					'Installed size, the package alone',
					rows.map((row) => row?.install?.packageBytes),
					mib
				),
				deltaRow(
					'Runtime dependencies, transitive',
					rows.map((row) => row?.dependencies?.transitive),
					num
				),
			]
		),
		...para(
			'The last column is ts-pptx measured against pptxgenjs, so a positive number is ours ' +
				'costing more and a negative one is ours costing less. It is a percentage of the ' +
				'pptxgenjs figure rather than a difference in bytes, because two of its rows are in ' +
				'mebibytes and the third is a count, and a reader comparing them needs a ' +
				'number that does not change meaning between rows.'
		),
	]

	const ours = hygieneOf(snapshot, OURS)
	const upstream = hygieneOf(snapshot, UPSTREAM)
	if (typeof ours?.install?.bytes === 'number' && ours.install.bytes > upstream?.install?.bytes)
		lines.push(
			...para(
				'ts-pptx installs larger than pptxgenjs despite carrying fewer dependencies. The largest ' +
					'share of that weight is source maps: `dist/` ships a `.js.map` beside every module, and ' +
					'each one embeds the original TypeScript. The unminified `.js` is the next largest share. ' +
					'No consumer build keeps either, which is why the bundled figures below are much closer ' +
					'together than the installed ones.'
			)
		)

	lines.push(
		...table(
			['', 'ts-pptx', 'pptxgenjs'],
			[
				comparedRow(
					'Runtime dependencies, direct',
					rows.map((row) => row?.dependencies?.direct),
					(value) => (value.length === 0 ? 'none' : value.map(code).join(', '))
				),
				comparedRow(
					'Entry points',
					rows.map((row) => row?.entryPoints),
					(value) => value.map(code).join(', ')
				),
				comparedRow(
					'Module formats',
					rows.map((row) => row?.moduleFormats),
					(value) => value.join(', ')
				),
				comparedRow(
					'`engines.node`',
					rows.map((row) => row?.engines ?? null),
					(value) => (value === null ? 'not declared' : code(value))
				),
			]
		),
		...sectionBundles(snapshot)
	)
	return lines
}

/**
 * What a consumer's build weighs, over a corpus of whole decks rather than one program.
 *
 * A hello world is one point on a curve, and it is the point where two tree-shaken bundles
 * are most alike, because nearly everything either library holds has been shaken out of it.
 * The corpus climbs to a deck using every construct both libraries emit, so the table can
 * be read for what the second slide costs as well as the first.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionBundles(snapshot) {
	/** @type {any[]} */
	const programs = snapshot.hygiene?.programs ?? []
	const bundles = COLUMNS.map((subject) => hygieneOf(snapshot, subject)?.bundles)
	if (programs.length === 0) return []

	// What the corpus was built to find out, read off the two ends of it. Placed with the
	// table rather than after the method, because it is the sentence a reader needs before
	// they start doing the subtraction themselves.
	const first = programs[0]
	const last = programs[programs.length - 1]
	const growth = bundles.map((byId) => {
		const from = byId?.[first.id]?.initialBytes
		const to = byId?.[last.id]?.initialBytes
		return typeof from === 'number' && typeof to === 'number' ? to - from : undefined
	})
	const reading =
		programs.length > 1 && growth.every((value) => typeof value === 'number')
			? para(
					`The column is nearly flat, and that is the result. From ${first.label.toLowerCase()} to ` +
						`${last.label.toLowerCase()}, ts-pptx grows by ${kib(growth[0] ?? 0)} and pptxgenjs by ` +
						`${kib(growth[1] ?? 0)}, which is about what the programs' own literals weigh. Neither ` +
						'library splits along feature lines: importing either one costs almost everything it ' +
						'will ever cost, and the deck written afterwards is close to free. So a hello world was ' +
						'never a flattering measurement of either library, and a consumer weighing bundle size ' +
						'is choosing between two roughly fixed costs rather than between two slopes.'
				)
			: []

	return [
		'### What a bundled program costs',
		'',
		...para(
			`Each row is a whole deck both libraries build: ${programs.length} consumer programs, from the ` +
				'smallest one anyone writes up to one using every construct the shared baseline above ' +
				'shows both of them emitting. Each is written in its own idiom on both sides, and the ' +
				'calls behind every row are on [porting from PptxGenJS](comparison-syntax.md).'
		),
		...programs.flatMap((program) => bullet(`**${program.label}.** ${program.what}`)),
		'',
		...table(
			['Program', 'ts-pptx', 'pptxgenjs', 'Difference'],
			programs.map((program) =>
				deltaRow(
					program.label,
					bundles.map((byId) => byId?.[program.id]?.initialBytes),
					kib
				)
			)
		),
		...reading,
		...para(
			'Both columns construct the library the way every consumer of pptxgenjs constructs it, ' +
				'with the class that carries everything. ts-pptx has a lower floor than that, reached by ' +
				'composing a presentation from only the construct families a program uses, and ' +
				'[smaller bundles](bundle-size.md) carries those figures. It is deliberately not a row here: ' +
				'pptxgenjs has no counterpart to compose, so the cell beside it would be empty and the ' +
				'percentage would be comparing two different programs.'
		),
		...para(
			'Every program is identical in intent on both sides. Each is bundled with esbuild for the ' +
				'browser, minified, and gzipped at level 9, following the conventions ' +
				'`scripts/bundle-size-ratchet.mjs` documents, with one difference that matters. The ' +
				'ratchet never bundles, so it cannot drop unreachable code and its figures are an upper ' +
				'bound on what the package ships; this bundles and does tree-shake, because a ' +
				"consumer's build is precisely the thing being compared here. **The two sets of numbers " +
				'will not agree, and neither is wrong.**'
		),
		...para(
			'Code splitting is on, so each figure is the entry chunk: what the program pays before its ' +
				'first line runs. Chunks a bundler defers behind a dynamic import are not counted, ' +
				'because a program that never takes that path never fetches them.'
		),
		...para(
			'Each program is run against both libraries before it is bundled. A bundler compiles a ' +
				'call that does not exist, and a misspelled method comes out as a smaller bundle rather ' +
				'than as an error, because the tree-shaker keeps less: running the programs first is ' +
				'what stops a typo being published here as a saving.'
		),
	]
}

/**
 * Milliseconds, to one decimal below ten and to none above it.
 *
 * A deck that takes 4.7 ms and one that takes 838 ms are both on this page, and one format
 * cannot serve both: a decimal on the large figure is precision the measurement does not
 * have, and no decimal on the small one rounds two different rows to the same number.
 * @param {number} ms
 * @returns {string}
 */
const millis = (ms) => (ms < 10 ? ms.toFixed(1) : ms.toFixed(0)) + ' ms'

/**
 * One mode's table: every deck, both libraries, and the difference between them.
 * @param {any} timing
 * @param {string} modeId
 * @returns {string[]}
 */
function timingTable(timing, modeId) {
	/** @type {any[]} */
	const cases = timing.cases ?? []
	return table(
		['Deck', 'ts-pptx', 'pptxgenjs', 'Difference'],
		cases.map((row) => {
			const measured = row.modes?.[modeId]
			if (!measured || isUnavailable(measured)) return null
			return deltaRow(
				row.label,
				COLUMNS.map((subject) => measured[subject]?.median),
				millis
			)
		})
	)
}

/**
 * Generation time, which is the one measurement here that a clock took.
 *
 * Its own section rather than a row in package hygiene, because everything else on this
 * page is a fact about an artifact and this is a fact about a machine running a library.
 * The section is largely method for that reason: a time nobody can see the conditions
 * behind is not a measurement, it is a claim.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionTiming(snapshot) {
	/** @type {any} */
	const timing = snapshot.timing
	if (!timing?.cases?.length) return []
	const machine = timing.machine ?? {}
	/** @type {any[]} */
	const modes = timing.modes ?? []

	const lines = [
		'## Generation time',
		'',
		...para(
			'How long each library takes to turn a deck into bytes: the same decks the bundle table ' +
				'above weighs, plus three larger ones built for this measurement alone, because the ' +
				'largest program up there is three slides and a clock has almost nothing to see in it. ' +
				'The calls behind every row are on [porting from PptxGenJS](comparison-syntax.md).'
		),
		...para(
			'**A `.pptx` is a zip, so the compression setting is not a detail of this measurement, it ' +
				'is the measurement.** The two libraries do not default to the same one. ts-pptx ' +
				'deflates unless told not to. pptxgenjs passes no compression option to JSZip on the ' +
				'`outputType` path, and JSZip stores by default. Its own `compression` argument is ' +
				'honoured on the stream and browser paths and ignored on the one in between, which is ' +
				'the path `writeFile` takes in Node. Timing the two default calls against each other ' +
				'would compare deflating with not deflating and report the difference as a library ' +
				'being slow, so both tables below are matched pairs.'
		),
	]

	for (const mode of modes) {
		const heading = mode.id === 'deflate' ? 'Compressed' : 'Stored, the control'
		lines.push(
			`### ${heading}`,
			'',
			...para(
				mode.id === 'deflate'
					? 'Both libraries asked for a compressed deck, which is what a consumer writing a file ' +
							'they intend to keep gets. This is the table that matters.'
					: 'Both libraries asked not to compress. The zip drops out of the measurement, leaving ' +
							"each library's own work: building the XML and assembling the package."
			),
			...timingTable(timing, mode.id)
		)
	}

	const flip = modeDeltas(timing)
	if (flip)
		lines.push(
			...para(
				'The two tables point in opposite directions, and that is the finding. Stored, ts-pptx is ' +
					`slower on every deck, by ${flip.stored.toFixed(0)}% on average, so our XML generation and ` +
					"package assembly cost more than upstream's. Compressed, ts-pptx is faster on every deck, by " +
					`${Math.abs(flip.compressed).toFixed(0)}% on average, because fflate deflates faster than ` +
					'JSZip does and the compressor dominates the total. A consumer writing a file they intend ' +
					'to keep gets the first table. A consumer who has turned compression off gets the second, ' +
					'and should know that is where we are behind.'
			)
		)

	lines.push(
		...para(
			'The measurement is a median over repeated rounds, taken after a warm-up that is thrown ' +
				'away, with the two libraries interleaved and the order alternated so that a machine ' +
				'which slows down mid-run cannot hand either column a result it did not earn. Building ' +
				'the deck and writing it are both inside the clock; constructing the presentation ' +
				'object is not.'
		),
		...para(
			'**The milliseconds belong to the machine that took them and do not transfer; the ratios ' +
				'mostly do.** These were taken on ' +
				`${machine.cpu ?? 'an unrecorded CPU'} (${machine.cores ?? '?'} cores) under Node ` +
				`${machine.node ?? '?'} on ${machine.platform ?? '?'}` +
				(machine.measured ? `, on ${machine.measured}` : '') +
				'. Repeating a run on the same machine moves a difference by a few points in either ' +
				'direction, so read the columns for their direction and rough size rather than for ' +
				'their last digit.'
		)
	)
	return lines
}

/**
 * @param {any} coverage
 * @returns {string}
 */
function formatCoverage(coverage) {
	if (typeof coverage === 'string') return coverage
	if (typeof coverage?.pct !== 'number') return 'not measured'
	const lane = coverage.lane === 'merged' ? 'Node and browser lanes merged' : `${coverage.lane} lane only`
	return `${coverage.pct}% (${lane})`
}

/**
 * @param {any} evidence
 * @returns {string}
 */
function formatTestEvidence(evidence) {
	if (!evidence) return 'not measured'
	const scripts = evidence.testScripts?.length ?? 0
	const specs = evidence.specFiles ?? 0
	/** @type {string[]} */
	const dirs = evidence.testDirs ?? []
	if (scripts === 0 && specs === 0 && dirs.length === 0) return 'no test script, no spec file, no test directory'
	const files =
		dirs.length > 0
			? `${num(specs)} spec files under ${dirs.map((dir) => code(dir + '/')).join(', ')}`
			: `${num(specs)} spec files`
	return `${num(scripts)} test scripts, ${files}`
}

/**
 * How the project health figures are chosen, kept in their own section for the reason they
 * have their own snapshot key: they measure the projects rather than what they emit, and
 * mixing them into the tables above would let a star count read as a property of the output.
 * The table itself is on the short page.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionHealth(snapshot) {
	const lines = [
		'## Project health',
		'',
		...para(
			'The [health table](comparison.md#adoption-and-project-health) describes how the two projects ' +
				'are run, not what either one emits. Stars and downloads measure adoption, adoption measures ' +
				'history as much as merit, and none of it belongs in the same table as a construct a library ' +
				'does or does not write. This is how each figure in it is taken.'
		),
		...para(
			"The last commit on the default branch is reported rather than the repository's last " +
				'push, which the same API offers and which counts activity on any branch. The two ' +
				'disagree for pptxgenjs by several months, and reporting the later one would say ' +
				'something the default branch does not support.'
		),
		...para(
			'Line counts come from the same walk on both sides: every code file under `src/`, raw ' +
				'lines with comments and blanks included, and test lines are spec files plus anything ' +
				'under a test directory, counted once each. No normalisation makes two libraries ' +
				'formatted to different rules comparable, and a large part of the ts-pptx figure is the ' +
				'documentation comments the bundled sizes above shed. Read it as an order of magnitude ' +
				'for how much there is to maintain, and as nothing at all about whether it is good.'
		),
	]

	const upstreamEvidence = healthOf(snapshot, UPSTREAM)?.source?.testEvidence
	if (
		upstreamEvidence &&
		upstreamEvidence.testScripts?.length === 0 &&
		upstreamEvidence.specFiles === 0 &&
		upstreamEvidence.testDirs?.length === 0
	)
		lines.push(
			...para(
				'The empty pptxgenjs test row is what this walk can see, and it is not the same claim ' +
					'as untested. That repository documents a manual, demo-driven process instead, which ' +
					'nothing measured here can weigh. The row is about an automated suite, and the ' +
					'coverage figure beside it exists for ts-pptx only because there is a suite to ' +
					'instrument.'
			)
		)

	const ourNpm = healthOf(snapshot, OURS)?.npm
	if (ourNpm?.names?.length > 1 && ourNpm.downloadsByName)
		lines.push(
			...para(
				'ts-pptx is published under two names carrying the same bytes, ' +
					ourNpm.names.map(code).join(' and ') +
					'. The download figure is their sum (' +
					Object.entries(ourNpm.downloadsByName)
						.map(([name, count]) => code(name) + ' ' + num(Number(count)))
						.join(', ') +
					'), because either name alone understates the total, and the canonical name alone ' +
					'happens to understate it by most.'
			)
		)

	const upstream = healthOf(snapshot, UPSTREAM)
	if (upstream?.lastDefaultBranchCommit && upstream?.npm?.lastPublish)
		lines.push(
			...para(
				`The pptxgenjs column shows no npm release since ${upstream.npm.lastPublish} and no commit on ` +
					`${code(upstream.defaultBranch)} since ${upstream.lastDefaultBranchCommit}. That is what the two ` +
					'APIs report, and it is all these pages say about it: from outside, a stable library ' +
					'that has stopped needing changes looks exactly like one between maintainers, and ' +
					'this measurement cannot tell them apart. It is worth weighing either way, next to ' +
					`${num(upstream.openIssues ?? 0)} open issues and ${num(upstream.openPullRequests ?? 0)} open pull ` +
					'requests.'
			)
		)
	return lines
}

/**
 * The method page: the corpus, every full table, and how each figure was taken.
 * @param {Snapshot} snapshot
 * @returns {string}
 */
export function renderMethodPage(snapshot) {
	return finish([
		...methodFrontmatter(snapshot),
		'',
		...banner('FILE'),
		'',
		...sectionCorpus(snapshot),
		...sectionCoverage(snapshot),
		...sectionValidity(snapshot),
		...sectionHygiene(snapshot),
		...sectionTiming(snapshot),
		...sectionHealth(snapshot),
	])
}

/* ────────────────────────────────────────────────────────────────────────────
 * The syntax page
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * The probes both libraries have an arm for, and how many of those arms are the same code.
 *
 * Both figures are read off the recorded sources rather than asserted, because the second
 * one is a claim about how close the two APIs still are and that is exactly the kind of
 * claim a comparison written by one of the two parties should not be making by hand.
 * @param {Snapshot} snapshot
 * @returns {{both: CoverageRow[], identical: CoverageRow[]}}
 */
function armAgreement(snapshot) {
	const both = snapshot.coverage.filter((row) => COLUMNS.every((subject) => row.source?.[subject]))
	return { both, identical: both.filter((row) => row.source?.[OURS] === row.source?.[UPSTREAM]) }
}

/**
 * One fenced JavaScript block.
 * @param {string} source
 * @returns {string[]}
 */
function fence(source) {
	return ['```js', ...source.split('\n'), '```', '']
}

/**
 * Where two arms differ, as runs of changed lines.
 *
 * A longest-common-subsequence alignment rather than a comparison by position: one arm often
 * carries a line the other does not (`type: 'bar'` inside the options, where pptxgenjs passes the
 * chart type as an argument), and comparing by position would then mark every later line as
 * changed. Lines are compared and returned trimmed, so indentation alone is never a difference.
 * @param {string} ours
 * @param {string} upstream
 * @returns {Array<{ours: string[], upstream: string[]}>}
 */
export function lineHunks(ours, upstream) {
	const a = ours.split('\n').map((line) => line.trim())
	const b = upstream.split('\n').map((line) => line.trim())
	const width = b.length + 1
	// The longest common subsequence of a[i..] and b[j..], at i * width + j.
	const lengths = new Int32Array((a.length + 1) * width)
	const at = (/** @type {number} */ i, /** @type {number} */ j) => lengths[i * width + j] ?? 0
	for (let i = a.length - 1; i >= 0; i--)
		for (let j = b.length - 1; j >= 0; j--)
			lengths[i * width + j] = a[i] === b[j] ? at(i + 1, j + 1) + 1 : Math.max(at(i + 1, j), at(i, j + 1))

	/** @type {Array<{ours: string[], upstream: string[]}>} */
	const hunks = []
	/** @type {{ours: string[], upstream: string[]}} */
	let open = { ours: [], upstream: [] }
	const close = () => {
		if (open.ours.length > 0 || open.upstream.length > 0) hunks.push(open)
		open = { ours: [], upstream: [] }
	}
	let i = 0
	let j = 0
	while (i < a.length || j < b.length) {
		if (i < a.length && j < b.length && a[i] === b[j]) {
			close()
			i++
			j++
		} else if (j >= b.length || (i < a.length && at(i + 1, j) >= at(i, j + 1))) {
			open.ours.push(a[i++] ?? '')
		} else {
			open.upstream.push(b[j++] ?? '')
		}
	}
	close()
	return hunks
}

/**
 * The calls where a port stops being a rename.
 *
 * Every run of lines that differs between the two arms of a probe, a bundle program or the timing
 * deck, with identical runs merged into one row that names every place it appears. Read off the
 * recorded sources like everything else on the page, so a difference cannot be listed here that
 * the corpus no longer has.
 * @param {Snapshot} snapshot
 * @returns {Array<{ours: string, upstream: string, where: string[]}>}
 */
function callDifferences(snapshot) {
	/** @type {Array<{label: string, source?: Record<string, string>}>} */
	const sources = [
		...snapshot.coverage,
		...(snapshot.hygiene?.programs ?? []),
		...(snapshot.timing?.shape?.source ? [{ label: 'The timing decks', source: snapshot.timing.shape.source }] : []),
	]
	/** @type {Map<string, {ours: string, upstream: string, where: string[]}>} */
	const found = new Map()
	for (const { label, source } of sources) {
		const ours = source?.[OURS]
		const upstream = source?.[UPSTREAM]
		if (!ours || !upstream || ours === upstream) continue
		for (const hunk of lineHunks(ours, upstream)) {
			const row = { ours: hunk.ours.join(' '), upstream: hunk.upstream.join(' ') }
			const key = row.ours + '\0' + row.upstream
			const existing = found.get(key)
			if (!existing) found.set(key, { ...row, where: [label] })
			else if (!existing.where.includes(label)) existing.where.push(label)
		}
	}
	return [...found.values()]
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionCallDifferences(snapshot) {
	const differences = callDifferences(snapshot)
	const cell = (/** @type {string} */ value) => (value === '' ? 'nothing' : code(value.replaceAll('|', '\\|')))
	const lines = ['## The calls that change', '']
	if (differences.length === 0)
		return [...lines, ...para('Every intent and program in the corpus is called identically in both libraries.')]
	lines.push(
		...para(
			'Each row is a run of lines that differs between the two arms of an intent or a program printed ' +
				'further down, with every place it appears. Everything else in those arms is the same code.'
		),
		'| ts-pptx | pptxgenjs | Where |',
		'|---|---|---|',
		...differences.map((row) => `| ${cell(row.ours)} | ${cell(row.upstream)} | ${row.where.join(', ')} |`),
		''
	)
	return lines
}

/**
 * What a port changes besides the calls, written by hand because none of it is a measurement.
 * @returns {string[]}
 */
function sectionAroundTheCalls() {
	return [
		'## What changes around the calls',
		'',
		...bullet(
			'**The import.** `import TsPptx from "pptx-ts"` in place of `import pptxgen from "pptxgenjs"`. From ' +
				'CommonJS, `const { default: TsPptx } = require("pptx-ts")`: the package is an ES module, so ' +
				'`require()` returns its namespace and the class is on `.default`.'
		),
		...bullet('**Node.js 24 or later.** ts-pptx declares `>=24`; pptxgenjs runs on older releases.'),
		...bullet(
			'**One build.** Node, bundlers and browsers all load the same ES module through the package ' +
				'exports, so there is no separate CommonJS or browser file to choose. See ' +
				'[where it runs](getting-started/runtime.md).'
		),
		...bullet(
			'**No upstream file paths.** Code that loads `dist/pptxgen.cjs.js`, `dist/pptxgen.js`, ' +
				'`dist/pptxgen.es.js`, `dist/pptxgen.bundle.js` or `dist/pptxgen.min.js` by path imports the ' +
				'package instead, and the global the classic-script bundle defined has no equivalent. A page ' +
				'with no build step uses `import TsPptx from "https://esm.sh/pptx-ts/browser"` in a module script.'
		),
		'',
	]
}

/**
 * One probe, as the two libraries express it.
 *
 * Three shapes, and which one a probe gets is decided by the snapshot rather than by an
 * editor: two blocks where the arms differ, one block where they are character-identical,
 * and our arm alone where upstream has no API. The identical case is collapsed on purpose --
 * printing the same line twice under two headings would make a reader compare two things
 * that are the same, and the heading is the finding.
 * @param {CoverageRow} row
 * @returns {string[]}
 */
function probeSection(row) {
	const ours = row.source?.[OURS]
	const upstream = row.source?.[UPSTREAM]
	const outcome = (/** @type {string} */ subject) =>
		OUTCOME_LABELS[row.results[subject] ?? ''] ?? row.results[subject] ?? 'not measured'
	const lines = [
		`### ${row.label}`,
		'',
		...para(
			`${code(row.construct)} in ${code(row.part)}. ` + `${OURS}: ${outcome(OURS)}. ${UPSTREAM}: ${outcome(UPSTREAM)}.`
		),
	]

	if (ours && upstream && ours === upstream) {
		lines.push(...para('Both libraries, called identically:'), ...fence(ours))
	} else if (ours && upstream) {
		lines.push(`**${OURS}**`, '', ...fence(ours), `**${UPSTREAM}**`, '', ...fence(upstream))
	} else if (ours) {
		lines.push(`**${OURS}**`, '', ...fence(ours))
	}

	// No sentence for the ordinary "upstream has no API" case: the status line above already
	// said so, and eleven repetitions of it would train a reader to skip the line that carries
	// the sighting note on the one probe that has one.
	if (!ours && !upstream) lines.push(...para('Neither library has an API for this intent, so neither has a block.'))
	else if (!upstream && row.notes?.[UPSTREAM])
		lines.push(...para(`The token is in the ${UPSTREAM} bundle regardless. It ${row.notes[UPSTREAM]}.`))
	// A note on a subject that *does* have an arm is the harness reporting what happened when
	// that code ran -- a throw, or a part that came back without the construct in it. It belongs
	// under the block it is about, since the block alone would read as working code.
	const ran = COLUMNS.filter((subject) => row.source?.[subject] && row.notes?.[subject])
	if (ran.length > 0) lines.push(...ran.flatMap((subject) => bullet(`${subject}: ${row.notes?.[subject]}.`)), '')
	return lines
}

/**
 * The bundle corpus as code, under the same rule as the probes above it.
 *
 * These are whole decks rather than single constructs, and the method page reports what a
 * consumer's build weighs for each of them. The programs belong on this page for the reason
 * the probes do: a size that nobody can see the program behind is a number a reader has to
 * take on trust.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function bundleCorpusSection(snapshot) {
	/** @type {any[]} */
	const programs = snapshot.hygiene?.programs ?? []
	if (programs.length === 0) return []
	/** @type {Record<string, string>} */
	const frames = snapshot.hygiene?.frame ?? {}

	const lines = [
		'## The bundle corpus',
		'',
		...para(
			'The [comparison](comparison.md) also reports what a bundler leaves after tree-shaking, ' +
				'over whole decks rather than single constructs. These are those programs. Each one ' +
				'builds the same deck with both libraries, using only constructs the shared baseline ' +
				'shows both of them emitting, and each is run before it is bundled.'
		),
	]
	if (frames[OURS] && frames[UPSTREAM])
		lines.push(
			...para(
				'The frame is a consumer program rather than the probe harness, so it imports the ' +
					'package root and keeps the exported bytes:'
			),
			...fence(frames[OURS]),
			...para('and, for the other column:'),
			...fence(frames[UPSTREAM])
		)

	for (const program of programs) {
		const ours = program.source?.[OURS]
		const upstream = program.source?.[UPSTREAM]
		lines.push(`### ${program.label}`, '', ...para(program.what))
		if (ours && upstream && ours === upstream)
			lines.push(...para('Both libraries, called identically:'), ...fence(ours))
		else if (ours && upstream) lines.push(`**${OURS}**`, '', ...fence(ours), `**${UPSTREAM}**`, '', ...fence(upstream))
	}
	return lines
}

/**
 * The scale decks behind the generation-time table.
 *
 * Only the three large ones. The five programs timed alongside them are the bundle corpus
 * printed above, and a second copy of each would be a second thing to keep true.
 *
 * One block rather than two wherever the arms agree, on the rule the rest of this page
 * follows, which here means each deck is printed once, since the two libraries differ in a
 * single `addChart` call.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function timingCorpusSection(snapshot) {
	/** @type {any} */
	const shape = snapshot.timing?.shape
	const ours = shape?.source?.[OURS]
	const upstream = shape?.source?.[UPSTREAM]
	if (!ours || !upstream) return []
	/** @type {number[]} */
	const sizes = shape.sizes ?? []

	const lines = [
		'## The timing corpus',
		'',
		...para(
			'The [comparison](comparison.md) also reports how long each library takes to turn a deck ' +
				'into bytes. The small end of that measurement is the bundle corpus above; the large ' +
				'end is this, one deck shape built at ' +
				(sizes.length > 0 ? listOf(sizes.map((size) => String(size))) + ' slides' : 'several sizes') +
				'. It is printed once because the slide count is the only thing that changes between ' +
				'them, and it is deliberately dull: a timing corpus is not hunting for the slowest ' +
				'construct, it is making the per-slide cost visible.'
		),
		...para('`slides` below is that count. Everything else is the same code at every size.'),
	]
	if (ours === upstream) lines.push(...para('Both libraries, called identically:'), ...fence(ours))
	else lines.push(`**${OURS}**`, '', ...fence(ours), `**${UPSTREAM}**`, '', ...fence(upstream))
	return lines
}

/**
 * The whole corpus as code, one page.
 * @param {Snapshot} snapshot
 * @returns {string}
 */
export function renderSyntaxPage(snapshot) {
	const ours = snapshot.subjects[OURS]
	const upstream = snapshot.subjects[UPSTREAM]
	const { both, identical } = armAgreement(snapshot)

	const lines = [
		'---',
		'doc-schema-version: 1',
		'title: "Porting from PptxGenJS"',
		`summary: "Moving a deck script from pptxgenjs ${upstream?.version ?? ''} to ts-pptx ${ours?.version ?? ''}: the calls that change, what changes around them, and every intent in the comparison corpus as code in both libraries."`,
		'read_when:',
		'  - Porting a deck script from pptxgenjs to ts-pptx',
		'  - Reading a comparison row and wanting the calls behind it',
		'  - Looking for the ts-pptx call that emits a particular construct',
		'  - Reading a bundle size and wanting the program that was measured',
		'doc_type: "guide"',
		'---',
		'',
		...banner('FILE'),
		'',
		'# Porting from PptxGenJS',
		'',
		...para(
			'ts-pptx descends from pptxgenjs, so much of a pptxgenjs script carries across as it is: of the ' +
				`${both.length} intents both libraries build in the comparison corpus, ${identical.length} are ` +
				'called with identical code. This page is the rest: the calls that change, what changes around ' +
				'them, and then every intent and program in the corpus as code in both libraries.'
		),
		...para(
			`Measured on ${snapshot.generatedAt}: ts-pptx ${ours?.version ?? ''} built from this ` +
				`repository, against pptxgenjs ${upstream?.version ?? ''} installed from npm.`
		),
		...sectionCallDifferences(snapshot),
		...sectionAroundTheCalls(),
		'## How to read a snippet',
		'',
		...para(
			'Every snippet is the code the comparison ran. The harness lifts it out of the build function as ' +
				'it measures and records it in the snapshot beside the outcome it produced, so no snippet here ' +
				"can illustrate a row that some earlier version of it produced. Each intent is written in the library's " +
				'own idiom rather than transcribed from one into the other, and where the two arms come out the ' +
				'same anyway the page prints one block and says so.'
		),
		...para(
			'Each block is the body of a build function. The harness puts the same frame around every ' +
				'one of them, so this page states the frame once rather than repeating it on every intent:'
		),
		...fence(
			[
				"import TsPptx from 'pptx-ts/node'",
				'',
				'const pres = new TsPptx()',
				'// the snippet goes here',
				"await pres.writeFile({ fileName: 'probe.pptx' })",
			].join('\n')
		),
		...para('and, for the other column:'),
		...fence(
			[
				"import PptxGenJS from 'pptxgenjs'",
				'',
				'const pres = new PptxGenJS()',
				'// the snippet goes here',
				"await pres.writeFile({ fileName: 'probe.pptx' })",
			].join('\n')
		),
		...para(
			'A snippet that needs a value the corpus declares once, such as a data URL or a chart ' +
				'series, carries that declaration above it. Those lines come from the value the ' +
				'measurement used, cut short with an ellipsis past 60 characters, and a path is written ' +
				'relative to the repository root.'
		),
		...para(
			'A library with no API for an intent has no block. The line under each heading says what the ' +
				'harness read for both of them.'
		),
	]

	/** @type {string[]} */
	const groups = []
	for (const row of snapshot.coverage) if (!groups.includes(row.group)) groups.push(row.group)
	for (const group of groups) {
		lines.push(`## ${groupLabel(group)}`, '')
		for (const row of snapshot.coverage.filter((row) => row.group === group)) lines.push(...probeSection(row))
	}

	lines.push(...bundleCorpusSection(snapshot), ...timingCorpusSection(snapshot))

	lines.push(
		'## Adding one',
		'',
		...para(
			'The corpus is `scripts/comparison/probes.mjs`, one object per intent, and both arms of a ' +
				'probe are ordinary code. A pull request that adds an intent is welcome, including one ' +
				'ts-pptx fails. The harness reports the four outcomes it reads, and the comparison page ' +
				'prints an intent upstream emits and we do not rather than dropping it.'
		),
		...para(
			'The bundle corpus is `scripts/comparison/programs.mjs`, one object per program, on one ' +
				'extra condition: both arms have to build the same deck. A program only one library can ' +
				'build measures two different pieces of work and reports the difference as a size.'
		)
	)

	return finish(lines)
}

/* ────────────────────────────────────────────────────────────────────────────
 * The README region, and the command
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * The README summary: the same snapshot, shorter, pointing at the pages for the rest.
 * @param {Snapshot} snapshot
 * @returns {string}
 */
export function renderReadmeRegion(snapshot) {
	const total = snapshot.coverage.length
	const counts = COLUMNS.map((subject) => snapshot.coverage.filter((row) => emitted(row, subject)).length)
	const ourHealth = healthOf(snapshot, OURS)
	const upstreamHealth = healthOf(snapshot, UPSTREAM)
	const ourValidity = validityOf(snapshot, OURS)
	const upstreamValidity = validityOf(snapshot, UPSTREAM)

	const lines = [
		REGION_START,
		...banner('REGION'),
		'',
		'## How this compares with PptxGenJS',
		'',
		...para(
			'ts-pptx is an independent derivative of ' +
				'[PptxGenJS](https://github.com/gitbrent/PptxGenJS), detached at its v4.0.1. Both were ' +
				`measured on ${snapshot.generatedAt} by building the same ${total} deck intents with each ` +
				'library and reading the bytes that came out.'
		),
		...bullet(
			`**Construct coverage:** ts-pptx emitted ${counts[0]} of ${total}, pptxgenjs ${counts[1]} of ${total}. ` +
				'Nothing in the corpus is emitted by pptxgenjs and not by ts-pptx.'
		),
	]
	if (ourValidity && upstreamValidity)
		lines.push(
			...bullet(
				`**Schema validity:** of the decks each library built, ${ourValidity.cleanDecks} of ` +
					`${ourValidity.decks} ts-pptx decks and ${upstreamValidity.cleanDecks} of ${upstreamValidity.decks} ` +
					'pptxgenjs decks validate with no error against the Open XML SDK.'
			)
		)
	if (typeof upstreamHealth?.npm?.downloadsLastMonth === 'number')
		lines.push(
			...bullet(
				`**Adoption:** pptxgenjs is downloaded ${num(upstreamHealth.npm.downloadsLastMonth)} times a month, ` +
					`against ${num(ourHealth?.npm?.downloadsLastMonth ?? 0)} for ts-pptx. If a large installed base ` +
					'matters to you more than the differences above, use pptxgenjs.'
			)
		)
	if (ourHealth?.lastDefaultBranchCommit && upstreamHealth?.lastDefaultBranchCommit)
		lines.push(
			...bullet(
				`**Activity:** last commit on the default branch, ${ourHealth.lastDefaultBranchCommit} for ts-pptx ` +
					`and ${upstreamHealth.lastDefaultBranchCommit} for pptxgenjs. Last npm publish, ` +
					`${ourHealth.npm?.lastPublish} and ${upstreamHealth.npm?.lastPublish}.`
			)
		)
	lines.push(
		'',
		...para(
			'Where the two libraries part company is on the [comparison page](docs/comparison.md), and ' +
				'[how it was measured](docs/comparison-method.md) has every full table. Every intent as each ' +
				'library expresses it, including the calls that differ, is on ' +
				'[porting from PptxGenJS](docs/comparison-syntax.md).'
		),
		REGION_END
	)
	return lines.join('\n')
}

/**
 * Replace the generated region of the README, leaving every hand-written line alone.
 * @param {string} readme
 * @param {string} region
 * @returns {string}
 */
export function spliceRegion(readme, region) {
	const start = readme.indexOf(REGION_START)
	const end = readme.indexOf(REGION_END)
	if (start === -1 || end === -1 || end < start)
		throw new Error(
			`README.md has no generated region. Add the ${REGION_START} and ${REGION_END} markers ` +
				'around the block this script owns, or restore them if an edit removed one.'
		)
	return readme.slice(0, start) + region + readme.slice(end + REGION_END.length)
}

/** @returns {number} process exit code */
function main() {
	const { values } = parseCli(process.argv.slice(2), {
		usage: USAGE,
		options: { check: { type: 'boolean', default: false } },
	})

	/** @type {Snapshot} */
	const snapshot = JSON.parse(fs.readFileSync(SNAPSHOT, 'utf8'))
	/** @type {Array<[string, string]>} */
	const outputs = [
		[PAGE, renderPage(snapshot)],
		[METHOD_PAGE, renderMethodPage(snapshot)],
		[SYNTAX_PAGE, renderSyntaxPage(snapshot)],
		[README, spliceRegion(fs.readFileSync(README, 'utf8'), renderReadmeRegion(snapshot))],
	]
	const stale = outputs.filter(([file, text]) => !fs.existsSync(file) || fs.readFileSync(file, 'utf8') !== text)

	if (values.check) {
		if (stale.length === 0) {
			console.log(
				'comparison: docs/comparison.md, docs/comparison-method.md, docs/comparison-syntax.md and README.md match the snapshot.'
			)
			return 0
		}
		for (const [file] of stale) console.error(path.relative(ROOT, file) + ' does not match the snapshot.')
		console.error(
			'Run `pnpm run comparison:render`. If a number moved, the snapshot was re-measured and that is the news.'
		)
		return 1
	}

	for (const [file, text] of outputs) fs.writeFileSync(file, text)
	console.log(
		stale.length === 0
			? 'comparison: already up to date.'
			: 'comparison: wrote ' + stale.map(([file]) => path.relative(ROOT, file)).join(' and ') + '.'
	)
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
