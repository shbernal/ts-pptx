#!/usr/bin/env node
/**
 * Turn the committed snapshot into the comparison page and the README summary.
 *
 * This is the cheap half of a deliberate split. `./measure.mjs` is expensive and needs the
 * network: it builds the corpus with both libraries, installs upstream, packs ours, clones
 * a tree and calls two public APIs, and it runs on release cadence. This reads the JSON
 * that run left behind and writes markdown. It touches no network, builds nothing, and
 * finishes in milliseconds, which is what lets `--check` sit in a per-push gate.
 *
 * The consequence is the point: a push can never silently move a published number, because
 * numbers only move when someone re-measures, and a push never pays for measuring one
 * either. The page prints the date it was measured on, which is the honest way to present a
 * figure refreshed on a slower cadence than the file around it.
 *
 * ## Two outputs, one source
 *
 * `docs/comparison.md` is written whole. `README.md` is not: it is hand-written prose with
 * one generated region between {@link REGION_START} and {@link REGION_END}, spliced in
 * place so the rest of the file survives. Both come from the same snapshot, so the short
 * version in the README cannot date while the long version moves. That failure mode is the
 * entire reason the README block is generated rather than typed once and forgotten.
 *
 * The page is **committed**, unlike `docs/doc-index.md`, which is generated and gitignored.
 * It has to be: a fresh checkout must build the docs site without network access and
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
 * comparison table reads as a measured zero, and on this page a zero is never neutral. The
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

Renders scripts/comparison/snapshot.json into docs/comparison.md and the generated region
of README.md.

Options:
  --check   report drift and exit 1; write nothing
  -h, --help`

/**
 * @typedef {{version: string, published?: string, source: string}} Subject
 * @typedef {{id: string, label: string, group: string, construct: string, part: string,
 *   results: Record<string, string>, notes?: Record<string, string>}} CoverageRow
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
 * @property {Record<string, any>} health
 */

/**
 * How each construct family is titled in the table.
 *
 * A group the corpus grows without a label here still renders, capitalised, rather than
 * failing the build: a missing heading is a cosmetic problem, and stopping the docs build
 * over one would be out of proportion.
 * @type {Record<string, string>}
 */
const GROUP_LABELS = {
	shared: 'Shared baseline',
	motion: 'Motion',
	embedding: 'Embedding',
	shapes: 'Shapes',
	text: 'Text',
	fills: 'Fills',
	tables: 'Tables',
	charts: 'Charts',
	navigation: 'Navigation',
	diagrams: 'Diagrams',
}

/**
 * The four outcomes, as the table prints them.
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
 * Kibibytes under the label the rest of the repo uses.
 *
 * `scripts/bundle-size-ratchet.mjs` divides by 1024 and writes "kB", and every size this
 * project has published sits on that footing. Switching units here alone would make two of
 * our own numbers disagree for no reader's benefit.
 * @param {number} bytes
 * @returns {string}
 */
const kb = (bytes) => (bytes / 1024).toFixed(0) + ' kB'

/**
 * @param {number} bytes
 * @returns {string}
 */
const mb = (bytes) => (bytes / 1024 / 1024).toFixed(1) + ' MB'

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
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function frontmatter(snapshot) {
	const ours = snapshot.subjects[OURS]?.version ?? ''
	const upstream = snapshot.subjects[UPSTREAM]?.version ?? ''
	return [
		'---',
		'doc-schema-version: 1',
		'title: "Comparison With PptxGenJS"',
		`summary: "What ts-pptx ${ours} and pptxgenjs ${upstream} each emit, what validates, what each costs to install, and how the two projects are run."`,
		'read_when:',
		'  - Choosing between ts-pptx and pptxgenjs',
		'  - Checking whether a construct is emitted by one library or by both',
		'  - Weighing what ts-pptx gives up against what it adds',
		'doc_type: "reference"',
		'---',
	]
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionPremise(snapshot) {
	const ours = snapshot.subjects[OURS]
	const upstream = snapshot.subjects[UPSTREAM]
	return [
		'# Comparison With PptxGenJS',
		'',
		...para(
			'ts-pptx is an independent derivative of ' +
				'[gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS), detached at its v4.0.1 ' +
				'(see [project target](project-target.md)). Descending from a project is a poor reason ' +
				'to be trusted over it, so every difference below was produced by running both ' +
				'libraries and reading what came out.'
		),
		...para(
			`Measured on ${snapshot.generatedAt}: ts-pptx ${ours?.version ?? 'unknown'} built from this ` +
				`repository, against pptxgenjs ${upstream?.version ?? 'unknown'} installed from npm` +
				(upstream?.published ? ` (published ${upstream.published}).` : '.')
		),
		'## What this measures, and how',
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
		...para(
			'Every number on this page comes from `scripts/comparison/snapshot.json`, which is ' +
				'refreshed on release cadence and carries the date above. Nothing here is edited by hand.'
		),
	]
}

/**
 * The concessions, hand-written and placed above the tables.
 *
 * Only the adoption figures are interpolated. The rest is a standing statement of what this
 * package does not do, and it belongs where a reader meets it before the tables rather than
 * in a footnote under them: a page that puts its own weaknesses first is the only kind whose
 * strengths are worth reading.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionConcessions(snapshot) {
	const ourNpm = healthOf(snapshot, OURS)?.npm
	const upstreamNpm = healthOf(snapshot, UPSTREAM)?.npm
	const lines = [
		'## What ts-pptx gives up',
		'',
		...bullet(
			'**No CommonJS build.** pptxgenjs ships one, so it runs unchanged on Node versions and ' +
				"toolchains ts-pptx cannot serve at all. `require('pptx-ts')` does work, through the " +
				'ESM interop Node has had since 22.12, which every Node ts-pptx supports has. See ' +
				'[runtime and package support](runtime-and-package-support.md).'
		),
		...bullet(
			'**No global bundle, and no CDN script tag.** pptxgenjs can be dropped into a page with a ' +
				'`<script>` tag and used from a global. ts-pptx requires a bundler or a runtime that ' +
				'loads ES modules.'
		),
		...bullet('**Node.js `>=24` only.** pptxgenjs declares no engine floor and runs much further back.'),
		...bullet(
			'**Not a drop-in continuation of the upstream release line.** The API is close by descent, ' +
				'not by contract, and it has moved since. Migrating is a port, not an upgrade.'
		),
		...bullet(
			'**No SmartArt on the write side.** Neither library generates it, so this is not a ' +
				'difference between them, but it is a real gap in both.'
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
					`the last month, against ${num(ourNpm.downloadsLastMonth)} for ts-pptx. That gap buys real ` +
					'things: answers that already exist, examples written by people who are not the ' +
					'maintainer, and reasonable odds that a bug on a common path was hit by someone else ' +
					'first. Anyone who weighs those above the differences measured below should use pptxgenjs.'
			)
		)
	lines.push('')
	return lines
}

/**
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionCoverage(snapshot) {
	const counts = COLUMNS.map((subject) => snapshot.coverage.filter((row) => row.results[subject] === 'emitted').length)
	const total = snapshot.coverage.length
	const labelOf = (/** @type {string} */ id) => snapshot.coverage.find((row) => row.id === id)?.label ?? id
	const baseline = snapshot.coverage.filter((row) => row.group === 'shared').length

	const lines = [
		'## Construct coverage',
		'',
		...para(`Of ${total} probes, ts-pptx emitted ${counts[0]} and pptxgenjs emitted ${counts[1]}.`),
		...para(
			'The middle column is the token the harness looks for. It is the OOXML element in every ' +
				'case but one, where the intent is speaker notes and the token is the note text itself; ' +
				'the part each token has to appear in is recorded in the snapshot.'
		),
	]

	/** @type {string[]} */
	const groups = []
	for (const row of snapshot.coverage) if (!groups.includes(row.group)) groups.push(row.group)

	for (const group of groups) {
		const rows = snapshot.coverage.filter((row) => row.group === group)
		lines.push(`### ${groupLabel(group)}`, '', '| Intent | Looked for | ts-pptx | pptxgenjs |', '|---|---|---|---|')
		for (const row of rows) {
			const outcomes = COLUMNS.map((subject) => OUTCOME_LABELS[row.results[subject] ?? ''] ?? row.results[subject])
			lines.push(`| ${row.label} | ${code(row.construct)} | ${outcomes[0]} | ${outcomes[1]} |`)
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
			['', 'ts-pptx', 'pptxgenjs'],
			[
				comparedRow(
					'Installed size, with dependencies',
					rows.map((row) => row?.install?.bytes),
					mb
				),
				comparedRow(
					'Installed size, the package alone',
					rows.map((row) => row?.install?.packageBytes),
					mb
				),
				comparedRow(
					'Runtime dependencies, transitive',
					rows.map((row) => row?.dependencies?.transitive),
					num
				),
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
				comparedRow(
					'Hello world, first chunk',
					rows.map((row) => row?.helloWorld?.initialBytes),
					kb
				),
				comparedRow(
					'Hello world, every chunk',
					rows.map((row) => row?.helloWorld?.totalBytes),
					kb
				),
			]
		),
		...para(
			'The hello world program is identical in intent on both sides and written in each ' +
				"library's own idiom: one slide, one text box, then export. It is bundled with esbuild " +
				'for the browser, minified, and gzipped at level 9, following the conventions ' +
				'`scripts/bundle-size-ratchet.mjs` documents, with one difference that matters. The ' +
				'ratchet never bundles, so it cannot drop unreachable code and its figures are an upper ' +
				'bound on what the package ships; this bundles and does tree-shake, because a ' +
				"consumer's build is precisely the thing being compared here. **The two sets of numbers " +
				'will not agree, and neither is wrong.**'
		),
		...para(
			'Two figures rather than one, because code splitting is on. The first chunk is what the ' +
				'program pays to start; every chunk is what it can reach. They differ for ts-pptx ' +
				'because font metrics load `opentype.js` behind a dynamic import that only runs once a ' +
				'font is registered, and a bundler that can defer that will. Charging the program for a ' +
				'chunk it may never fetch, and hiding bytes it might, are both misleading, so both are ' +
				'printed.'
		),
	]
	const ours = hygieneOf(snapshot, OURS)
	const upstream = hygieneOf(snapshot, UPSTREAM)
	if (typeof ours?.install?.bytes === 'number' && ours.install.bytes > upstream?.install?.bytes)
		lines.push(
			...para(
				'ts-pptx installs larger than pptxgenjs despite carrying fewer dependencies. Its ' +
					'`dist/` ships unminified, and a large share of that weight is documentation comments ' +
					'that no consumer build keeps, which is why the bundled figures above are much closer ' +
					'together than the installed ones.'
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
function sectionReadSide() {
	return [
		'## The read side',
		'',
		...para(
			'pptxgenjs generates decks. It does not read them, and it does not claim to. So there is ' +
				'nothing to compare here and no table: this is a capability one library has, which is a ' +
				'different statement from one library being better at something both do.'
		),
		...para('ts-pptx also reads:'),
		...bullet(
			'[Inspection](reference/pptx-inspection.md) reports what a package contains without ' + 'parsing it into a model.'
		),
		...bullet(
			'[Reading](reference/pptx-read.md) loads a deck into an addressable object model, edits ' +
				'it in place, and writes the package back out.'
		),
		...bullet(
			'[Deck to script](reference/pptx-to-script.md) turns an existing deck into runnable ' +
				'TypeScript, reporting what it could not express rather than dropping it silently.'
		),
		'',
		...para('If you only generate decks, none of this is a reason to choose either library.'),
	]
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
 * Project health, kept in its own section for the reason it has its own snapshot key: it
 * measures the projects rather than what they emit, and mixing it into the tables above
 * would let a star count read as a property of the output.
 * @param {Snapshot} snapshot
 * @returns {string[]}
 */
function sectionHealth(snapshot) {
	const rows = COLUMNS.map((subject) => healthOf(snapshot, subject))
	const lines = [
		'## Project health',
		'',
		...para(
			'Separate from everything above, and on purpose. These figures describe how the two ' +
				'projects are run, not what either one emits. Stars and downloads measure adoption, ' +
				'adoption measures history as much as merit, and none of it belongs in the same table as ' +
				'a construct a library does or does not write.'
		),
		...table(
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
					'. The download figure above is their sum (' +
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
				`The pptxgenjs row shows no npm release since ${upstream.npm.lastPublish} and no commit on ` +
					`${code(upstream.defaultBranch)} since ${upstream.lastDefaultBranchCommit}. That is what the two ` +
					'APIs report, and it is all this page says about it: from outside, a stable library ' +
					'that has stopped needing changes looks exactly like one between maintainers, and ' +
					'this measurement cannot tell them apart. It is worth weighing either way, next to ' +
					`${num(upstream.openIssues ?? 0)} open issues and ${num(upstream.openPullRequests ?? 0)} open pull ` +
					'requests.'
			)
		)
	return lines
}

/**
 * @param {Snapshot} snapshot
 * @returns {string}
 */
export function renderPage(snapshot) {
	const lines = [
		...frontmatter(snapshot),
		'',
		...banner('FILE'),
		'',
		...sectionPremise(snapshot),
		...sectionConcessions(snapshot),
		...sectionCoverage(snapshot),
		...sectionValidity(snapshot),
		...sectionHygiene(snapshot),
		...sectionReadSide(),
		...sectionHealth(snapshot),
	]
	return (
		lines
			.join('\n')
			.replace(/\n{3,}/g, '\n\n')
			.trimEnd() + '\n'
	)
}

/**
 * The README summary: the same snapshot, shorter, pointing at the page for the rest.
 * @param {Snapshot} snapshot
 * @returns {string}
 */
export function renderReadmeRegion(snapshot) {
	const total = snapshot.coverage.length
	const counts = COLUMNS.map((subject) => snapshot.coverage.filter((row) => row.results[subject] === 'emitted').length)
	const ourHealth = healthOf(snapshot, OURS)
	const upstreamHealth = healthOf(snapshot, UPSTREAM)
	const ourValidity = validityOf(snapshot, OURS)
	const upstreamValidity = validityOf(snapshot, UPSTREAM)

	const lines = [
		REGION_START,
		...banner('REGION'),
		'',
		'## How This Compares With PptxGenJS',
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
			'The full tables, the method behind them, and what ts-pptx gives up (no CommonJS, no CDN ' +
				'script tag, Node.js `>=24` only) are on the ' +
				'[comparison page](docs/comparison.md).'
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
		[README, spliceRegion(fs.readFileSync(README, 'utf8'), renderReadmeRegion(snapshot))],
	]
	const stale = outputs.filter(([file, text]) => !fs.existsSync(file) || fs.readFileSync(file, 'utf8') !== text)

	if (values.check) {
		if (stale.length === 0) {
			console.log('comparison: docs/comparison.md and README.md match the snapshot.')
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
