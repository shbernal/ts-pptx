/**
 * The comparison page's charts, as data: what each component draws, with no DOM and no Vue in it.
 *
 * `comparison.data.ts` runs {@link shapeComparison} at build time over
 * `scripts/comparison/snapshot.json`, so the page ships the handful of fields the charts draw
 * rather than the snapshot's recorded sources. The components are markup around the result, and
 * `test/regression/www/comparison-charts.test.js` covers the arithmetic here against the
 * committed snapshot.
 */

/** The two libraries, in the column order every comparison table uses. */
export const SUBJECTS = ['ts-pptx', 'pptxgenjs'] as const
export type Subject = (typeof SUBJECTS)[number]

/** The four outcomes a probe can have, in the order the legend lists them. */
export const OUTCOMES = ['emitted', 'no-api', 'absent', 'error'] as const
export type Outcome = (typeof OUTCOMES)[number]

const OUTCOME_LABELS: Record<Outcome, string> = {
	emitted: 'Emitted',
	'no-api': 'No API',
	absent: 'API, but not emitted',
	error: 'Build threw',
}

/** The accessible name of a mark, and the legend text beside it. */
export function outcomeLabel(outcome: Outcome | null): string {
	return outcome === null ? 'Not measured' : OUTCOME_LABELS[outcome]
}

interface CoverageInput {
	id: string
	label: string
	group: string
	results: Record<string, string | undefined>
}

interface ModeMeasure {
	[subject: string]: { median?: unknown } | undefined
}

interface TimingCaseInput {
	id: string
	label: string
	modes?: Record<string, ModeMeasure | undefined>
}

/** The parts of a snapshot the charts read. Everything else in it is ignored. */
export interface SnapshotInput {
	generatedAt: string
	coverage: CoverageInput[]
	validity?: Record<string, unknown>
	timing?: { cases?: TimingCaseInput[] }
}

export interface MatrixRow {
	id: string
	label: string
	outcomes: Record<Subject, Outcome | null>
}

export interface MatrixGroup {
	id: string
	label: string
	rows: MatrixRow[]
	emitted: Record<Subject, number>
}

function asOutcome(value: string | undefined): Outcome | null {
	return (OUTCOMES as readonly string[]).includes(value ?? '') ? (value as Outcome) : null
}

/**
 * The coverage rows grouped as the tables group them, in snapshot order.
 *
 * `labels` is `scripts/comparison/groups.json`, the same file `render.mjs` titles its tables from,
 * with the same fallback for a group it has no label for.
 */
export function coverageGroups(
	coverage: readonly CoverageInput[],
	labels: Readonly<Record<string, string>>
): MatrixGroup[] {
	const groups: MatrixGroup[] = []
	for (const row of coverage) {
		let group = groups.find((candidate) => candidate.id === row.group)
		if (!group) {
			group = {
				id: row.group,
				label: labels[row.group] ?? row.group.charAt(0).toUpperCase() + row.group.slice(1),
				rows: [],
				emitted: { 'ts-pptx': 0, pptxgenjs: 0 },
			}
			groups.push(group)
		}
		const outcomes = { 'ts-pptx': asOutcome(row.results['ts-pptx']), pptxgenjs: asOutcome(row.results['pptxgenjs']) }
		group.rows.push({ id: row.id, label: row.label, outcomes })
		for (const subject of SUBJECTS) if (outcomes[subject] === 'emitted') group.emitted[subject] += 1
	}
	return groups
}

export interface ValidityBar {
	subject: Subject
	/** Decks that validated with no error. */
	clean: number
	/** Decks built that carry at least one schema error. */
	withErrors: number
	/** Intents with no deck at all. */
	notBuilt: number
	/** Every intent in the corpus: the bar's full width. */
	total: number
}

function isSubjectValidity(value: unknown): value is { decks: number; cleanDecks: number } {
	const candidate = value as { decks?: unknown; cleanDecks?: unknown } | null
	return typeof candidate?.decks === 'number' && typeof candidate.cleanDecks === 'number'
}

/**
 * One bar per library over every intent, so a library that built fewer decks shows the gap
 * rather than a shorter bar that reads as a better score.
 */
export function validityBars(validity: SnapshotInput['validity'], total: number): ValidityBar[] {
	return SUBJECTS.flatMap((subject) => {
		const measured = validity?.[subject]
		if (!isSubjectValidity(measured)) return []
		return [
			{
				subject,
				clean: measured.cleanDecks,
				withErrors: measured.decks - measured.cleanDecks,
				notBuilt: total - measured.decks,
				total,
			},
		]
	})
}

export interface ValiditySegment {
	key: 'clean' | 'errors' | 'none'
	count: number
	label: string
}

/** A bar's segments, left to right, leaving out the empty ones. */
export function validitySegments(bar: ValidityBar): ValiditySegment[] {
	const segments: ValiditySegment[] = [
		{ key: 'clean', count: bar.clean, label: 'with no schema error' },
		{ key: 'errors', count: bar.withErrors, label: 'with schema errors' },
		{ key: 'none', count: bar.notBuilt, label: 'with no deck to validate' },
	]
	return segments.filter((segment) => segment.count > 0)
}

/** The sentence a bar stands for, for its accessible name and the line under it. */
export function validitySummary(bar: ValidityBar): string {
	return (
		`${bar.subject}: ${bar.clean} ${bar.clean === 1 ? 'deck' : 'decks'} with no schema error, ` +
		`${bar.withErrors} with errors, and ${bar.notBuilt} of ${bar.total} intents with no deck`
	)
}

export interface RatioRow {
	id: string
	label: string
	/** ts-pptx's median over pptxgenjs's, both asked to compress. */
	compressed: number
	/** The same, both asked to store. */
	stored: number
}

function ratioOf(measured: ModeMeasure | undefined): number | null {
	const ours = measured?.['ts-pptx']?.median
	const theirs = measured?.['pptxgenjs']?.median
	return typeof ours === 'number' && typeof theirs === 'number' && theirs > 0 ? ours / theirs : null
}

/** Every timing case measured in both modes, as two ratios. A case missing either is left out. */
export function timingRatios(cases: readonly TimingCaseInput[]): RatioRow[] {
	return cases.flatMap((row) => {
		const compressed = ratioOf(row.modes?.['deflate'])
		const stored = ratioOf(row.modes?.['store'])
		return compressed === null || stored === null ? [] : [{ id: row.id, label: row.label, compressed, stored }]
	})
}

export interface RatioDomain {
	min: number
	max: number
}

/**
 * A log scale symmetric around 1×, so "half the time" sits as far left of the centre as "twice
 * the time" sits right of it. It reaches at least 0.5× to 2×, and 15% past the furthest ratio, so
 * no dot lands on the edge.
 */
export function ratioDomain(rows: readonly RatioRow[]): RatioDomain {
	const furthest = Math.max(
		Math.log(2),
		...rows.flatMap((row) => [Math.abs(Math.log(row.compressed)), Math.abs(Math.log(row.stored))])
	)
	const reach = furthest * 1.15
	return { min: Math.exp(-reach), max: Math.exp(reach) }
}

/** Ticks come in reciprocal pairs, so each one left of 1× has its mirror on the right. */
const TICK_PAIRS: ReadonlyArray<readonly [number, number]> = [
	[2 / 3, 1.5],
	[0.5, 2],
	[0.25, 4],
	[0.2, 5],
	[0.1, 10],
]

/** 1× and the two outermost reciprocal pairs that fit the domain, in ascending order. */
export function ratioTicks(domain: RatioDomain): number[] {
	const kept = TICK_PAIRS.filter(([low, high]) => low >= domain.min && high <= domain.max).slice(-2)
	return [...kept.map(([low]) => low), 1, ...kept.map(([, high]) => high)].sort((a, b) => a - b)
}

/** Where a ratio sits across the plot, as a percentage of its width. */
export function ratioPosition(value: number, domain: RatioDomain): number {
	const low = Math.log(domain.min)
	return ((Math.log(value) - low) / (Math.log(domain.max) - low)) * 100
}

/** The thin line joining a row's two dots. */
export function ratioSpan(row: RatioRow, domain: RatioDomain): { left: number; width: number } {
	const a = ratioPosition(row.compressed, domain)
	const b = ratioPosition(row.stored, domain)
	return { left: Math.min(a, b), width: Math.abs(a - b) }
}

/** A measured ratio, as the comparison page's table prints it. */
export function formatRatio(value: number): string {
	return (value < 10 ? value.toFixed(2) : value.toFixed(1)) + '×'
}

/** A tick label, with no trailing zeros: `0.5×`, `1×`, `1.5×`. */
export function formatTick(value: number): string {
	return Number(value.toFixed(2)).toString() + '×'
}

/** The sentence a timing row stands for. */
export function ratioSummary(row: RatioRow): string {
	return `${row.label}: ts-pptx takes ${formatRatio(row.compressed)} pptxgenjs's time compressed, ${formatRatio(row.stored)} stored`
}

export interface ComparisonData {
	generatedAt: string
	total: number
	groups: MatrixGroup[]
	validity: ValidityBar[]
	ratios: RatioRow[]
	domain: RatioDomain
	ticks: number[]
}

/** Everything the three charts draw, from one snapshot. */
export function shapeComparison(snapshot: SnapshotInput, labels: Readonly<Record<string, string>>): ComparisonData {
	const ratios = timingRatios(snapshot.timing?.cases ?? [])
	const domain = ratioDomain(ratios)
	return {
		generatedAt: snapshot.generatedAt,
		total: snapshot.coverage.length,
		groups: coverageGroups(snapshot.coverage, labels),
		validity: validityBars(snapshot.validity, snapshot.coverage.length),
		ratios,
		domain,
		ticks: ratioTicks(domain),
	}
}
