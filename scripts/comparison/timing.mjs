/**
 * How long each library takes to turn a deck into bytes.
 *
 * The rest of this directory measures things that do not move: a construct is emitted or it
 * is not, a tarball is the size it is, and a second run on another machine gets the same
 * answer. A clock is not like that. It reports the machine, the Node build, what else was
 * running, and how warm the JIT was, and only somewhere underneath all of that, the
 * library. Everything below is about getting the library out from under the rest.
 *
 * ## Matched compression, and why the whole measurement turns on it
 *
 * A `.pptx` is a zip, and most of the time this measures is spent making one. So the
 * compression setting is not a detail of the benchmark, it *is* the benchmark, and the two
 * libraries do not default to the same one:
 *
 *   - ts-pptx deflates at level 6 unless told not to.
 *   - pptxgenjs passes no compression option to JSZip on the `outputType` path, and JSZip's
 *     own default is STORE. Its `compression` argument is honoured on the STREAM path and
 *     the browser blob path, and silently ignored on the one in between — which is the one
 *     `writeFile` takes in Node.
 *
 * So the obvious measurement — call `write` on both and start a clock — compares deflating
 * against not deflating and reports the difference as a library being slow. Both modes in
 * {@link MODES} are therefore matched pairs, and both go through pptxgenjs's STREAM path so
 * that the mode is the only thing changing on that side.
 *
 * The match is not asserted, it is **evidenced**, and by the right evidence. The obvious
 * check — do the two arms produce the same number of bytes — is the wrong one, because they
 * genuinely do not: the two libraries emit different volumes of XML for the same deck, most
 * visibly for charts, and a matched mode does nothing to change that. What a matched mode
 * predicts is narrower and stronger: that *each* arm's stored output is much larger than its
 * own compressed output. `compression` records that ratio per library per case, and a run
 * where either side stops compressing when asked shows up there as a ratio falling towards
 * 1, rather than as a timing regression nobody can explain.
 *
 * `bytesAgree` is kept beside it as the other half of the picture — how far apart the two
 * decks are in size — because a time is only half a fact when the two libraries are writing
 * outputs of different sizes, and the page has to be able to say so.
 *
 * ## What the protocol does about noise
 *
 * Four things, each aimed at a different way a clock lies:
 *
 *   - **Warm up, and throw the warm-up away.** V8 re-optimises for a long time. Measured
 *     cold, the first arm to run carries the tiering cost for both, and the ratio moves by
 *     tens of percent between repeats of the identical run.
 *   - **Interleave the arms, alternating which goes first.** A laptop that throttles, or an
 *     endpoint scanner that wakes up mid-run, does not respect a benchmark's structure.
 *     Measuring all of one library and then all of the other hands whichever ran during the
 *     quiet half a result it did not earn; alternating spreads any drift across both
 *     columns.
 *   - **Report the median, and carry the spread with it.** The mean is at the mercy of one
 *     descheduled run. The spread is recorded next to it because a difference smaller than
 *     the noise is not a difference, and the page cannot say so unless the number is there.
 *   - **Spend a budget, not a fixed count.** A 500-slide deck and a one-slide deck differ by
 *     two orders of magnitude, and a round count that suits either one is absurd for the
 *     other.
 *
 * What none of that fixes is that the absolute milliseconds belong to the machine that took
 * them. They are recorded with the CPU and the Node version beside them, and the page leads
 * with the ratio, which is the part that travels.
 */
import os from 'node:os'
import { PROGRAMS, programArm, resetProgramData } from './programs.mjs'
import { isUnavailable, unavailable } from './unavailable.mjs'
import { scaleSource, WORKLOADS } from './workloads.mjs'

/**
 * One matched pair of write options, one per library.
 * @typedef {object} Mode
 * @property {string} id - kebab-case, stable; the snapshot's key for this mode
 * @property {string} label - how the mode reads on the page
 * @property {Record<string, object>} props - the write options, per subject
 */

/**
 * The two matched pairs.
 *
 * `deflate` is first because it is what each library does when a consumer asks for a file
 * they intend to keep, and because it is the mode where the output sizes prove the match.
 * `store` is the control: it takes the compressor out of the measurement, leaving the part
 * that is actually each library's own code — building the XML and assembling the package.
 * Between them they say whether a difference is in the deck or in the zip.
 *
 * Upstream goes through `STREAM` in both, so its code path is fixed and the compression
 * flag is the only thing that moves. `nodebuffer` is our equivalent: the same bytes, in the
 * same shape, on the path a Node consumer takes.
 * @type {Mode[]}
 */
export const MODES = [
	{
		id: 'deflate',
		label: 'Compressed',
		props: {
			'ts-pptx': { outputType: 'nodebuffer' },
			pptxgenjs: { outputType: 'STREAM', compression: true },
		},
	},
	{
		id: 'store',
		label: 'Stored',
		props: {
			'ts-pptx': { outputType: 'nodebuffer', compression: false },
			pptxgenjs: { outputType: 'STREAM', compression: false },
		},
	},
]

/** Rounds thrown away before any are kept, and the time cap on doing that. */
const WARMUP_ROUNDS = 12
const WARMUP_BUDGET_MS = 1500
/** Kept rounds: at least the floor, then until the count or the budget runs out. */
const MIN_ROUNDS = 7
const MAX_ROUNDS = 25
const ROUND_BUDGET_MS = 2500

/** @param {bigint} ns @returns {number} */
const toMs = (ns) => Number(ns) / 1e6

/**
 * The value at a quantile of a sample, interpolating between neighbours.
 * @param {number[]} values - need not be sorted
 * @param {number} q - 0 to 1
 * @returns {number}
 */
export function quantile(values, q) {
	if (values.length === 0) throw new Error('cannot take a quantile of an empty sample')
	const sorted = [...values].sort((a, b) => a - b)
	/** @param {number} index @returns {number} */
	const at = (index) => /** @type {number} */ (sorted[index])
	const position = (sorted.length - 1) * q
	const low = Math.floor(position)
	const high = Math.ceil(position)
	return low === high ? at(low) : at(low) + (at(high) - at(low)) * (position - low)
}

/**
 * A sample reduced to what the snapshot keeps.
 *
 * The median is the headline. `spread` is the interquartile range as a percentage of it,
 * which is the number that says whether a difference between two columns means anything:
 * a gap narrower than either column's spread is noise wearing a percentage sign.
 * @param {number[]} values - milliseconds, one per kept round
 * @returns {{median: number, min: number, spread: number, rounds: number}}
 */
export function summarise(values) {
	const median = quantile(values, 0.5)
	return {
		median,
		min: Math.min(...values),
		spread: median === 0 ? 0 : ((quantile(values, 0.75) - quantile(values, 0.25)) / median) * 100,
		rounds: values.length,
	}
}

/**
 * Build a deck and write it, once, timed.
 *
 * The reset is outside the clock and so is constructing the presentation object, because
 * neither is work a consumer's deck pays for twice. What is inside is exactly what a
 * consumer's program does: hand the library its content, then ask for bytes.
 * @param {() => any} construct
 * @param {(pres: any) => unknown} build
 * @param {object} props - the write options for this subject and mode
 * @returns {Promise<{ms: number, bytes: number}>}
 */
async function timeOnce(construct, build, props) {
	// The bundle programs share their corpus values, and pptxgenjs edits what it is handed.
	// See `./corpus-data.mjs`; the timing workloads need no such help, by construction.
	resetProgramData()
	const pres = construct()
	const started = process.hrtime.bigint()
	await build(pres)
	const written = await pres.write(props)
	const ms = toMs(process.hrtime.bigint() - started)
	return { ms, bytes: written?.byteLength ?? written?.length ?? 0 }
}

/**
 * Time one case in one mode, both arms, interleaved.
 *
 * A throw from either arm ends the case rather than the run: a library that cannot build
 * one of these decks is a fact about that library, and the remaining cases still have
 * something to say. The reason travels as an {@link unavailable} marker so it reaches the
 * snapshot gate instead of becoming a silently missing row.
 * @param {{id: string, build: Record<string, (pres: any) => unknown>}} deck
 * @param {Mode} mode - one of {@link MODES}
 * @param {Record<string, () => any>} subjects
 * @param {(deck: any, subject: string) => (pres: any) => unknown} armOf
 * @returns {Promise<Record<string, unknown> | import('./unavailable.mjs').Unavailable>}
 */
async function timeCase(deck, mode, subjects, armOf) {
	/** @type {Array<{subject: string, run: () => Promise<{ms: number, bytes: number}>, samples: number[], bytes: number}>} */
	const arms = []
	for (const [subject, construct] of Object.entries(subjects)) {
		const props = mode.props[subject]
		if (!props) return unavailable('no ' + mode.id + ' write options registered for ' + subject)
		const build = armOf(deck, subject)
		arms.push({ subject, run: () => timeOnce(construct, build, props), samples: [], bytes: 0 })
	}

	try {
		const warmupStart = process.hrtime.bigint()
		for (let round = 0; round < WARMUP_ROUNDS; round++) {
			for (const arm of arms) await arm.run()
			if (toMs(process.hrtime.bigint() - warmupStart) > WARMUP_BUDGET_MS) break
		}

		const start = process.hrtime.bigint()
		for (let round = 0; round < MAX_ROUNDS; round++) {
			// Alternate the order so neither column always runs into the other's leftovers.
			const order = round % 2 === 0 ? arms : [...arms].reverse()
			for (const arm of order) {
				const { ms, bytes } = await arm.run()
				arm.samples.push(ms)
				arm.bytes = bytes
			}
			if (round + 1 >= MIN_ROUNDS && toMs(process.hrtime.bigint() - start) > ROUND_BUDGET_MS) break
		}
	} catch (error) {
		return unavailable(
			'the ' +
				deck.id +
				' deck could not be timed in ' +
				mode.id +
				' mode: ' +
				(error instanceof Error ? error.message : String(error))
		)
	}

	/** @type {Record<string, unknown>} */
	const result = {}
	for (const arm of arms) result[arm.subject] = { ...summarise(arm.samples), bytes: arm.bytes }
	return { ...result, bytesAgree: bytesAgreement(arms.map((arm) => arm.bytes)) }
}

/**
 * How far apart the two arms' output sizes are, as a percentage of the larger.
 *
 * Not a match check — see the header: the two libraries emit different amounts of XML for
 * the same deck, and this number says how much. It belongs in the snapshot because a
 * generation time means less on its own than it does beside the size of what was generated.
 * `null` where there is nothing to compare.
 * @param {number[]} sizes
 * @returns {number | null}
 */
export function bytesAgreement(sizes) {
	if (sizes.length !== 2) return null
	const [ours, theirs] = sizes
	if (typeof ours !== 'number' || typeof theirs !== 'number' || ours <= 0 || theirs <= 0) return null
	return (Math.abs(ours - theirs) / Math.max(ours, theirs)) * 100
}

/**
 * How much smaller each library's compressed deck is than its own stored one.
 *
 * The evidence that the matched modes matched, and the only form of it that survives the
 * two libraries emitting different XML: whatever else differs, a library asked to compress
 * has to produce something substantially smaller than the same library asked not to. A
 * ratio near 1 means that side ignored the mode and its "compressed" row is timing a
 * different operation from the column beside it.
 * @param {Record<string, unknown>} modes - the finished per-mode results for one case
 * @param {string[]} subjects
 * @returns {Record<string, number | null>}
 */
function compressionRatios(modes, subjects) {
	/** @param {string} mode @param {string} subject @returns {number | undefined} */
	const bytesOf = (mode, subject) => {
		const measured = /** @type {any} */ (modes[mode])
		const arm = measured && !isUnavailable(measured) ? measured[subject] : undefined
		return typeof arm?.bytes === 'number' ? arm.bytes : undefined
	}
	/** @type {Record<string, number | null>} */
	const ratios = {}
	for (const subject of subjects) {
		const stored = bytesOf('store', subject)
		const deflated = bytesOf('deflate', subject)
		ratios[subject] =
			typeof stored === 'number' && typeof deflated === 'number' && deflated > 0 ? stored / deflated : null
	}
	return ratios
}

/**
 * What the numbers below belong to.
 *
 * Recorded because milliseconds without a machine are a measurement of nowhere. A reader
 * who wants to know whether our figures apply to them needs to see the CPU they were taken
 * on, and a reader comparing two snapshots needs to know whether the machine changed under
 * them.
 * @returns {Record<string, unknown>}
 */
function machine() {
	const cpus = os.cpus()
	return {
		cpu: cpus[0]?.model?.trim() ?? 'unknown',
		cores: cpus.length,
		platform: process.platform,
		node: process.version,
		measured: new Date().toISOString().slice(0, 10),
	}
}

/**
 * Time every deck in both corpora, in both modes.
 *
 * The bundle programs and the scale workloads run through one loop because they are the
 * same kind of thing measured at two ends of a range, and the page prints them as two
 * tables only because a reader reads milliseconds and hundreds of milliseconds differently.
 * @param {object} opts
 * @param {Record<string, () => any>} opts.subjects - a fresh presentation per library
 * @returns {Promise<Record<string, unknown>>}
 */
export async function measureTiming({ subjects }) {
	/** @type {Array<{group: string, decks: any[], armOf: (deck: any, subject: string) => (pres: any) => unknown}>} */
	const groups = [
		{ group: 'programs', decks: PROGRAMS, armOf: programArm },
		{
			group: 'scale',
			decks: WORKLOADS,
			armOf: (deck, subject) => {
				const build = deck.build[subject]
				if (!build) throw new Error('the timing workload "' + deck.id + '" has no ' + subject + ' arm')
				return build
			},
		},
	]

	/** @type {any[]} */
	const cases = []
	for (const { group, decks, armOf } of groups)
		for (const deck of decks) {
			/** @type {Record<string, unknown>} */
			const modes = {}
			for (const mode of MODES) modes[mode.id] = await timeCase(deck, mode, subjects, armOf)
			cases.push({
				id: deck.id,
				group,
				label: deck.label,
				what: deck.what,
				...(typeof deck.slides === 'number' ? { slides: deck.slides } : {}),
				modes,
				compression: compressionRatios(modes, Object.keys(subjects)),
			})
		}

	return {
		machine: machine(),
		modes: MODES.map((mode) => ({ id: mode.id, label: mode.label })),
		// The scale decks are one shape at three sizes, so the shape is carried once. The five
		// programs beside them need nothing here: they are the bundle corpus, already printed on
		// the syntax page, and a second copy would be a second thing to keep true.
		shape: {
			sizes: WORKLOADS.map((workload) => workload.slides),
			source: Object.fromEntries(Object.keys(subjects).map((subject) => [subject, scaleSource(subject)])),
		},
		cases,
	}
}
