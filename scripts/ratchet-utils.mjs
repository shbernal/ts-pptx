/**
 * Budget mechanics the two size gates share: how much room a freeze leaves, when a measurement is
 * over its budget or far enough under to bank, and how a budget file is read, written and checked
 * against what the gate measured.
 *
 * `bundle-size-ratchet.mjs` and `bundle-tier-size.mjs` run the same ratchet over different
 * measurements. The tier gate factored its half of this out and tested it; the size gate kept
 * inline copies nothing tested, and the two had drifted in what they left unchecked.
 * `raw-xml-ratchet.mjs` fails on a budget entry that no longer matches anything, and neither size
 * gate did, so removing an entry or a tier from a gate left a budget row that nothing reported.
 * Two gates that disagree about what counts as ordinary work would also teach a reader that the
 * numbers are arbitrary.
 */

import fs from 'node:fs'

/** Room a freeze leaves above the measurement, so ordinary work is not a re-freeze. */
export const HEADROOM_PCT = 5

/** Re-freeze is only worth asking for when a measurement comes in this far under budget. */
export const SLACK_PCT = 15

/**
 * ...and this far under in absolute terms, which is what keeps the small entries usable.
 *
 * `--freeze` rounds the budget up to a whole KiB, and on a 5 KiB entry that rounding alone is
 * larger than {@link SLACK_PCT} of it: `zip.js` froze at 6 KiB, measured 5 KiB, and was
 * immediately 16% under — a nag no re-freeze could clear, because the next freeze rounds to the
 * same 6 KiB. A percentage of a tiny number is noise; asking for a re-freeze over 1 KiB is asking
 * for a gate to be switched off.
 */
export const SLACK_MIN_BYTES = 2048

/**
 * A byte count as the gates print it, in kibibytes: the figure is divided by 1024, so it carries
 * the binary unit rather than `kB`.
 * @param {number} bytes
 * @returns {string}
 */
export function kb(bytes) {
	return (bytes / 1024).toFixed(1) + ' KiB'
}

/**
 * Compare one measurement against one budget, on the ratchet's terms.
 *
 * The whole of the verdict logic, and it needs no build to exercise — see
 * `test/scripts/ratchet-utils.test.js`.
 * @param {number} bytes - what was measured
 * @param {number} budget - what is frozen
 * @returns {'over' | 'under' | 'ok'} `under` means far enough under to be worth banking
 */
export function verdictFor(bytes, budget) {
	if (bytes > budget) return 'over'
	if (bytes < budget * (1 - SLACK_PCT / 100) && budget - bytes >= SLACK_MIN_BYTES) return 'under'
	return 'ok'
}

/**
 * The budget a measurement freezes to: {@link HEADROOM_PCT} above it, rounded up to a whole
 * KiB so the file reads as a decision someone made rather than a build artifact copied in.
 * @param {number} bytes
 * @returns {number}
 */
export function frozenBudget(bytes) {
	return Math.ceil((bytes * (1 + HEADROOM_PCT / 100)) / 1024) * 1024
}

/**
 * A budget file's contents. The shape is the gate's own: a number per entry, or figures per tier.
 * @param {string} file - absolute path of the budget file
 * @returns {Record<string, unknown>}
 */
export function readBudget(file) {
	return JSON.parse(fs.readFileSync(file, 'utf8'))
}

/**
 * Write a budget file as every freeze writes one: tab-indented JSON and a trailing newline, so a
 * re-freeze diffs as the numbers that moved.
 * @param {string} file - absolute path of the budget file
 * @param {Record<string, unknown>} budget - the budget to write
 */
export function writeBudget(file, budget) {
	fs.writeFileSync(file, JSON.stringify(budget, null, '\t') + '\n')
}

/**
 * Where a budget file and a measurement disagree about which keys exist.
 *
 * `missing` are measured keys without a complete budget, which a gate cannot judge. `stale` are
 * budgeted keys the gate no longer measures: an entry or tier removed from the gate, whose row
 * would otherwise sit in the file with nothing reporting it.
 * @template V
 * @param {Iterable<string>} measuredKeys - what the gate measured
 * @param {Record<string, V>} budget - the budget file's contents
 * @param {(value: V | undefined) => boolean} [isBudgeted] - whether a key's value is a complete
 *   budget; a number, unless the gate says otherwise
 * @returns {{ missing: string[], stale: string[] }}
 */
export function budgetKeyDrift(measuredKeys, budget, isBudgeted = (value) => typeof value === 'number') {
	const measured = [...measuredKeys]
	return {
		missing: measured.filter((key) => !isBudgeted(budget[key])),
		stale: Object.keys(budget).filter((key) => !measured.includes(key)),
	}
}
