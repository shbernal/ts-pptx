/**
 * ts-pptx: option-bag normalization helpers — one spelling of absent.
 *
 * `exactOptionalPropertyTypes` draws a line this library needs drawn: a `foo?: T` declaration says
 * the key is either *missing* or holds a `T`, and a key present with an `undefined` in it is a
 * third state. Readers here cannot see that state — every option is read with truthiness or `?.`
 * — but spreads can, and the generator spreads these bags constantly: a layout placeholder's
 * options onto a slide's, a column default under a cell's own, a combo subchart's overrides onto
 * the chart's. There the side that wins is decided by whether the key *exists*, not by what it
 * holds, so "the caller said nothing" and "the caller said nothing, in writing" produce different
 * output.
 *
 * These two helpers are how a normalizer stays on the right side of that. They are the write-side
 * twin of `compact()` (`script/from-read/values.ts`), which keeps the same invariant on the read
 * side and for the same reason.
 *
 * Internal: nothing here is exported from an entrypoint, the same way `constants-internal.ts` and
 * `units-internal.ts` are not.
 */

import type { OptionalKeysOf } from './types/internal.js'

/**
 * Write a normalized value onto an option bag, spelling "no value" as an *absent* key.
 *
 * Reach for this wherever a normalizer's result may be nothing —
 * `setOrClear(opts, 'holeSize', clampChartPct(opts.holeSize, 10, 90, 'holeSize'))` — rather than
 * assigning the result directly, which leaves the caller's rejected value replaced by an
 * `undefined` instead of removed.
 *
 * Only optional keys are accepted, so this cannot be used to unset something required.
 * @param bag - the options bag to write onto
 * @param key - the key to set or remove
 * @param value - the normalized value, or `undefined` to remove the key
 */
export function setOrClear<T extends object, K extends OptionalKeysOf<T> & keyof T>(bag: T, key: K, value: T[K]): void {
	if (value === undefined) delete bag[key]
	else bag[key] = value
}

/**
 * The named properties of `source` that are actually set, as a new object.
 *
 * For the other shape this problem takes: a literal that projects a fixed key list off a bag that
 * may not state them all (`{ x: opts.x, y: opts.y, rotate: opts.rotate, … }`), which writes a key
 * for every name whether or not the source had one. Naming the keys once is also shorter than
 * spelling out the copy, and it cannot drift from itself the way a hand-written list can.
 *
 * Every key comes back optional and with `undefined` excluded from its type, which is the whole
 * claim: what this returns is the subset that was actually stated.
 * @param source - the bag to read from
 * @param keys - the keys to carry over
 */
export function pickDefined<T extends object, K extends keyof T>(
	source: T,
	keys: readonly K[]
): { [P in K]?: Exclude<T[P], undefined> } {
	const out: { [P in K]?: Exclude<T[P], undefined> } = {}
	for (const key of keys) {
		const value = source[key]
		if (value !== undefined) out[key] = value as Exclude<T[K], undefined>
	}
	return out
}
