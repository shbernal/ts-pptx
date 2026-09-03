/**
 * ts-pptx: pairing two lists of the same things for comparison.
 *
 * The round trip compares a deck against the deck its script produced, and both the slide
 * calls and the chrome masters have to be paired up before anything can be diffed. Position
 * alone is wrong the moment one item drops out: every later one shifts by one, and the report
 * fills with mismatches that are really one missing shape. So a stable key — a shape's
 * `objectName`, a master's title — aligns what it can, with position as the fallback.
 *
 * Both sides wrote this out, down to two near-identical `nextUnclaimed` closures, and they
 * differed in one respect that matters. `diffChrome`'s own reasoning is recorded there: a
 * mutation that removed layout-title deduplication left two layouts sharing a title and the
 * round trip came back clean, so its key has to identify exactly one item *on both sides*.
 * The call aligner checked the actual side only, and that reasoning was never carried back.
 * This is the both-sides rule, once.
 */

/**
 * Pair `expected` with `actual`, keyed where the key is unambiguous and positional elsewhere.
 *
 * The result holds one entry per pairing, in expected order, then one `[null, item]` for each
 * unclaimed actual item. A `[item, null]` means the item was lost; a `[null, item]` means it
 * was added.
 *
 * @param expected - the source deck's items, in document order
 * @param actual - the round-tripped deck's items, in document order
 * @param keyOf - the alignment key, or `null` for an item that has none
 */
export function alignByKey<T>(
	expected: readonly T[],
	actual: readonly T[],
	keyOf: (item: T) => string | null
): Array<[T | null, T | null]> {
	// A key is usable only where it names exactly one item. Counted on each side separately,
	// and required on both: a key that is unique among the expected items but duplicated among
	// the actual ones would pair the first of the duplicates and let the second pass unexamined.
	const uniqueBy = (items: readonly T[]): Map<string, T> => {
		const counts = new Map<string, number>()
		for (const item of items) {
			const key = keyOf(item)
			if (key !== null) counts.set(key, (counts.get(key) ?? 0) + 1)
		}
		const out = new Map<string, T>()
		for (const item of items) {
			const key = keyOf(item)
			if (key !== null && counts.get(key) === 1) out.set(key, item)
		}
		return out
	}

	const byKey = uniqueBy(actual)
	const expectedKeys = uniqueBy(expected)
	const claimed = new Set<T>()

	// Every actual item some expected item will claim BY KEY, reserved before the positional walk
	// begins. Without the reservation the walk claims whatever it lands on, including an item a
	// later expected entry is going to match by name: with `expected = [<unnamed>, "A"]` against
	// `actual = ["A", <unnamed>]`, the unnamed expected item took `"A"` positionally, `"A"` then
	// found its own match already claimed and fell to the unnamed one, and the report described
	// two shapes that each had an exact counterpart as a full set of differences.
	const reserved = new Set<T>()
	for (const before of expected) {
		const key = keyOf(before)
		if (key === null || !expectedKeys.has(key)) continue
		const named = byKey.get(key)
		if (named !== undefined) reserved.add(named)
	}

	// Positional cursor for items no key can align, advanced only past claimed and reserved items
	// so a keyed match does not lose its slot to an unkeyed one.
	let cursor = 0
	const nextUnclaimed = (): T | null => {
		while (cursor < actual.length) {
			const candidate = actual[cursor++]
			if (candidate !== undefined && !claimed.has(candidate) && !reserved.has(candidate)) return candidate
		}
		return null
	}

	const pairs: Array<[T | null, T | null]> = []
	for (const before of expected) {
		const key = keyOf(before)
		const named = key !== null && expectedKeys.has(key) ? byKey.get(key) : undefined
		const after = named !== undefined && !claimed.has(named) ? named : nextUnclaimed()
		if (after !== null && after !== undefined) claimed.add(after)
		pairs.push([before, after ?? null])
	}
	for (const after of actual) if (!claimed.has(after)) pairs.push([null, after])
	return pairs
}
