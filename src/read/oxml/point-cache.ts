/**
 * The one place a chart point cache is turned into an idx-ordered array.
 *
 * Both chart families cache their category/value data the same way: a run of indexed point
 * elements plus a declared count, which the reader widens back into a dense array so
 * `series.values[3]` means the point at `idx="3"` and a gap reads as `null`. Classic `c:`
 * charts and 2016 `cx:` charts spell it differently — `c:ptCount/@val` and a `c:v` child
 * against `cx:lvl/@ptCount` and the point's own text — and nothing else about the two decodes
 * differs, so the shape parameterizes over exactly those two differences.
 *
 * **Why neither number out of the file is trusted.** `@ptCount` is `xsd:unsignedInt`, so
 * `4294967295` is schema-valid, and `new Array(n).fill(null)` at that size is not a slow path
 * but a process kill: V8 answers `FATAL ERROR: invalid table size` and the host dies with no
 * exception to catch. That was reachable through the public read API — `Presentation.load()`
 * on a deck with an edited `c:ptCount` returned normally and the first `chart.series` access
 * killed the process. A single `<c:pt idx="900000000"/>` does the same thing on its own.
 * Reading a deck is the one part of this library whose input comes from somewhere other than
 * the calling program, so both numbers are bounded here.
 *
 * The array is sized by the points that are really there, not by the declared count. The cost
 * of that is a *trailing* gap: a cache declaring four points but carrying three (the fourth
 * cell blank) reads three long rather than four. No value is lost — the dropped slots are the
 * ones that would have been `null` — and the disagreement is warned about either way.
 */
import { attr, numberValue, type Element } from './dom.js'
import { warn } from '../../diagnostics.js'

/**
 * The highest point index this reader will honour.
 *
 * A cache references a range in the embedded workbook, and a worksheet has 1,048,576 rows, so
 * a point index at or above that cannot describe data any producer could have charted. Using
 * the sheet bound rather than a round number keeps the ceiling something a reader can check:
 * below it every real cache fits, above it the file is wrong by construction. It also caps the
 * worst-case allocation at about 8 MB of `null` — survivable, and reported — instead of an
 * unbounded one that takes the process with it.
 */
const MAX_POINT_INDEX = 1048576

/**
 * Read indexed cache points into a dense, idx-ordered array.
 *
 * The result is as long as the highest `@idx` actually present, and holds `null` at every
 * index no point claimed. A declared count that disagrees is warned about and not obeyed; a
 * point indexed past {@link MAX_POINT_INDEX} is warned about and dropped.
 * @param pts - the point elements, in document order
 * @param declaredCount - the cache's own count attribute, or `null` when absent
 * @param valueOf - reads one point's value; the two chart families differ only here
 * @param label - what to call the count in a diagnostic, e.g. `c:ptCount`
 */
export function readIndexedPoints(
	pts: Element[],
	declaredCount: number | null,
	valueOf: (pt: Element) => string | null,
	label: string
): (string | null)[] {
	// One pass to find the real extent — one past the highest index a point actually claims,
	// ignoring the ones that claim an impossible index. A cache carrying no points has no
	// extent at all, whatever it declares.
	let count = 0
	let dropped = 0
	for (const pt of pts) {
		const idx = numberValue(attr(pt, 'idx')) ?? 0
		if (idx < 0 || idx >= MAX_POINT_INDEX) {
			dropped++
			continue
		}
		if (idx + 1 > count) count = idx + 1
	}
	if (dropped > 0) {
		warn(
			'chart/point-index-out-of-range',
			`${label}: ${dropped} cache point(s) are indexed outside 0..${MAX_POINT_INDEX - 1}, which no worksheet-backed chart can reference; they were dropped`
		)
	}
	if (declaredCount !== null && declaredCount !== count) {
		warn(
			'chart/point-count-mismatch',
			`${label} declares ${declaredCount} point(s) but the cache carries ${pts.length}; reading ${count} rather than allocating the declared size`
		)
	}
	const points: (string | null)[] = new Array<string | null>(count).fill(null)
	for (const pt of pts) {
		const idx = numberValue(attr(pt, 'idx')) ?? 0
		if (idx < 0 || idx >= count) continue
		points[idx] = valueOf(pt)
	}
	return points
}
