/**
 * Re-express one coverage lane in another's instrumentation maps.
 *
 * The merged coverage gate needs the browser lane's hits inside the Node report's own
 * `statementMap`/`fnMap`/`branchMap`, because merging by *location key* would let two
 * spellings of the same open-ended range count as two statements and quietly inflate the
 * denominator. Projecting instead can only move hit counts: the denominator is the Node
 * report's, whatever the browser measured.
 *
 * Lifted out of `coverage-merge.mjs` so it can be exercised directly. Its failure mode is the
 * one a gate cannot show you — a projection that silently loses hits reports *lower* coverage
 * and looks like a test that stopped running, while one that mismatches a location reports a
 * pass on a line nothing executed.
 */

/**
 * Both sides have been through JSON by the time this runs, so an open-ended location's end
 * column is `null` on both — but the normalization is kept anyway, because a comparison
 * that reports a false disagreement is worse than useless: it would condemn a merge that
 * is in fact exact. See the round-trip in `convertBrowserCoverage` for where `Infinity`
 * becomes `null` and why it has to.
 */
/** @param {import('istanbul-lib-coverage').Location} p */
export const pos = (p) => `${p.line}:${Number.isFinite(p.column) ? p.column : -1}`
/** @param {import('istanbul-lib-coverage').Range} loc */
export const locKey = (loc) => `${pos(loc.start)}-${pos(loc.end)}`

/**
 * Every hit the browser lane recorded for one file, keyed by source location.
 *
 * Counts are summed rather than replaced: the same location can appear more than once
 * across the lane's eight scenarios, and what matters downstream is only whether the
 * total is zero.
 * @param {import('istanbul-lib-coverage').FileCoverageData} data
 * @returns {Map<string, number>}
 */
export function hitsByLocation(data) {
	/** @type {Map<string, number>} */
	const hits = new Map()
	/**
	 * @param {string} key
	 * @param {number} count
	 */
	const add = (key, count) => hits.set(key, (hits.get(key) ?? 0) + count)

	for (const [index, loc] of Object.entries(data.statementMap)) add('s:' + locKey(loc), data.s[index] ?? 0)
	for (const [index, fn] of Object.entries(data.fnMap)) add('f:' + locKey(fn.loc), data.f[index] ?? 0)
	for (const [index, branch] of Object.entries(data.branchMap)) {
		branch.locations.forEach(
			/**
			 * @param {import('istanbul-lib-coverage').Range} loc
			 * @param {number} arm
			 */
			(loc, arm) => add('b:' + locKey(loc), data.b[index]?.[arm] ?? 0)
		)
	}
	return hits
}

/**
 * Re-express the browser lane's coverage of one file in the Node report's own maps.
 *
 * The returned file coverage has the Node report's exact `statementMap`, `fnMap` and
 * `branchMap` — so merging it can only move hit counts, never the denominator — with each
 * count taken from whatever the browser recorded at that source location, or zero if it
 * recorded nothing there.
 * @param {import('istanbul-lib-coverage').FileCoverageData} nodeData
 * @param {import('istanbul-lib-coverage').FileCoverageData} browserData
 */
export function project(nodeData, browserData) {
	const hits = hitsByLocation(browserData)
	/** @type {Set<string>} */
	const used = new Set()
	/**
	 * @param {string} key
	 * @returns {number}
	 */
	const take = (key) => {
		const hit = hits.get(key)
		if (hit === undefined) return 0
		used.add(key)
		return hit
	}

	/** @type {Record<string, number>} */
	const s = {}
	for (const [index, loc] of Object.entries(nodeData.statementMap)) s[index] = take('s:' + locKey(loc))

	/** @type {Record<string, number>} */
	const f = {}
	for (const [index, fn] of Object.entries(nodeData.fnMap)) f[index] = take('f:' + locKey(fn.loc))

	/** @type {Record<string, number[]>} */
	const b = {}
	for (const [index, branch] of Object.entries(nodeData.branchMap)) {
		b[index] = branch.locations.map(
			/** @param {import('istanbul-lib-coverage').Range} loc */ (loc) => take('b:' + locKey(loc))
		)
	}

	return {
		coverage: {
			path: nodeData.path,
			statementMap: nodeData.statementMap,
			fnMap: nodeData.fnMap,
			branchMap: nodeData.branchMap,
			s,
			f,
			b,
		},
		// Locations the browser lane measured that the Node report has no slot for. Their
		// hits are dropped; see the header for why that is the safe direction.
		orphans: hits.size - used.size,
		measured: hits.size,
	}
}
