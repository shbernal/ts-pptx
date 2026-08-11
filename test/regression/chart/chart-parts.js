// Locating the chart part inside a generated package.
//
// `ppt/charts/chartN.xml` is numbered by emission order, so no chart test can hard-code its
// own path; sixteen of them each re-derived the same two lookups instead. Both assert rather
// than return null, because a chart test whose chart part is missing has already failed and
// should say so there, not three assertions later on an `undefined` XML string.
//
// Not a test file (no `.test.` in the name) — vitest's default glob skips it.

import { assert, listEntries, readEntry } from '../../helpers.js'

/** The XML of the package's first `ppt/charts/chartN.xml`. */
export function chartXml(zip) {
	const entry = listEntries(zip).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
	assert(entry, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, entry)
}

/** The part name of the package's first `ppt/charts/chartExN.xml` — the extended-chart family. */
export function chartExPath(zip) {
	const entry = listEntries(zip).find((name) => /^ppt\/charts\/chartEx\d+\.xml$/.test(name))
	assert(entry, 'expected a ppt/charts/chartExN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return entry
}

/** The XML of the package's first `ppt/charts/chartExN.xml`. */
export function chartExXml(zip) {
	return readEntry(zip, chartExPath(zip))
}
