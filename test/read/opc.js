// Relationship-graph helpers over a loaded package's `presentation.opc`.
//
// Both of these were re-derived in four to seven read tests each, with the usual drift: three
// spellings of "the single rel of this type" (`.find`, `.filter()[0]`, and a `length === 0`
// early return) that all mean the same thing. The behaviour they assert is a property of the
// package, not of any one suite, so it belongs in one place.
//
// Not a test file (no `.test.` in the name) — vitest's default glob skips it.

import { assert } from '../helpers.js'

/**
 * The resolved target part name of `partName`'s single relationship of `type`, or null when it
 * has none. Takes the first if there are several — every caller asks about a relationship the
 * format allows at most one of.
 */
export function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const match = [...rels].find((rel) => rel.type === type)
	return match ? rels.resolveTarget(match.id) : null
}

/**
 * Every internal relationship in the package resolves to a part that exists.
 *
 * This is the check that catches the whole class of splice bugs where a part is removed or
 * renamed and something still points at where it used to be. PowerPoint reports the result as
 * a repair prompt naming nothing, and schema validation does not see it at all: each part is
 * individually valid, and it is the graph between them that is broken.
 */
export function assertNoDanglingRels(opc) {
	for (const partName of opc.parts.keys()) {
		if (partName.endsWith('.rels')) continue
		for (const rel of opc.relationshipsFor(partName)) {
			if (rel.targetMode === 'External') continue
			const target = opc.relationshipsFor(partName).resolveTarget(rel.id)
			assert(opc.part(target), `${partName} → ${rel.id} targets an existing part (${target})`)
		}
	}
}
