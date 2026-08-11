// The read-side fixture corpus: where the decks are, how many there should be, and what
// each one converts to.
//
// Two jobs, and the second is the reason this is a module rather than four copies.
//
// **One definition of the corpus.** Twenty-six read tests each re-derived
// `path.join(__dirname, 'fixtures', ...)`, and four more each re-derived the `.pptx` glob
// that enumerates it. A glob is a claim about how many decks are under test, and a claim
// spelled four times is one that can quietly become false in one of them.
//
// **A floor under that claim.** `scripts/script-roundtrip.mjs` refuses to run against an
// empty corpus; the suites that enumerate it had no such guard, so a `FIXTURES` that
// resolved somewhere else — or a filter that stopped matching — would have turned every
// corpus invariant into a test that iterated nothing and passed. That failure mode is
// indistinguishable from success in a reporter, which is precisely why it is checked here
// once, loudly, at collection time.
//
// Not a test file (no `.test.` in the name) — vitest's default glob skips it.

import { readFile, readdir } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures')

/** Committed printer output, so codegen churn shows up in review rather than only as a pass. */
export const SNAPSHOTS = path.resolve(FIXTURES, '..', 'snapshots')

/** The package root, which a few tests read `package.json` out of. */
export const REPO = path.resolve(FIXTURES, '..', '..', '..')

/**
 * Where the script-running suites lay a printed script down before executing it.
 *
 * Inside the repo rather than the OS temp directory, and required to be: an emitted script
 * imports this package by its published name, which Node resolves by the self-reference rule
 * only from a path underneath the package root. `/.tmp/` is gitignored.
 */
export const SCRATCH = path.join(REPO, '.tmp')

/**
 * Resolve a fixture by name. `.pptx` is implied, so `fixturePath('mixed')` and
 * `fixturePath('mixed.pptx')` are the same file and both spellings already in the suite
 * keep working; anything carrying its own extension (`template.potx`,
 * `slide-transition.oracle.json`) is taken as written.
 *
 * @param {string} name
 */
export function fixturePath(name) {
	return path.join(FIXTURES, /\.[A-Za-z0-9]+$/.test(name) ? name : `${name}.pptx`)
}

/** @param {string} name */
export function readFixture(name) {
	return readFile(fixturePath(name))
}

/** Every `.pptx` in the corpus, in a stable order. */
export const fixtureNames = (await readdir(FIXTURES)).filter((name) => name.endsWith('.pptx')).sort()

// A floor, not a pin: fixtures are added often and pinning the count would fail on every
// addition, which trains people to edit the number without reading it. Raise it when the
// corpus grows enough that the current value stops being evidence of anything.
const MIN_CORPUS = 40
if (fixtureNames.length < MIN_CORPUS) {
	throw new Error(
		`only ${fixtureNames.length} .pptx fixtures found in ${FIXTURES} (expected at least ${MIN_CORPUS}). ` +
			'Every corpus invariant iterates this list, so a short one makes them all pass while measuring nothing.'
	)
}

// `dist/read.js` and `dist/script.js` are pulled in on first use rather than at import time.
// Two dozen read tests want nothing from this module but `fixturePath`, and making them each
// load the script converter to get it would trade one duplication for a slower one.
let deps = null
function loadDeps() {
	deps ??= Promise.all([import('../../dist/read.js'), import('../../dist/script.js')])
	return deps
}

/** @type {Map<string, Promise<import('../../dist/script.js').DeckIr>>} */
const irCache = new Map()

/**
 * The deck IR for a fixture, converted once per test file and shared thereafter.
 *
 * Loading and converting the whole corpus costs ~590 ms, and the four `script-*.test.js`
 * files did it eighteen times between them — the conversion is ~96 % of what each of those
 * loops spends, since printing over an IR already in hand is ~20 ms for all 44.
 *
 * **The returned IR is shared, so treat it as frozen.** `canonicalDeckIr` builds its own
 * structure and the printers only read, so the existing callers are safe; a test that needs
 * to perturb one must `structuredClone` it first, and one that needs two independent
 * conversions wants {@link freshIr}.
 *
 * @param {string} name
 */
export function irFor(name) {
	const key = fixturePath(name)
	let pending = irCache.get(key)
	if (pending === undefined) {
		pending = freshIr(name)
		irCache.set(key, pending)
	}
	return pending
}

/**
 * A conversion that does not touch the cache — for the determinism check, which compares two
 * runs and would compare a cached IR against itself.
 *
 * @param {string} name
 * @returns {Promise<import('../../dist/script.js').DeckIr>}
 */
export async function freshIr(name) {
	const [{ Presentation }, { readModelToIr }] = await loadDeps()
	return readModelToIr(await Presentation.load(await readFixture(name)))
}

/**
 * The fixture loaded into the deep read model. Not cached — most callers then mutate it.
 *
 * @param {string} name
 * @returns {Promise<import('../../dist/read.js').Presentation>}
 */
export async function openFixture(name) {
	const [{ Presentation }] = await loadDeps()
	return Presentation.load(await readFixture(name))
}
