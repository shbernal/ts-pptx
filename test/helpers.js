// Tests intentionally read the generated .pptx with jszip rather than the
// library's own src/zip.ts (fflate). The write path uses fflate, so reading
// back with a *different* zip implementation makes jszip an independent oracle:
// a round-trip bug in fflate can't mask itself by being used on both sides.
// Keep jszip as a devDep for this reason — do not "consolidate" onto src/zip.ts.
import JSZip from 'jszip'
import TsPptx, { setDiagnosticHandler } from '../dist/node.js'
import { describe, test } from 'vitest'

/**
 * A 1x1 transparent PNG, in the bare `type;base64,…` spelling `addImage` takes.
 *
 * The same 67 bytes were pasted into nineteen files under six different names (`PNG_DATA`,
 * `PNG_1X1`, `PNG_1PX`, `PNG_A`, `PNG_B`, `PNG`), which made a grep for "the tests' image"
 * miss most of them and made two files look like they used different images when they did
 * not. Where a test needs a *second*, distinguishable image, keep a local constant and say
 * what makes it different — that is a real distinction, unlike a sixth alias for this one.
 */
const PNG_1X1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

/** The same bytes with the `data:` scheme, for the paths that assert both spellings are taken. */
const PNG_1X1_DATA_URI = `data:${PNG_1X1}`

async function build(buildFn) {
	const pres = new TsPptx()
	buildFn(pres)
	// stream() is typed for every output target (string/Blob/ArrayBuffer/Uint8Array);
	// under Node it resolves to a Uint8Array, which the tests rely on for byte reads.
	const buf = /** @type {Uint8Array} */ (await pres.stream())
	const zip = await JSZip.loadAsync(buf)
	return { pres, zip, buf }
}

async function readEntry(zip, path) {
	const entry = zip.file(path)
	if (!entry) throw new Error('zip entry not found: ' + path)
	return entry.async('string')
}

function listEntries(zip) {
	return Object.keys(zip.files)
}

/**
 * Every part of a `.pptx`, keyed by zip entry name, as raw bytes.
 *
 * Sixteen read tests each re-derived this. Fifteen read `'uint8array'` and one read
 * `'string'`; this is the byte spelling, which is the stronger of the two — a decoded
 * comparison cannot see a BOM or an encoding change, and every caller is asking whether
 * the bytes moved.
 *
 * @param {Uint8Array | Buffer} pptxBytes
 * @returns {Promise<Map<string, Uint8Array>>}
 */
async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

/**
 * Every part in `before` survives into `after` with identical bytes, except those named in
 * `allowedToChange`.
 *
 * This is the "did the edit stay local?" assertion, which was written out longhand in a
 * dozen read tests with three different filter spellings. Two properties it adds over the
 * hand-rolled loops:
 *
 * **It fails when it compared nothing.** A loop whose filter stops matching passes having
 * checked nothing, and that is indistinguishable from success in a reporter — the same
 * failure mode `test/read/corpus.js` guards the fixture list against.
 *
 * **A part missing from `after` is a failure, not a skip.** One hand-rolled copy skipped
 * absent parts, which turns "this part was deleted" into a pass.
 *
 * `allowedToChange` is permission, not obligation: it says nothing about whether those
 * parts actually differ. Where that matters the caller asserts it separately, which keeps
 * the two claims legible instead of folding them into one helper that means both.
 *
 * @param {Map<string, Uint8Array>} before
 * @param {Map<string, Uint8Array>} after
 * @param {Iterable<string>} [allowedToChange]
 * @param {string} [label]
 */
function assertUnchangedExcept(before, after, allowedToChange = [], label = '') {
	const allowed = new Set(allowedToChange)
	const prefix = label ? label + ': ' : ''
	let checked = 0
	for (const [name, body] of before) {
		if (allowed.has(name)) continue
		const actual = after.get(name)
		assert(actual, `${prefix}${name} is missing from the saved package`)
		assert(bytesEqual(body, actual), `${prefix}${name} should be untouched`)
		checked++
	}
	assert(
		checked > 0,
		`${prefix}compared no parts — every one of the ${before.size} input parts was allowed to change, ` +
			'so this assertion proved nothing'
	)
}

/**
 * Declare a regression suite from an array of `{ name, fn }` cases.
 *
 * **One signature.** There used to be two — a three-argument form carrying a provenance tag
 * (`'legacy bug-14'`, `'upstream-issue-1451'`) which was destructured out and then dropped on
 * the floor, so thirty-six suites recorded where their regression came from in a string that
 * reached no reporter and no reader who was not looking at the call site. Those tags are now
 * part of the suite name, where they are visible; a second positional argument is an error
 * rather than a silently ignored one.
 *
 * **`fn` is handed to vitest as-is**, not wrapped in `async () => await fixture.fn()`. The
 * wrapper put this file at the top of every regression failure's stack, above the case that
 * actually failed.
 *
 * **Modifiers are per case**, because a case cannot reach `test.skipIf` / `test.todo` /
 * `test.concurrent` from inside a plain array: `{ name, fn, skipIf: !validatorInstalled }`,
 * `{ name, todo: true }`, `{ name, fn, timeout: 30_000 }`. Note that `concurrent` is safe only
 * for cases that touch no process global — `captureDiagnostics` and `setDiagnosticHandler` are
 * process-wide, and their safety rests on cases within a file running serially.
 *
 * @param {string} suiteName
 * @param {{ name: string, fn?: () => unknown, only?: boolean, skip?: boolean, skipIf?: unknown,
 *           runIf?: unknown, todo?: boolean, fails?: boolean, concurrent?: boolean,
 *           timeout?: number }[]} cases
 */
function defineRegressionSuite(suiteName, cases) {
	if (!Array.isArray(cases)) {
		throw new Error(
			`defineRegressionSuite(${JSON.stringify(suiteName)}, …) takes an array of test cases as its ` +
				'second argument. The three-argument form carrying a provenance tag is gone — fold the tag ' +
				'into the suite name, where it is actually reported.'
		)
	}

	describe(suiteName, () => {
		for (const fixture of cases) {
			if (fixture.todo) {
				test.todo(fixture.name)
				continue
			}
			// `any`, because this *is* dynamic dispatch: each step narrows vitest's chainable to a
			// different member of the family, and the whole point is that a case picks its own.
			/** @type {any} */
			let define = fixture.concurrent ? test.concurrent : test
			if (fixture.fails) define = define.fails
			if (fixture.only) define = define.only
			else if (fixture.skip) define = define.skip
			else if (fixture.skipIf !== undefined) define = define.skipIf(fixture.skipIf)
			else if (fixture.runIf !== undefined) define = define.runIf(fixture.runIf)
			define(fixture.name, fixture.fn, fixture.timeout)
		}
	})
}

/**
 * @param {unknown} cond
 * @param {string} [msg]
 * @returns {asserts cond}
 */
function assert(cond, msg) {
	if (!cond) throw new Error('assertion failed: ' + msg)
}

function assertEqual(actual, expected, msg) {
	if (actual !== expected)
		throw new Error(
			'assertion failed: ' + (msg || '') + ' expected ' + JSON.stringify(expected) + ' got ' + JSON.stringify(actual)
		)
}

function assertIncludes(haystack, needle, label) {
	assert(haystack.includes(needle), `expected ${label || 'value'} to include ${needle}; got: ${haystack}`)
}

function assertNotIncludes(haystack, needle, label) {
	assert(!haystack.includes(needle), `expected ${label || 'value'} not to include ${needle}; got: ${haystack}`)
}

/**
 * Byte-for-byte equality of two `Uint8Array`s (or anything array-like), null-safe.
 *
 * Both spellings that were in the tree — with and without the `a && b` guard — are folded into
 * the guarded one, which is the superset: an absent part compares unequal rather than throwing,
 * which is what "did this part change?" means at the fifteen call sites that ask it.
 */
function bytesEqual(a, b) {
	return Boolean(a && b && a.length === b.length && a.every((value, index) => value === b[index]))
}

/**
 * True when `fn` throws. Deliberately says nothing about *what* it threw, so pair it with a
 * separate assertion on the message when the distinction matters — a bare `throws()` passes
 * just as happily on a `TypeError` from a typo in the test as on the guard under test.
 */
function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

function xmlBlocks(xml, tagName) {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[\\s\\S]*?<\\/${escapedName}>`, 'g')
	return xml.match(re) || []
}

function firstXmlBlock(xml, tagName, label = tagName) {
	const block = xmlBlocks(xml, tagName)[0]
	assert(block, `expected ${label} block in XML; got: ${xml}`)
	return block
}

function xmlAttributes(tag) {
	const attrs = {}
	for (const match of tag.matchAll(/\s([\w:-]+)="([^"]*)"/g)) {
		attrs[match[1]] = match[2]
	}
	return attrs
}

function selfClosingTags(xml, tagName) {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[^>]*/>`, 'g')
	return xml.match(re) || []
}

function xmlOpeningTags(xml, tagName) {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[^>]*(?:/>|>)`, 'g')
	return xml.match(re) || []
}

function contentTypeDefaultExtensions(xml) {
	return selfClosingTags(xml, 'Default').map((tag) => xmlAttributes(tag).Extension)
}

function contentTypeOverrideParts(xml) {
	return selfClosingTags(xml, 'Override').map((tag) => xmlAttributes(tag).PartName)
}

function contentTypeForExtension(xml, extension) {
	const tag = selfClosingTags(xml, 'Default').find((t) => xmlAttributes(t).Extension === extension)
	return tag ? xmlAttributes(tag).ContentType : undefined
}

function assertContentTypeDefault(xml, extension) {
	const extensions = contentTypeDefaultExtensions(xml)
	assert(
		extensions.includes(extension),
		`expected Content_Types Default for ${extension}; got: ${extensions.join(', ')}`
	)
}

function assertNoContentTypeDefault(xml, extension) {
	const extensions = contentTypeDefaultExtensions(xml)
	assert(
		!extensions.includes(extension),
		`did not expect Content_Types Default for ${extension}; got: ${extensions.join(', ')}`
	)
}

function assertContentTypeOverride(xml, partName) {
	const parts = contentTypeOverrideParts(xml)
	assert(parts.includes(partName), `expected Content_Types Override for ${partName}; got: ${parts.join(', ')}`)
}

function assertXmlOrder(xml, before, after, label) {
	const beforeIndex = xml.indexOf(before)
	const afterIndex = xml.indexOf(after)
	assert(beforeIndex !== -1, `expected ${before} in ${label || 'XML'}; got: ${xml}`)
	assert(afterIndex !== -1, `expected ${after} in ${label || 'XML'}; got: ${xml}`)
	assert(
		beforeIndex < afterIndex,
		`expected ${before} before ${after} in ${label || 'XML'}; got order ${beforeIndex} then ${afterIndex}: ${xml}`
	)
}

function nonVisualDrawingProperties(xml) {
	const tags = xmlOpeningTags(xml, 'p:cNvPr')
	return tags.map((tag) => ({ tag, attrs: xmlAttributes(tag) }))
}

function findNonVisualDrawingProperty(xml, attrs) {
	return nonVisualDrawingProperties(xml).find(({ attrs: actual }) =>
		Object.entries(attrs).every(([name, value]) => actual[name] === value)
	)
}

function assertNonVisualDrawingProperty(xml, attrs, label) {
	const match = findNonVisualDrawingProperty(xml, attrs)
	assert(match, `expected ${label || 'p:cNvPr'} with ${JSON.stringify(attrs)}; got: ${xml}`)
	return match
}

/**
 * Run `fn` with a diagnostic handler installed, returning what the library emitted alongside the
 * function's own result. Prefer asserting on `codes` -- a diagnostic's `code` is API and its
 * `message` explicitly is not, so a message assertion breaks on any wording improvement.
 *
 * The handler is process-global (see `setDiagnosticHandler`), so this must not be used from two
 * concurrently-running cases; vitest runs cases within a file serially, which is what makes it safe.
 */
async function captureDiagnostics(fn) {
	const diagnostics = []
	setDiagnosticHandler((d) => diagnostics.push(d))
	try {
		const result = await fn()
		return {
			result,
			diagnostics,
			codes: diagnostics.map((d) => d.code),
			messages: diagnostics.map((d) => d.message),
		}
	} finally {
		setDiagnosticHandler(null)
	}
}

export {
	TsPptx,
	setDiagnosticHandler,
	captureDiagnostics,
	PNG_1X1,
	PNG_1X1_DATA_URI,
	build,
	readEntry,
	listEntries,
	partBodies,
	assertUnchangedExcept,
	defineRegressionSuite,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
	bytesEqual,
	throws,
	xmlBlocks,
	firstXmlBlock,
	xmlAttributes,
	selfClosingTags,
	contentTypeDefaultExtensions,
	contentTypeForExtension,
	contentTypeOverrideParts,
	assertContentTypeDefault,
	assertNoContentTypeDefault,
	assertContentTypeOverride,
	assertXmlOrder,
	nonVisualDrawingProperties,
	findNonVisualDrawingProperty,
	assertNonVisualDrawingProperty,
	xmlOpeningTags,
}
