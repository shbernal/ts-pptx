// Contract for the construct-family seam (src/families/shared.ts): what a presentation can author
// is a list of families it is composed with, not a set of imports its slide class makes. Three
// things have to hold for that list to be safe to vary, and none of them are visible to a schema
// check:
//
//   1. The order a caller lists families in must not reach the output. Zip part order and
//      `[Content_Types].xml` are byte-significant, so two programs composed with the same families
//      in different orders have to emit the same deck.
//   2. A family a deck never uses must cost it nothing: dropping it changes no byte.
//   3. A method whose family was not composed must say so, naming the family. `undefined is not a
//      function` names neither the family nor the fix.
//
// These import from `src/` rather than `dist/` because the seam is internal: `PresentationCore`'s
// second constructor argument is not on the public surface.
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest'
import PresentationCore from '../../../src/presentation.ts'
import { createNodeRuntime } from '../../../src/runtime/node.ts'
import { ALL_CONSTRUCT_FAMILIES } from '../../../src/entry-families.ts'
import { chartFamily } from '../../../src/families/chart.ts'
import { measureFamily } from '../../../src/families/measure.ts'
import {
	CHILD_DESCRIPTOR_FAMILIES,
	composeFamilies,
	PRESENTATION_METHOD_FAMILIES,
	SLIDE_METHOD_FAMILIES,
} from '../../../src/families/shared.ts'
// From `src/`, not `dist/`: `warn` here is the src-side module, and `test/helpers.js`'s
// `captureDiagnostics` installs its handler on the built one, which is a different singleton.
import { setDiagnosticHandler } from '../../../src/diagnostics.ts'
import { assert, assertEqual } from '../../helpers.js'

const SERIES = [{ name: 'Rev', labels: ['Q1', 'Q2'], values: [1, 2] }]

const composed = (families) => new PresentationCore(createNodeRuntime(), families)
const withoutCharts = ALL_CONSTRUCT_FAMILIES.filter((family) => family !== chartFamily)

// Freeze the clock for the whole file, because two parts of these decks read it and only one of
// them is text. `docProps/core.xml` carries `new Date()` dcterms stamps, and so does the
// `docProps/core.xml` *inside* every `ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` a chart
// embeds (`gen/chart/embed-xlsx.ts`). Blanking the first with a regex was the previous fix here,
// and it could never have reached the second: the embedded workbook is a deflated zip by the time
// this sees it, so its timestamp is not a string to match. That left `assertSameParts` failing on
// `Microsoft_Excel_Worksheet2.xlsx` whenever the two builds in a test straddled a one-second tick
// — rare, unreproducible in isolation, and nothing to do with the family seam under test. One
// frozen `Date` covers both stamps, and any third that grows later.
beforeAll(() => {
	vi.useFakeTimers({ toFake: ['Date'] })
	vi.setSystemTime(new Date('2026-01-01T00:00:00Z'))
})
afterAll(() => {
	vi.useRealTimers()
})

// latin1 rather than utf-8, so every byte round-trips to one code point and the comparison below
// is byte-exact. Most of these parts are XML, but the embedded workbooks are compressed, and
// utf-8 decoding folds distinct invalid sequences onto the same U+FFFD.
const decoder = new TextDecoder('latin1')
function decode(bytes) {
	return decoder.decode(bytes)
}

/** Build a deck through a family list and return its parts as `path -> text`, in emission order. */
async function partsOf(families, author) {
	const pres = composed(families)
	author(pres)
	return new Map((await pres.toParts()).map((part) => [part.path, decode(part.data)]))
}

function assertSameParts(a, b, label) {
	assertEqual(JSON.stringify([...a.keys()]), JSON.stringify([...b.keys()]), `${label}: part paths or their order`)
	for (const [path, text] of a) assert(b.get(path) === text, `${label}: ${path} differs`)
}

/** A deck that reaches several families at once, including one inside a slide master. */
function richDeck(pres) {
	pres.defineSlideMaster({
		title: 'MASTER',
		objects: [
			{ rect: { x: 0, y: 0, w: 10, h: 0.4, fill: { color: 'E7E7E7' } } },
			{ chart: { type: 'bar', data: SERIES, options: { x: 7, y: 0.6, w: 3, h: 2 } } },
		],
	})
	const slide = pres.addSlide({ masterTitle: 'MASTER' })
	slide.addText('composed', { x: 1, y: 1, w: 4, h: 1 })
	slide.addChart(SERIES, { type: 'bar', x: 1, y: 2, w: 4, h: 3 })
	slide.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Tighten this', x: 1, y: 0.5 })
	slide.addNotes('speaker notes')
	slide.addGroup([{ rect: { x: 1, y: 5, w: 1, h: 0.5 } }, { text: { text: 'in a group', options: { x: 2, y: 5 } } }])
}

describe('construct families', () => {
	test('the order the caller lists them in does not reach the package', async () => {
		const canonical = await partsOf(ALL_CONSTRUCT_FAMILIES, richDeck)
		const reversed = await partsOf([...ALL_CONSTRUCT_FAMILIES].reverse(), richDeck)
		// Both halves of the chart family have to have run for this to be worth anything: the slide
		// method, and the `{ chart }` descriptor inside the master. Two chart parts is what says so.
		assertEqual(
			[...canonical.keys()].filter((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path)).length,
			2,
			'chart parts in the rich deck'
		)
		// Families are addressed by the method or descriptor key they claim, so listing them in
		// another order can only matter if two claimed the same one. None do, and this is what says so.
		assertSameParts(canonical, reversed, 'reversed family list')
	})

	test('dropping a family the deck never uses changes nothing', async () => {
		const author = (pres) => pres.addSlide().addText('text only', { x: 1, y: 1, w: 4, h: 1 })
		assertSameParts(
			await partsOf(ALL_CONSTRUCT_FAMILIES, author),
			await partsOf(withoutCharts, author),
			'text deck without the chart family'
		)
	})

	test('a method whose family was not composed names the family', () => {
		const slide = composed(withoutCharts).addSlide()
		expect(() => slide.addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 4, h: 3 })).toThrow(/"chart"/)
		try {
			slide.addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 4, h: 3 })
		} catch (err) {
			assertEqual(err.code, 'family/not-composed', 'error code')
			assertEqual(err.name, 'UnsupportedFeatureError', 'error class')
		}
		// The families that *are* composed still answer, on the same slide.
		assert(slide.addText('still here', { x: 1, y: 1, w: 2, h: 1 }) === slide, 'addText chains')
	})

	test('a presentation method whose family was not composed names the family', () => {
		const pres = composed(ALL_CONSTRUCT_FAMILIES.filter((family) => family !== measureFamily))
		expect(() => pres.measureText('wide enough?', { wIn: 3, fontSize: 18 })).toThrow(/"measure"/)
		try {
			pres.tableLayout([[{ text: 'a' }]], { x: 1, y: 1, w: 4 })
		} catch (err) {
			assertEqual(err.code, 'family/not-composed', 'error code')
		}
		// Measuring is a convenience over the `ts-pptx/measure` subpath, so a presentation without it
		// still authors and still writes.
		assertEqual(
			typeof composed(ALL_CONSTRUCT_FAMILIES).measureText('x', { wIn: 3, fontSize: 18 }).heightIn,
			'number',
			'measured'
		)
	})

	test('every slide method the seam names is supplied by the full list', () => {
		const { authors } = composeFamilies(ALL_CONSTRUCT_FAMILIES)
		const names = new Set(ALL_CONSTRUCT_FAMILIES.map((family) => family.name))
		for (const [method, family] of Object.entries(SLIDE_METHOD_FAMILIES)) {
			// Both directions: the seam attributes every method to a family that is in the list, and
			// the list between them supplies every method the seam names. A method with no author is
			// bound to a thrower, which looks like a function from the outside and fails every call.
			assert(names.has(family), `${method} is attributed to "${family}", which is not in the full list`)
			assert(typeof authors[method] === 'function', `${method} is not supplied by any family`)
		}
		// Same for the presentation's own methods, minus the one that needs a DOM: `tableToSlides`
		// comes from the table family's live-DOM half, which only the browser entry composes.
		const { presentationAuthors } = composeFamilies(ALL_CONSTRUCT_FAMILIES)
		for (const [method, family] of Object.entries(PRESENTATION_METHOD_FAMILIES)) {
			assert(names.has(family), `${method} is attributed to "${family}", which is not in the full list`)
			if (method === 'tableToSlides') continue
			assert(typeof presentationAuthors[method] === 'function', `${method} is not supplied by any family`)
		}
	})

	test('every child descriptor the seam names is supplied by the full list', () => {
		const { children } = composeFamilies(ALL_CONSTRUCT_FAMILIES)
		const names = new Set(ALL_CONSTRUCT_FAMILIES.map((family) => family.name))
		// Both directions, as for the methods above: nothing claims a descriptor key the full list
		// cannot author, and nothing authors one the seam cannot attribute to a family.
		for (const [key, family] of Object.entries(CHILD_DESCRIPTOR_FAMILIES)) {
			assert(names.has(family), `the '${key}' descriptor is attributed to "${family}", which is not in the full list`)
			assert(typeof children[key] === 'function', `the '${key}' descriptor is not supplied by any family`)
		}
		for (const key of Object.keys(children))
			assert(CHILD_DESCRIPTOR_FAMILIES[key] !== undefined, `the '${key}' descriptor is attributed to no family`)
	})

	// The condition the seam made reachable: a descriptor key is only rejected by the *types* when
	// nothing claims it, and every key here is claimed by some family. So `{ chart: … }` on a
	// presentation composed without charts type-checks, and used to leave the deck silently.
	describe('a child descriptor whose family was not composed', () => {
		/** Collect the src-side diagnostics `fn` raises. */
		function diagnosticsOf(fn) {
			const seen = []
			setDiagnosticHandler((d) => seen.push(d))
			try {
				fn()
			} finally {
				setDiagnosticHandler(null)
			}
			return seen
		}

		/** @type {import('../../../src/types/index.ts').SlideMasterObject} */
		const chartChild = { chart: { type: 'bar', data: SERIES, options: { x: 1, y: 1, w: 3, h: 2 } } }

		test('warns from a slide master, naming the family', () => {
			const pres = composed(withoutCharts)
			const seen = diagnosticsOf(() =>
				pres.defineSlideMaster({ title: 'MASTER', objects: [{ rect: { x: 0, y: 0, w: 10, h: 0.4 } }, chartChild] })
			)
			assertEqual(seen.length, 1, 'one diagnostic')
			assertEqual(seen[0].code, 'family/child-not-composed', 'code')
			assert(seen[0].message.includes('"chart"'), `message names the family: ${seen[0].message}`)
			assert(seen[0].message.includes('defineSlideMaster()'), `message names the call: ${seen[0].message}`)
			// The descriptor the composed families *do* claim still landed. Read off the layout the
			// master became, which is internal state a cast reaches rather than the public surface.
			const [layout] = /** @type {any} */ (pres)._slideLayouts.slice(-1)
			assertEqual(layout._slideObjects.length, 1, 'only the rect was authored')
		})

		test('warns from addGroup, naming the family', () => {
			const slide = composed(ALL_CONSTRUCT_FAMILIES.filter((family) => family.name !== 'image')).addSlide()
			const seen = diagnosticsOf(() =>
				slide.addGroup([
					{ text: { text: 'kept', options: { x: 1, y: 1, w: 2, h: 0.5 } } },
					{ image: { data: 'image/png;base64,iVBORw0KGgo=', x: 1, y: 2, w: 1, h: 1 } },
				])
			)
			const [first] = seen
			assertEqual(first.code, 'family/child-not-composed', 'code')
			assert(first.message.includes('"image"'), `message names the family: ${first.message}`)
			assert(first.message.includes('addGroup()'), `message names the call: ${first.message}`)
		})

		test('a key no family has ever claimed still reports the typo, not a family', () => {
			const slide = composed(ALL_CONSTRUCT_FAMILIES).addSlide()
			// The cast is the point: with no family claiming it, the types reject this key outright,
			// which is exactly the protection a claimed-but-uncomposed key does not get.
			const seen = diagnosticsOf(() => slide.addGroup([/** @type {any} */ ({ notAThing: { x: 1, y: 1 } })]))
			assertEqual(seen[0].code, 'group/unrecognized-child', 'code')
			assert(seen[0].message.includes('notAThing'), `message names the key: ${seen[0].message}`)
		})
	})
})
