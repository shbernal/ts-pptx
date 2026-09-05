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
import { describe, expect, test } from 'vitest'
import PresentationCore from '../../../src/presentation.ts'
import { createNodeRuntime } from '../../../src/runtime/node.ts'
import { ALL_CONSTRUCT_FAMILIES } from '../../../src/entry-families.ts'
import { chartFamily } from '../../../src/families/chart.ts'
import { composeFamilies, SLIDE_METHOD_FAMILIES } from '../../../src/families/shared.ts'
import { assert, assertEqual } from '../../helpers.js'

const SERIES = [{ name: 'Rev', labels: ['Q1', 'Q2'], values: [1, 2] }]

const composed = (families) => new PresentationCore(createNodeRuntime(), families)
const withoutCharts = ALL_CONSTRUCT_FAMILIES.filter((family) => family !== chartFamily)

// docProps/core.xml carries `new Date()` dcterms timestamps; blank them so a pair of builds
// straddling a clock tick is not flaky. Every other part of these decks is deterministic.
const decoder = new TextDecoder()
function decode(bytes) {
	return decoder
		.decode(bytes)
		.replace(/(<dcterms:(?:created|modified)[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/g, '$1$2')
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
	})
})
