// Contract for the part-contributor seam (src/package/parts/shared.ts): which construct
// families put parts in a package is a list handed to `buildPackageParts`, not a set of imports
// it makes itself. Two properties have to hold for that list to be safe to vary, and both are
// byte-level, so neither shows up in a schema check:
//
//   1. The order a caller lists contributors in must not reach the output. Zip part order and
//      `[Content_Types].xml` are both byte-significant, so the same deck written by two programs
//      that list the same families differently has to come out identical.
//   2. Dropping a family a deck does not use must change nothing, and dropping one it does use
//      must remove exactly that family's parts.
//
// These import from `src/` rather than `dist/` because the seam is internal: `PackageSource` and
// the contributor list are not on the public surface.
import { describe, test } from 'vitest'
import TsPptx from '../../../src/node.ts'
import { buildPackageParts } from '../../../src/package/assemble.ts'
import { composeFamilies } from '../../../src/families/shared.ts'
import { ALL_CONSTRUCT_FAMILIES } from '../../../src/entry-families.ts'
import { ALL_PART_CONTRIBUTORS } from '../../../src/package/contributors.ts'
import { chartContributor } from '../../../src/package/parts/chart.ts'
import { assert, assertEqual } from '../../helpers.js'

const SERIES = [{ name: 'Rev', labels: ['Q1', 'Q2'], values: [1, 2] }]

/** A deck that reaches all three families: a chart, a comment, and the notes every slide carries. */
function richDeck() {
	const pres = new TsPptx()
	const slide = pres.addSlide()
	slide.addText('contributed parts', { x: 1, y: 1, w: 4, h: 1 })
	slide.addChart(SERIES, { type: 'bar', x: 1, y: 2, w: 4, h: 3 })
	slide.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Tighten this', x: 1, y: 0.5 })
	pres.addSlide().addText('second slide', { x: 1, y: 1, w: 4, h: 1 })
	return pres
}

/** A deck that reaches none of the optional families -- text boxes and the notes slides only. */
function textDeck() {
	const pres = new TsPptx()
	pres.addSlide().addText('text only', { x: 1, y: 1, w: 4, h: 1 })
	return pres
}

/**
 * The slice of deck state the packager reads. `Presentation` assembles this privately with the
 * full contributor list; here the list is the variable under test.
 */
function sourceWith(pres, partContributors) {
	return {
		runtime: pres._runtime,
		presentation: pres.internalPresentation,
		customProperties: pres._customProperties,
		fontMetrics: pres._fontMetrics,
		renderers: composeFamilies(ALL_CONSTRUCT_FAMILIES).renderers,
		partContributors,
	}
}

// docProps/core.xml carries `new Date()` dcterms timestamps; blank them so a pair of builds
// straddling a clock tick is not flaky. Every other part of these decks is deterministic.
const decoder = new TextDecoder()
function decode(bytes) {
	return decoder
		.decode(bytes)
		.replace(/(<dcterms:(?:created|modified)[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/g, '$1$2')
}

/** Build a deck through the seam and return its parts as `path -> text`, in emission order. */
async function partsOf(makePres, partContributors) {
	const parts = await buildPackageParts(sourceWith(makePres(), partContributors), {})
	return new Map(parts.map((part) => [part.path, decode(part.data)]))
}

function assertSameParts(a, b, label) {
	assertEqual(JSON.stringify([...a.keys()]), JSON.stringify([...b.keys()]), `${label}: part paths or their order`)
	for (const [path, text] of a) {
		assert(b.get(path) === text, `${label}: ${path} differs`)
	}
}

describe('part contributors', () => {
	test('the order the caller lists them in does not reach the package', async () => {
		const canonical = await partsOf(richDeck, ALL_PART_CONTRIBUTORS)
		const reversed = await partsOf(richDeck, [...ALL_PART_CONTRIBUTORS].reverse())
		// Every contributor's `order` decides where its parts and content-type entries land, so
		// reversing the list has to be invisible -- otherwise a tier shipping a subset in a
		// different order would emit a different deck for the same input.
		assertSameParts(canonical, reversed, 'reversed contributor list')
	})

	test('dropping a family the deck never uses changes nothing', async () => {
		const full = await partsOf(textDeck, ALL_PART_CONTRIBUTORS)
		const withoutCharts = await partsOf(
			textDeck,
			ALL_PART_CONTRIBUTORS.filter((contributor) => contributor !== chartContributor)
		)
		// This is what a tier buys: a text-only program that never links the chart part builders
		// writes the same bytes as one that does.
		assertSameParts(full, withoutCharts, 'text deck without the chart contributor')
	})

	test('dropping the chart contributor removes the chart parts and nothing else', async () => {
		const full = await partsOf(richDeck, ALL_PART_CONTRIBUTORS)
		const without = await partsOf(
			richDeck,
			ALL_PART_CONTRIBUTORS.filter((contributor) => contributor !== chartContributor)
		)

		const isChartPart = (path) => path.startsWith('ppt/charts/') || path.startsWith('ppt/embeddings/')
		const dropped = [...full.keys()].filter((path) => !without.has(path))
		assert(dropped.length > 0, 'the rich deck should have produced chart parts to drop')
		assertEqual(
			JSON.stringify(dropped),
			JSON.stringify(dropped.filter(isChartPart)),
			'only chart parts should disappear'
		)
		assertEqual(
			JSON.stringify([...full.keys()].filter((path) => !isChartPart(path))),
			JSON.stringify([...without.keys()]),
			'the remaining parts keep their paths and their order'
		)

		// Everything the chart family did not write is byte-identical; `[Content_Types].xml` is the
		// one part both versions write, and it loses exactly the chart family's entries.
		for (const [path, text] of without) {
			if (path === '[Content_Types].xml') continue
			assert(full.get(path) === text, `${path} should not depend on the chart contributor`)
		}
		const entriesOf = (xml) => (xml ?? '').match(/<(?:Default|Override)\b[^>]*\/>/g) ?? []
		const kept = new Set(entriesOf(without.get('[Content_Types].xml')))
		const lost = entriesOf(full.get('[Content_Types].xml')).filter((entry) => !kept.has(entry))
		assert(lost.length > 0, 'expected the chart content-type entries to be gone')
		for (const entry of lost) {
			assert(
				entry.includes('/ppt/charts/') || entry.includes('Extension="xlsx"'),
				'only the chart entries should be gone; also lost: ' + entry
			)
		}
	})
})
