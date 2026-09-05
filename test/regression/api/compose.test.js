// Acceptance for the composable surface: `createPresentation({ use })` against the published
// entry, not the internal seam. What it has to hold up:
//
//   1. A composed deck is a real deck. The core tier writes the same package a default `TsPptx`
//      writes for the same slides -- part for part, byte for byte.
//   2. A family that was not composed is absent, loudly. `addChart` on a core-tier deck raises
//      `family/not-composed` naming the family, rather than failing as a missing property.
//   3. A family that was composed is there, and puts its parts in the package.
//
// Imports from `dist/` because this is the public surface: the entry's `createPresentation` and
// the `pptx-ts/families` subpath are what a consumer reaches for.
import { describe, expect, test } from 'vitest'
import TsPptx, { createPresentation } from '../../../dist/node.js'
import { charts, comments, tables } from '../../../dist/families.js'
import { assert, assertEqual } from '../../helpers.js'

const SERIES = [{ name: 'Rev', labels: ['Q1', 'Q2'], values: [1, 2] }]

const decoder = new TextDecoder()
/** Parts as `path -> text`, with the two timestamps a pair of builds can straddle blanked. */
async function partsOf(pres) {
	const parts = await pres.toParts()
	return new Map(
		parts.map((part) => [
			part.path,
			decoder
				.decode(part.data)
				.replace(/(<dcterms:(?:created|modified)[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/g, '$1$2'),
		])
	)
}

/** The slides both halves of the byte-for-byte case author, using only core-tier families. */
function coreDeck(pres) {
	const slide = pres.addSlide()
	slide.addText('composed', { x: 1, y: 1, w: 4, h: 1 })
	slide.addShape('rect', { x: 1, y: 2.5, w: 2, h: 1, fill: { color: '0088CC' } })
	slide.addNotes('speaker notes')
	slide.addGroup([{ rect: { x: 4, y: 2.5, w: 1, h: 1 } }])
	return pres
}

describe('createPresentation', () => {
	test('the core tier writes the same package the full class does', async () => {
		const composed = await partsOf(coreDeck(createPresentation()))
		const full = await partsOf(coreDeck(new TsPptx()))
		// Not "a deck it can open" -- the same deck. A tier that emitted a near-miss would be a
		// second implementation of the write path, which is the thing the family seam exists to avoid.
		assertEqual(JSON.stringify([...composed.keys()]), JSON.stringify([...full.keys()]), 'part paths and order')
		for (const [path, text] of composed) assert(full.get(path) === text, `${path} differs from the full tier`)
	})

	test('a family that was not composed says which one it is', () => {
		const slide = createPresentation({ use: [tables] }).addSlide()
		// The one that was composed answers...
		slide.addTable([[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 4 })
		// ...and the one that was not names itself, with a code rather than a TypeError. The cast is
		// half the point: the type has already refused this call, and a consumer who ignores it (or
		// reaches the slide from JavaScript) is who the runtime message is for.
		const untyped = /** @type {{ addChart: (...args: unknown[]) => unknown }} */ (/** @type {unknown} */ (slide))
		let thrown
		try {
			untyped.addChart(SERIES, { type: 'bar', x: 1, y: 2, w: 4, h: 3 })
		} catch (err) {
			thrown = err
		}
		assert(thrown, 'addChart must throw on a deck composed without charts')
		assertEqual(thrown.code, 'family/not-composed', 'error code')
		assert(/"chart"/.test(thrown.message), `the message must name the family; got: ${thrown.message}`)
	})

	test('a composed family puts its parts in the package', async () => {
		const pres = createPresentation({ use: [charts, comments] })
		const slide = pres.addSlide()
		slide.addText('with charts', { x: 1, y: 1, w: 4, h: 1 })
		slide.addChart(SERIES, { type: 'bar', x: 1, y: 2, w: 4, h: 3 })
		slide.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Tighten this', x: 1, y: 0.5 })

		const paths = [...(await partsOf(pres)).keys()]
		for (const expected of ['ppt/charts/chart1.xml', 'ppt/comments/comment1.xml', 'ppt/commentAuthors.xml'])
			assert(
				paths.includes(expected),
				`expected ${expected} among: ${paths.filter((p) => p.includes('chart') || p.includes('omment')).join(', ')}`
			)
		assert(
			paths.some((path) => path.startsWith('ppt/embeddings/')),
			'a chart carries its embedded workbook'
		)
	})

	test('composing a family the core tier already has changes nothing', async () => {
		// A caller being explicit about a core family is a thing to allow, not to punish: the same
		// contributor collected twice would write its parts twice.
		const { notes } = await import('../../../dist/families.js')
		const explicit = await partsOf(coreDeck(createPresentation({ use: [notes, notes] })))
		const implicit = await partsOf(coreDeck(createPresentation()))
		assertEqual(JSON.stringify([...explicit.keys()]), JSON.stringify([...implicit.keys()]), 'part paths and order')
		for (const [path, text] of explicit) assert(implicit.get(path) === text, `${path} differs`)
	})

	test('the default class still carries every family', () => {
		const slide = new TsPptx().addSlide()
		for (const method of [
			'addChart',
			'addTable',
			'addMedia',
			'addComment',
			'addOleObject',
			'addModel3d',
			'addSlideZoom',
		])
			expect(typeof slide[method], `${method} on the full tier`).toBe('function')
		// Not merely present: a method with no family behind it is a thrower, which is also a function.
		expect(() => slide.addTable([[{ text: 'a' }]], { x: 1, y: 1, w: 2 })).not.toThrow()
	})
})
