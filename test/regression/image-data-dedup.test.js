import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// Identical base64 `data:` images added to a slide previously embedded one media part PER
// insertion: the duplicate-media check matched on `path`, but inline images carry no real
// path (they all share the `preencoded.<extn>` placeholder), so it never fired. Such images
// are now matched by their data payload, so an identical inline image reuses the original
// `Target` and is embedded once (issue #1339). Distinct images must still embed separately.

// Two visibly-distinct 1x1 PNGs (different pixel bytes → different base64 payloads).
const PNG_A =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const PNG_B =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

function countMedia(zip) {
	// JSZip also lists the `ppt/media/` directory itself as an entry — exclude it.
	return listEntries(zip).filter((p) => p.startsWith('ppt/media/') && !p.endsWith('/')).length
}

function countPics(xml) {
	return (xml.match(/<p:pic>/g) || []).length
}

defineRegressionSuite('Image base64 data de-duplication (#1339)', [
	{
		// Same inline image twice on one slide → both pictures render, but one media part.
		name: 'identical data images on a slide embed a single media part',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_A, x: 1, y: 1, w: 1, h: 1 })
				s.addImage({ data: PNG_A, x: 3, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(countPics(xml) === 2, `expected 2 pictures; got ${countPics(xml)}`)
			assert(countMedia(zip) === 1, `expected 1 media part for the duplicate image; got ${countMedia(zip)}`)
		},
	},
	{
		// Both pictures reference the same media Target via the slide rels.
		name: 'duplicate image relationships share one Target',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_A, x: 1, y: 1, w: 1, h: 1 })
				s.addImage({ data: PNG_A, x: 3, y: 1, w: 1, h: 1 })
			})
			const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			const targets = [...rels.matchAll(/Target="(\.\.\/media\/[^"]+)"/g)].map((m) => m[1])
			assert(targets.length === 2, `expected 2 image relationships; got ${targets.length}`)
			assert(
				targets[0] === targets[1],
				`expected both relationships to share one Target; got ${JSON.stringify(targets)}`
			)
		},
	},
	{
		// Distinct inline images must NOT be collapsed — each keeps its own media part.
		name: 'distinct data images embed separate media parts',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_A, x: 1, y: 1, w: 1, h: 1 })
				s.addImage({ data: PNG_B, x: 3, y: 1, w: 1, h: 1 })
			})
			assert(countMedia(zip) === 2, `expected 2 media parts for distinct images; got ${countMedia(zip)}`)
		},
	},
])
