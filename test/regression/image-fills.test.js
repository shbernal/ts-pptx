import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { ShapeType } from '../../dist/node.js'
import {
	setDiagnosticHandler,
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertIncludes,
	assertNotIncludes,
	firstXmlBlock,
	assertXmlOrder,
} from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/** Read a slide's XML straight out of a PowerPoint-authored fixture package. */
async function fixtureSlideXml(name, n = 1) {
	const buf = await readFile(path.join(__dirname, '..', 'read', 'fixtures', `${name}.pptx`))
	const zip = await JSZip.loadAsync(buf)
	return zip.file(`ppt/slides/slide${n}.xml`).async('string')
}

/** Collapse inter-tag whitespace so an indented oracle and our compact output compare alike. */
function squash(xml) {
	return xml.replace(/>\s+</g, '><')
}

/** Every `<a:tcPr>…</a:tcPr>` (or self-closed `<a:tcPr/>`) in document order. */
function tcPrBlocks(xml) {
	return squash(xml).match(/<a:tcPr(?:\s[^>]*)?\/>|<a:tcPr(?:\s[^>]*)?>[\s\S]*?<\/a:tcPr>/g) || []
}

// 1x1 transparent PNG (data URI). Used to exercise the picture-fill (`<a:blipFill>`) path
// for shapes and text boxes.
const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

async function captureWarnings(fn) {
	const warnings = []
	setDiagnosticHandler((d) => warnings.push(d.message))
	try {
		await fn()
	} finally {
		setDiagnosticHandler(null)
	}
	return warnings
}

defineRegressionSuite('Image (blip) fills', [
	{
		name: 'shape image fill emits blipFill referencing an embedded media rel, before line props',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 3,
					h: 1,
					fill: { type: 'image', image: { data: PNG_1x1 } },
					line: { color: '111111', width: 1 },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapeBlock = firstXmlBlock(xml, 'p:sp', 'shape')

			const m = /<a:blip r:embed="(rId\d+)">/.exec(shapeBlock)
			assert(m, `expected a blipFill referencing a media rel; got: ${shapeBlock}`)
			assertIncludes(shapeBlock, '<a:stretch><a:fillRect/></a:stretch>', 'image fill stretch')
			// Fill must precede the line per CT_ShapeProperties order
			assertXmlOrder(shapeBlock, '<a:blipFill', '<a:ln', 'shape properties')

			// The referenced relationship must exist and target an embedded media file
			const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			assertIncludes(rels, `Id="${m[1]}"`, 'media relationship id present')
			assertIncludes(rels, '/relationships/image', 'media relationship is an image type')
			const media = (await listEntries(zip)).filter((e) => e.startsWith('ppt/media/') && !e.endsWith('/'))
			assert(media.length === 1, `expected exactly one embedded media file; got ${JSON.stringify(media)}`)
		},
	},
	{
		name: 'custGeom shape image fill emits a blipFill clipped by the custom geometry',
		fn: async () => {
			// The ledger capability `blip-fill-shape` is specifically "a <p:sp> whose fill is an
			// <a:blipFill>, with the shape's geometry as the clip". A custom-geometry shape is the
			// sharpest form: the <a:custGeom> path is the clip and the blipFill must follow it,
			// in CT_ShapeProperties order (geometry, then fill, then line).
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.custGeom, {
					x: 1,
					y: 1,
					w: 3,
					h: 2,
					points: [{ x: 0, y: 0 }, { x: 3, y: 0 }, { x: 3, y: 2 }, { close: true }],
					fill: { type: 'image', image: { data: PNG_1x1 } },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapeBlock = firstXmlBlock(xml, 'p:sp', 'shape')

			assertIncludes(shapeBlock, '<a:custGeom>', 'custom geometry present')
			const m = /<a:blip r:embed="(rId\d+)">/.exec(shapeBlock)
			assert(m, `expected a blipFill referencing a media rel; got: ${shapeBlock}`)
			// The geometry must precede the fill so the custom path clips the image
			assertXmlOrder(shapeBlock, '<a:custGeom>', '<a:blipFill', 'shape properties')

			const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			assertIncludes(rels, `Id="${m[1]}"`, 'media relationship id present')
			assertIncludes(rels, '/relationships/image', 'media relationship is an image type')
		},
	},
	{
		name: 'image fill set via `image` alone (no explicit type) still emits a blipFill',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { image: { data: PNG_1x1 } } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, '<a:blipFill', 'blip fill emitted from image-only fill')
		},
	},
	{
		name: 'image fill transparency emits an alphaModFix on the blip',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 3,
					h: 1,
					fill: { type: 'image', image: { data: PNG_1x1 }, transparency: 25 },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, '<a:alphaModFix amt="75000"/>', 'image fill transparency')
		},
	},
	{
		name: 'identical image fills are embedded once (media de-duplication)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { image: { data: PNG_1x1 } } })
				s.addShape(ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { image: { data: PNG_1x1 } } })
			})
			const media = (await listEntries(zip)).filter((e) => e.startsWith('ppt/media/') && !e.endsWith('/'))
			assert(media.length === 1, `expected a single shared media file; got ${JSON.stringify(media)}`)
		},
	},
	{
		name: 'text box accepts an image fill',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1, w: 3, h: 1, fill: { type: 'image', image: { data: PNG_1x1 } } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, '<a:blipFill', 'text box image fill')
		},
	},
	{
		name: 'image fill with neither path nor data warns and falls back (no blipFill)',
		fn: async () => {
			let zip
			const warnings = await captureWarnings(async () => {
				;({ zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape(ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { type: 'image', image: {} } })
				}))
			})
			assert(
				warnings.some((w) => w.includes('image fill requires')),
				`expected a warning; got ${JSON.stringify(warnings)}`
			)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNotIncludes(xml, '<a:blipFill', 'no blip fill when image source is missing')
		},
	},
	{
		name: 'SVG image fill is rejected with a warning (raster only)',
		fn: async () => {
			let zip
			const warnings = await captureWarnings(async () => {
				;({ zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape(ShapeType.rect, {
						x: 1,
						y: 1,
						w: 3,
						h: 1,
						fill: { type: 'image', image: { data: 'image/svg+xml;base64,PHN2Zy8+' } },
					})
				}))
			})
			assert(
				warnings.some((w) => w.includes('SVG image fills are not supported')),
				`expected an SVG warning; got ${JSON.stringify(warnings)}`
			)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNotIncludes(xml, '<a:blipFill', 'no blip fill for unsupported SVG source')
		},
	},
])

// Picture fill of a TABLE CELL — `a:blipFill` inside `a:tcPr`. Gated on the
// PowerPoint-authored oracle `test/read/fixtures/table-cell-image-fill.pptx`
// (4x2 table, "No Style, No Grid", so every fill in it was set explicitly).
//
// Before this was wired, a cell image fill type-checked, reached
// `genXmlColorSelection`'s `case 'image'` with no registered relationship, and
// silently degraded to `<a:noFill/>` with a warning — nothing registered cell
// fills as media. These pin the fix in both directions.
defineRegressionSuite('Table cell image (blip) fills', [
	{
		name: 'oracle: PowerPoint puts a cell picture fill last in a:tcPr, after the borders',
		fn: async () => {
			// Reads the fixture rather than a transcribed .oracle.json, so the assertion cannot
			// drift from the bytes PowerPoint actually wrote.
			const oracle = await fixtureSlideXml('table-cell-image-fill')
			const cells = tcPrBlocks(oracle)
			assert(cells.length === 8, `expected 8 cells in the oracle; got ${cells.length}`)

			// A1: picture fill, no borders but the style's lnB.
			assertIncludes(cells[0], '<a:blipFill>', 'A1 carries a blipFill')
			assertXmlOrder(cells[0], '<a:lnB', '<a:blipFill', 'oracle A1 tcPr')
			// B1 (index 2): picture fill AND all four explicit borders — the child-order case.
			// `CT_TableCellProperties` puts `EG_FillProperties` at order 7, after lnL/lnR/lnT/lnB.
			assertXmlOrder(cells[2], '<a:lnL', '<a:lnR', 'oracle B1 tcPr')
			assertXmlOrder(cells[2], '<a:lnB', '<a:blipFill', 'oracle B1 tcPr')
			// Merged row: the origin (gridSpan="2") holds the fill; the covered cell is bare.
			assertIncludes(cells[4], '<a:blipFill>', 'merged origin carries the fill')
			assert(cells[5] === '<a:tcPr/>', `covered cell is bare in the oracle; got ${cells[5]}`)
			// All four picture cells share ONE relationship — PowerPoint dedupes the source.
			const embeds = [...oracle.matchAll(/<a:blip r:embed="(rId\d+)"/g)].map((m) => m[1])
			assert(embeds.length === 4, `expected 4 blips in the oracle; got ${embeds.length}`)
			assert(new Set(embeds).size === 1, `expected one shared rel; got ${JSON.stringify(embeds)}`)
		},
	},
	{
		name: 'a cell image fill emits a blipFill in a:tcPr after the borders, matching the oracle',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(
					[
						[
							{ text: 'A1', options: { fill: { type: 'image', image: { data: PNG_1x1 } } } },
							{ text: 'A2', options: { fill: { color: 'FF0000' } } },
						],
					],
					{ x: 1, y: 1, w: 6, h: 2, border: { type: 'solid', color: '000000', width: 1 } }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cells = tcPrBlocks(xml)

			const m = /<a:blip r:embed="(rId\d+)"/.exec(cells[0])
			assert(m, `expected a blipFill referencing a media rel; got: ${cells[0]}`)
			assertIncludes(cells[0], '<a:stretch><a:fillRect/></a:stretch>', 'stretched cell fill')
			// Same child order the oracle shows, and the order CT_TableCellProperties requires.
			assertXmlOrder(cells[0], '<a:lnB', '<a:blipFill', 'cell tcPr')
			// The sibling solid fill is untouched.
			assertIncludes(cells[1], '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>', 'solid cell fill')

			const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			assertIncludes(rels, `Id="${m[1]}"`, 'media relationship id present')
			assertIncludes(rels, '/relationships/image', 'media relationship is an image type')

			// We spell the blipFill with `dpi="0" rotWithShape="1"` + `<a:srcRect/>`, where the
			// oracle's *stretched* cells carry a bare `<a:blipFill>`. Deliberate: both attributes
			// are optional with no schema default, and PowerPoint writes this exact attribute set
			// inside `a:tcPr` for its *tiled* cell (D1 in the fixture), so the form is authored by
			// PowerPoint itself. Changing the shared `genXmlImageFill` would move bytes for every
			// shape and text-box fill for no behavioural gain.
			assertIncludes(cells[0], '<a:blipFill dpi="0" rotWithShape="1">', 'shared emitter spelling')
			assertIncludes(squash(await fixtureSlideXml('table-cell-image-fill')), '<a:blipFill dpi="0" rotWithShape="1">')
		},
	},
	{
		name: 'a table-level image fill reaches every cell through one relationship',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(
					[
						['a', 'b'],
						['c', 'd'],
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 2,
						fill: { type: 'image', image: { data: PNG_1x1 } },
					}
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const embeds = [...xml.matchAll(/<a:blip r:embed="(rId\d+)"/g)].map((m) => m[1])
			assert(embeds.length === 4, `expected all 4 cells filled; got ${embeds.length}`)
			assert(new Set(embeds).size === 1, `expected one shared rel; got ${JSON.stringify(embeds)}`)
			const media = (await listEntries(zip)).filter((e) => e.startsWith('ppt/media/') && !e.endsWith('/'))
			assert(media.length === 1, `expected one embedded media file; got ${JSON.stringify(media)}`)
		},
	},
	{
		name: '`headerRow` and `columns` image fills each register exactly one relationship',
		fn: async () => {
			// Both sugars spread their options shallowly, so one fill object is shared by every
			// cell it styles; registration is keyed on that identity, not per cell.
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(
					[
						['h1', 'h2'],
						['b1', 'b2'],
						['b3', 'b4'],
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						headerRow: { fill: { type: 'image', image: { data: PNG_1x1 } } },
						columns: [{}, { fill: { image: { path: 'demos/common/images/cc_logo.jpg' } } }],
					}
				)
			})
			const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			const imageRels = [...rels.matchAll(/<Relationship [^>]*\/relationships\/image[^>]*>/g)]
			assert(imageRels.length === 2, `expected one rel per distinct fill object; got ${imageRels.length}`)
			const media = (await listEntries(zip)).filter((e) => e.startsWith('ppt/media/') && !e.endsWith('/'))
			assert(media.length === 2, `expected two distinct media files; got ${JSON.stringify(media)}`)
		},
	},
	{
		name: 'a merged region with an image fill covers the whole span',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(
					[
						[
							{
								text: 'merged',
								options: { colspan: 2, rowspan: 2, fill: { type: 'image', image: { data: PNG_1x1 } } },
							},
							{ text: 'c' },
						],
						[{ text: 'd' }],
					],
					{ x: 1, y: 1, w: 6, h: 2 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// The origin declares the span...
			assertIncludes(xml, 'gridSpan="2"', 'colspan emitted')
			assertIncludes(xml, 'rowSpan="2"', 'rowspan emitted')
			// ...and every cell of the region resolves to the same relationship, so no covered
			// cell can emit a dangling `r:embed`.
			const embeds = [...xml.matchAll(/<a:blip r:embed="(rId\d+)"/g)].map((m) => m[1])
			assert(embeds.length > 1, `expected the covered cells to inherit the fill; got ${embeds.length}`)
			assert(new Set(embeds).size === 1, `expected one shared rel; got ${JSON.stringify(embeds)}`)
		},
	},
	{
		name: 'an auto-paged table registers each overflow slide’s fill on that slide',
		fn: async () => {
			// Registration runs after the auto-pager has shredded the rows, mirroring
			// `createHyperlinkRels` — otherwise every relationship would pile onto slide 1
			// and the overflow slides would emit a dangling `r:embed`.
			const rows = Array.from({ length: 60 }, (_, i) => [
				{ text: `row ${i}`, options: { fill: { type: 'image', image: { data: PNG_1x1 } } } },
			])
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 6, h: 1, autoPage: true })
			})
			const slides = (await listEntries(zip)).filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e))
			assert(slides.length > 1, `expected the table to auto-page; got ${slides.length} slide(s)`)
			for (const slidePath of slides) {
				const n = /slide(\d+)\.xml$/.exec(slidePath)[1]
				const xml = await readEntry(zip, slidePath)
				const rels = await readEntry(zip, `ppt/slides/_rels/slide${n}.xml.rels`)
				for (const [, rid] of xml.matchAll(/<a:blip r:embed="(rId\d+)"/g)) {
					assertIncludes(rels, `Id="${rid}"`, `slide ${n} resolves ${rid} on its own rels part`)
				}
			}
		},
	},
	{
		name: 'a cell image fill with neither path nor data warns and emits no blipFill',
		fn: async () => {
			let zip
			const warnings = await captureWarnings(async () => {
				;({ zip } = await build((p) => {
					p.addSlide().addTable([[{ text: 'a', options: { fill: { type: 'image', image: {} } } }]], {
						x: 1,
						y: 1,
						w: 3,
						h: 1,
					})
				}))
			})
			assert(
				warnings.some((w) => w.includes('image fill requires')),
				`expected a warning; got ${JSON.stringify(warnings)}`
			)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNotIncludes(xml, '<a:blipFill', 'no blip fill when the cell image source is missing')
			assertNotIncludes(xml, 'r:embed', 'and no dangling relationship reference')
		},
	},
	{
		name: 'an SVG cell image fill is rejected with a warning (raster only)',
		fn: async () => {
			let zip
			const warnings = await captureWarnings(async () => {
				;({ zip } = await build((p) => {
					p.addSlide().addTable(
						[[{ text: 'a', options: { fill: { type: 'image', image: { data: 'image/svg+xml;base64,PHN2Zy8+' } } } }]],
						{ x: 1, y: 1, w: 3, h: 1 }
					)
				}))
			})
			assert(
				warnings.some((w) => w.includes('SVG image fills are not supported')),
				`expected an SVG warning; got ${JSON.stringify(warnings)}`
			)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNotIncludes(xml, '<a:blipFill', 'no blip fill for unsupported SVG cell source')
		},
	},
])
