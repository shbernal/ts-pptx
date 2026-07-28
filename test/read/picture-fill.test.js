// Picture (image) fills of a *surface* — `a:blipFill` inside `a:tcPr`, `p:spPr`,
// and `p:bgPr` — read through the shared decoder in src/read/api/picture-fill.ts.
//
// A picture fill is not a colour, so `TableCell.resolvedFill` /
// `AutoShape.resolvedFill` report `null` for one: before this, an image-filled
// cell was indistinguishable from an empty cell. Both halves are gated on
// PowerPoint-authored oracles rather than synthesized XML:
//
//   - `table-cell-image-fill.pptx` — a 4x2 COM-authored table ("No Style, No
//     Grid", so every fill in it was set explicitly): stretched cells, a tiled
//     cell, a merged origin, and one solid-filled control cell.
//   - `math-omml.pptx` — its `mc:Fallback` `equation-box` is a `p:sp` whose
//     `p:spPr` carries a blipFill with a *negative* `a:fillRect` bleed, which is
//     the shape-side case no writer of ours produces. It is the only genuine
//     `p:spPr/a:blipFill` in the fixture set, and the read model walks `mc:Choice`
//     only, so the test unwraps the AlternateContent (see below) to reach it.
//
// The `resolvedFill` precedence leg has no PowerPoint oracle (the fixture's table
// names a style that shades nothing), so it is authored with the write API: a
// header-shaded `tableStyle` plus an image-filled cell, proving the cell's own
// blipFill suppresses the style graph the way PowerPoint renders it.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead, firstShape, firstTable, schemaErrors, validatorInstalled } from './authored.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

// A 1x1 transparent PNG; the fill geometry lives on the blipFill, not the pixels.
const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

async function openFixture(name) {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', `${name}.pptx`)))
}

function tableOf(presentation) {
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) {
			if (shape.shapeType === 'graphicFrame' && shape.table) return shape.table
		}
	}
	return null
}

describe('TableCell.pictureFill — a:tcPr/a:blipFill (PowerPoint oracle)', () => {
	test('a stretched picture cell reads its image, mode, and stretch rect', async () => {
		const table = tableOf(await openFixture('table-cell-image-fill'))
		assert(table, 'the fixture table is read back')

		const fill = table.cell(0, 0).pictureFill
		assert(fill, 'A1 surfaces a picture fill')
		assertEqual(fill.relId, 'rId2', "the blip's r:embed, verbatim")
		assertEqual(fill.partName, '/ppt/media/image1.jpg', 'and resolves against the slide relationships')
		assertEqual(fill.mode, 'stretch', 'a:stretch → stretch')
		assert(fill.fillRect, 'the a:stretch/a:fillRect is present')
		assertEqual(fill.fillRect.left, 0, 'an empty fillRect reads zeros, not nulls')
		assertEqual(fill.tile, null, 'a stretched fill has no tile geometry')
		assertEqual(fill.srcRect, null, 'PowerPoint omits a:srcRect on its stretched cells')
		assertEqual(fill.alpha, null, 'no a:alphaModFix on the blip')
		assertEqual(fill.dpi, null, 'the stretched cells carry a bare <a:blipFill>')
		assertEqual(fill.rotWithShape, null, 'and so no @rotWithShape either')
	})

	test('a picture cell that also carries borders reads both (the child-order case)', async () => {
		// `CT_TableCellProperties` puts EG_FillProperties at order 7, i.e. after
		// lnL/lnR/lnT/lnB — B1 in the fixture has all four plus the fill.
		const cell = tableOf(await openFixture('table-cell-image-fill')).cell(1, 0)
		const fill = cell.pictureFill
		assert(fill, 'the fill is found past the four preceding border elements')
		assertEqual(fill.relId, 'rId2', 'and shares the deduped relationship')
		const borders = cell.borders
		assert(borders?.left && borders.right && borders.top && borders.bottom, 'all four borders still read')
	})

	test('a tiled picture cell reads its tile offsets, scales, flip and alignment', async () => {
		const fill = tableOf(await openFixture('table-cell-image-fill')).cell(3, 0).pictureFill
		assert(fill, 'D1 surfaces a picture fill')
		assertEqual(fill.mode, 'tile', 'a:tile → tile')
		assert(fill.tile, 'the tile geometry is decoded')
		assertEqual(fill.tile.offsetXEmu, 0, 'tx stays in EMU')
		assertEqual(fill.tile.offsetYEmu, 0, 'ty stays in EMU')
		assertEqual(fill.tile.scaleX, 1, 'sx=100000 → the 1 = 100 % fraction convention')
		assertEqual(fill.tile.scaleY, 1, 'sy likewise')
		assertEqual(fill.tile.flip, 'xy', 'the @flip token is kept verbatim')
		assertEqual(fill.tile.align, 'tl', 'and the @algn token')
		assertEqual(fill.fillRect, null, 'a tiled fill has no stretch rect')
		// This is the cell PowerPoint wrote with the same attribute pair our writer emits.
		assertEqual(fill.dpi, 0, '@dpi="0" means "use the image\'s own"')
		assertEqual(fill.rotWithShape, true, '@rotWithShape="1" decodes to a boolean')
		assert(fill.srcRect, 'the explicit <a:srcRect/> is reported')
		assertEqual(fill.srcRect.bottom, 0, 'an empty srcRect is an uncropped source, not an absent one')
	})

	test('a merged origin carries the fill and its covered cell reports none', async () => {
		const table = tableOf(await openFixture('table-cell-image-fill'))
		const origin = table.cell(2, 0)
		assertEqual(origin.gridSpan, 2, 'C1 is the merge origin')
		assert(origin.pictureFill, 'the origin holds the picture fill')
		const covered = table.cell(2, 1)
		assertEqual(covered.isMergeContinuation, true, 'C2 is the covered half')
		assertEqual(covered.pictureFill, null, 'PowerPoint writes a bare <a:tcPr/> there, so nothing to read')
	})

	test('non-picture cells report null, and a solid-filled cell still resolves its colour', async () => {
		const table = tableOf(await openFixture('table-cell-image-fill'))
		assertEqual(table.cell(0, 1).pictureFill, null, 'the solid control cell has no picture fill')
		assertEqual(table.cell(0, 1).resolvedFill?.effectiveHex, 'FF0000', 'and its solid fill still reads')
		assertEqual(table.cell(3, 1).pictureFill, null, 'a bare <a:tcPr/> cell has no picture fill')
		assertEqual(table.cell(1, 1).pictureFill, null, 'a borders-only cell has no picture fill')
	})

	test('every picture cell resolves to the one shared media part', async () => {
		const table = tableOf(await openFixture('table-cell-image-fill'))
		const parts = [table.cell(0, 0), table.cell(1, 0), table.cell(2, 0), table.cell(3, 0)].map(
			(cell) => cell.pictureFill?.partName
		)
		assertEqual(new Set(parts).size, 1, 'PowerPoint dedupes the source to one relationship')
		assertEqual(parts[0], '/ppt/media/image1.jpg', 'and it resolves to the embedded jpg')
	})
})

/**
 * The genuine PowerPoint `p:sp` whose `p:spPr` carries a blipFill lives in
 * `math-omml.pptx`'s `mc:Fallback` branch, and the read model walks `mc:Choice`
 * only — so the shape it enumerates is the a14 equation copy, which is
 * `a:noFill`. Unwrap the AlternateContent to its Fallback child so that PowerPoint
 * `p:sp` (and its untouched `rId2` → media relationship) becomes reachable.
 *
 * Only the `mc:` wrapper is removed; the shape XML and the part relationships are
 * PowerPoint's own bytes, which is what makes this an oracle rather than a
 * synthetic input. The wrapper is not what is under test — the reader never
 * traverses it either way.
 */
async function unwrappedFallbackShape() {
	const buf = await readFile(path.join(__dirname, 'fixtures', 'math-omml.pptx'))
	const zip = await JSZip.loadAsync(buf)
	const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
	const unwrapped = slideXml.replace(
		/<mc:AlternateContent[^>]*>[\s\S]*?<mc:Fallback>([\s\S]*?)<\/mc:Fallback><\/mc:AlternateContent>/,
		'$1'
	)
	assert(unwrapped !== slideXml, 'the AlternateContent wrapper was found and removed')
	zip.file('ppt/slides/slide1.xml', unwrapped)
	const presentation = await Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
	return firstShape(presentation, (s) => s.name === 'equation-box')
}

describe('AutoShape.pictureFill — p:spPr/a:blipFill (PowerPoint oracle)', () => {
	test("an image-filled shape surface reads its blip and the fillRect's negative bleed", async () => {
		const shape = await unwrappedFallbackShape()
		assert(shape, 'the PowerPoint-authored equation shape is read back')
		assertEqual(shape.resolvedFill, null, 'resolvedFill decodes solid colours only, so it reports nothing')

		const fill = shape.pictureFill
		assert(fill, 'the shape surface surfaces a picture fill')
		assertEqual(fill.relId, 'rId2', "the blip's r:embed")
		assert(fill.partName?.startsWith('/ppt/media/'), `resolves to a media part, got ${fill.partName}`)
		assertEqual(fill.mode, 'stretch', 'a:stretch → stretch')
		// A fillRect edge may be negative: the image bleeds past that edge rather than
		// being inset by it, so the sign has to survive the fraction conversion.
		assertEqual(fill.fillRect?.bottom, -0.06667, 'b="-6667" → -0.06667, sign intact')
		assertEqual(fill.fillRect?.left, 0, 'the unset edges read 0')
	})

	test('a shape with no blipFill reports null', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('plain', { x: 1, y: 1, w: 3, h: 1, fill: { color: '1F4E79' } })
		})
		const shape = firstShape(presentation, (s) => s.shapeType === 'autoShape')
		assertEqual(shape.pictureFill, null, 'a solid-filled shape has no picture fill')
		assertEqual(shape.resolvedFill?.effectiveHex, '1F4E79', 'and its solid fill is untouched')
	})

	test('a shape image fill authored by the write API reads back', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('img', { x: 1, y: 1, w: 3, h: 1, fill: { type: 'image', image: { data: PNG_1x1 } } })
		})
		const shape = firstShape(presentation, (s) => s.shapeType === 'autoShape')
		const fill = shape.pictureFill
		assert(fill, "the writer's blipFill decodes through the same reader")
		assert(fill.partName?.endsWith('.png'), `resolves to the embedded png, got ${fill.partName}`)
		assertEqual(fill.mode, 'stretch', 'the writer always stretches')
		assertEqual(fill.dpi, 0, 'and spells dpi="0"')
		assertEqual(fill.rotWithShape, true, 'and rotWithShape="1"')
		assertEqual(fill.alpha, null, 'an opaque fill records no a:alphaModFix')
	})

	test('a transparent image fill reads its alpha as an opacity fraction', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('img', {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: { type: 'image', image: { data: PNG_1x1 }, transparency: 25 },
			})
		})
		const fill = firstShape(presentation, (s) => s.shapeType === 'autoShape').pictureFill
		// The writer's `transparency` is a percentage *lost*; a:alphaModFix/@amt is the
		// opacity that remains, so 25 % transparent is 0.75 opaque.
		assertEqual(fill.alpha, 0.75, 'transparency 25 → a:alphaModFix amt=75000 → 0.75')
	})

	test('a blip pointing at a missing relationship reads its rel id and a null part', async () => {
		// A reader must not throw on a dangling r:embed: `Relationships.resolveTarget`
		// does, so the decoder checks the id exists before resolving it.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addText('img', { x: 1, y: 1, w: 3, h: 1, fill: { type: 'image', image: { data: PNG_1x1 } } })
		})
		const zip = await JSZip.loadAsync(buf)
		const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
		zip.file('ppt/slides/slide1.xml', slideXml.replace(/<a:blip r:embed="rId\d+"/, '<a:blip r:embed="rIdNope"'))
		const presentation = await Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
		const fill = firstShape(presentation, (s) => s.shapeType === 'autoShape').pictureFill
		assertEqual(fill.relId, 'rIdNope', 'the id is reported verbatim')
		assertEqual(fill.partName, null, 'and resolves to nothing rather than throwing')
	})
})

describe('TableCell.resolvedFill — a non-solid own fill suppresses the style graph', () => {
	/**
	 * A 1x2 table on a header-shading style: cell A1 image-filled, B1 left to the style.
	 *
	 * The style has to be a *custom* one: `makeXmlTableStyles` materialises only
	 * registered styles into `ppt/tableStyles.xml`, so a built-in `TableStyle.*` GUID
	 * writes no `a:tblStyle` for `Table.resolvedStyle` to resolve and the style-graph
	 * fallback this test is about would never fire in either direction.
	 */
	function styledTable(pres) {
		const brand = pres.defineTableStyle({ name: 'Picture Fill Probe', firstRow: { fill: '1A2B3C', color: 'FFFFFF' } })
		pres
			.addSlide()
			.addTable([[{ text: 'A', options: { fill: { type: 'image', image: { data: PNG_1x1 } } } }, { text: 'B' }]], {
				x: 1,
				y: 1,
				w: 6,
				h: 1,
				tableStyle: brand,
				hasHeader: true,
			})
	}

	test('an image-filled cell reports no colour while its plain neighbour inherits one', async () => {
		const { presentation } = await authorRead(styledTable)
		const table = firstTable(presentation)
		assert(table.resolvedStyle, 'the authored tableStyle resolves, so the style graph is live')

		// The control: same row, same style, no fill of its own — it must still inherit.
		const plain = table.cell(0, 1)
		assertEqual(plain.pictureFill, null, 'the neighbour is not image-filled')
		assert(plain.resolvedFill?.effectiveHex, 'and it inherits the firstRow shading from the style')

		const picture = table.cell(0, 0)
		assert(picture.pictureFill, 'the image-filled cell surfaces its picture')
		assertEqual(
			picture.resolvedFill,
			null,
			'and reports no colour: its own a:blipFill overrides the style graph, so reporting the header shading would be a colour PowerPoint never paints'
		)
	})

	test('a cell with no fill choice at all still falls through to the style', async () => {
		const { presentation } = await authorRead((pres) => {
			const brand = pres.defineTableStyle({ name: 'Header Only', firstRow: { fill: '1A2B3C' } })
			pres
				.addSlide()
				.addTable([[{ text: 'x' }, { text: 'y' }]], { x: 1, y: 1, w: 6, tableStyle: brand, hasHeader: true })
		})
		const cell = firstTable(presentation).cell(0, 0)
		assertEqual(cell.pictureFill, null, 'no picture fill')
		assert(cell.resolvedFill?.effectiveHex, 'the style-graph fallback is intact for the unfilled case')
	})

	test.skipIf(!validatorInstalled)('the authored image-filled styled table is schema-valid', async () => {
		const { buf } = await authorRead(styledTable)
		assertEqual((await schemaErrors(buf)).length, 0, 'no schema violations')
	})
})

describe('slide background — the image variant carries the full picture fill', () => {
	test('an authored image background reports geometry alongside relId/partName', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().background = { data: PNG_1x1 }
		})
		const bg = presentation.slides[0].background
		assert(bg.type === 'image', 'image background')
		assertEqual(bg.picture.relId, bg.relId, 'the flat relId still mirrors the picture fill')
		assertEqual(bg.picture.partName, bg.partName, 'and so does partName')
		assertEqual(bg.picture.mode, 'stretch', 'the geometry the flat fields never carried')
		assertEqual(bg.picture.tile, null, 'not tiled')
	})
})
