// The read model's contract is that `null` means "absent". Two accessors broke it in opposite
// directions: `Picture.imagePartName` threw where the same blip reached through a picture fill
// returned `null`, and `carriedDecorations` reported a `p:spTree/p:extLst` as a shape to carry.
// Neither construct is in the fixture corpus, so both inputs are authored here.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const SLIDE_PATH = 'ppt/slides/slide1.xml'
const SLIDE_RELS_PATH = 'ppt/slides/_rels/slide1.xml.rels'
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/**
 * The one picture on slide 1, narrowed off the shape union.
 * @param {import('../../dist/read.js').Presentation} pres
 */
function onlyPicture(pres) {
	const shape = pres.slides[0].shapes[0]
	assert(shape?.shapeType === 'picture', `expected a picture; got ${shape?.shapeType}`)
	return shape
}

/**
 * Author a one-picture deck, rewrite its slide part and rels, and load the result.
 * @param {{ editSlide?: (xml: string) => string, editRels?: (xml: string) => string }} [edits]
 */
async function pictureDeck({ editSlide, editRels } = {}) {
	const pres = new TsPptx()
	pres.addSlide().addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1 })
	const zip = await JSZip.loadAsync(await pres.toBytes())
	if (editSlide) await rewrite(zip, SLIDE_PATH, editSlide)
	if (editRels) await rewrite(zip, SLIDE_RELS_PATH, editRels)
	return Presentation.load(await zip.generateAsync({ type: 'nodebuffer' }))
}

/**
 * Apply `edit` to one zip entry in place, refusing a rewrite that changes nothing — a
 * `String.replace` that silently matched nothing would leave the case asserting on a CLEAN deck.
 * @param {JSZip} zip
 * @param {string} path
 * @param {(xml: string) => string} edit
 */
async function rewrite(zip, path, edit) {
	const entry = zip.file(path)
	assert(entry, `no ${path} in the authored package`)
	const before = await entry.async('string')
	const after = edit(before)
	assert(after !== before, `the rewrite of ${path} must actually change it`)
	zip.file(path, after)
}

describe('Picture partnames degrade like their picture-fill twin', () => {
	test('an intact embed still resolves to a partname', async () => {
		const picture = onlyPicture(await pictureDeck())
		assert(
			(picture.imagePartName ?? '').startsWith('/ppt/media/'),
			`expected a media partname; got ${picture.imagePartName}`
		)
		assertEqual(picture.mediaKind, 'raster', 'a plain raster picture')
	})

	test('a dangling r:embed reports null rather than throwing', async () => {
		const picture = onlyPicture(
			await pictureDeck({ editSlide: (xml) => xml.replace(/r:embed="[^"]*"/, 'r:embed="rIdGone"') })
		)
		assertEqual(picture.imagePartName, null, 'imagePartName')
		assertEqual(picture.svgPartName, null, 'svgPartName')
		assertEqual(picture.mediaPartName, null, 'mediaPartName — the "just give me the bytes" accessor')
	})

	test('a LINKED image reports null rather than throwing, and mediaKind still names it', async () => {
		// `mediaKind`'s own documentation calls out "a linked image" as the `'none'` case, and the
		// accessor beside it threw on exactly that.
		const picture = onlyPicture(
			await pictureDeck({
				editRels: (xml) =>
					xml.replace(
						/(<Relationship Id="[^"]*"[^>]*image[^>]*Target=")[^"]*("[^>]*)\/>/,
						'$1https://example.com/logo.png$2 TargetMode="External"/>'
					),
			})
		)
		assertEqual(picture.imagePartName, null, 'an external target names no part')
		assertEqual(picture.mediaPartName, null, 'and neither does the bytes accessor')
	})

	test('a whole shape walk survives one broken embed', async () => {
		const pres = await pictureDeck({ editSlide: (xml) => xml.replace(/r:embed="[^"]*"/, 'r:embed="rIdGone"') })
		const names = pres.slides[0].shapes.map((shape) =>
			shape.shapeType === 'picture' ? (shape.mediaPartName ?? '(none)') : shape.shapeType
		)
		assertEqual(names.join(','), '(none)', 'the walk completes instead of taking the slide down')
	})
})

describe('A p:spTree/p:extLst is the tree own child, not a shape', () => {
	/** Author a deck whose every layout and master carries a `p:spTree/p:extLst` of its own. */
	async function deckWithSpTreeExtLst() {
		const pres = new TsPptx()
		pres.defineSlideMaster({
			title: 'CARRY',
			objects: [{ text: { text: 'decoration', options: { x: 1, y: 1, w: 3, h: 1 } } }],
		})
		pres.addSlide({ masterTitle: 'CARRY' }).addText('body', { x: 1, y: 3, w: 3, h: 1 })
		const zip = await JSZip.loadAsync(await pres.toBytes())
		let patched = 0
		for (const name of Object.keys(zip.files)) {
			if (!/^ppt\/(slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue
			await rewrite(zip, name, (xml) =>
				xml.replace(
					'</p:spTree>',
					'<p:extLst><p:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"/></p:extLst></p:spTree>'
				)
			)
			patched++
		}
		assert(patched > 0, 'the layout/master rewrite must land somewhere')
		return { bytes: await zip.generateAsync({ type: 'nodebuffer' }) }
	}

	test('an importSlide carrying master graphics does not carry the extLst into the shape tree', async () => {
		const { bytes } = await deckWithSpTreeExtLst()
		const source = await Presentation.load(bytes)
		const target = await Presentation.load(await new TsPptx().toBytes())
		target.importSlide(source, 0, { theme: 'preserve', carryMasterGraphics: true })
		const saved = await target.save()
		const zip = await JSZip.loadAsync(saved)
		// Sensitivity: the carry has to have RUN, or the assertion below passes on an empty walk.
		const slideXml = (await zip.file(SLIDE_PATH)?.async('string')) ?? ''
		assert(slideXml.includes('decoration'), `expected the layout decoration to be carried; got ${slideXml}`)
		for (const name of Object.keys(zip.files)) {
			if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue
			const xml = (await zip.file(name)?.async('string')) ?? ''
			const spTree = xml.match(/<p:spTree>[\s\S]*<\/p:spTree>/)?.[0]
			if (!spTree) continue
			const extAt = spTree.indexOf('<p:extLst>')
			if (extAt === -1) continue
			// `CT_GroupShape` sequences `nvGrpSpPr, grpSpPr, (shape)*, extLst?`, so a tree-level
			// extLst may only ever be last.
			const lastShapeAt = Math.max(
				spTree.lastIndexOf('<p:sp>'),
				spTree.lastIndexOf('<p:pic>'),
				spTree.lastIndexOf('<p:grpSp>')
			)
			assert(extAt > lastShapeAt, `${name}: a spTree extLst must follow every shape; got ${spTree.slice(0, 400)}`)
		}
	})
})

describe('One reading per fact', () => {
	/**
	 * Author a one-table deck, rewrite its slide XML, and return the loaded cell (0, 0).
	 * @param {(xml: string) => string} edit
	 */
	async function cellFrom(edit) {
		const pres = new TsPptx()
		pres.addSlide().addTable([[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 6 })
		const zip = await JSZip.loadAsync(await pres.toBytes())
		await rewrite(zip, SLIDE_PATH, edit)
		const presentation = await Presentation.load(await zip.generateAsync({ type: 'nodebuffer' }))
		const shape = presentation.slides[0].shapes[0]
		assert(shape?.shapeType === 'graphicFrame', `expected a graphic frame; got ${shape?.shapeType}`)
		const table = shape.table
		assert(table, 'the graphic frame holds a table')
		return table.cell(0, 0)
	}

	test('a solid fill in a place CT_TableCell has none is not reported as the cell fill', async () => {
		// `a:tc/a:solidFill` is not a location the content model permits, so the old `?? this.tc`
		// fallback was unreachable on a well-formed deck and, on a malformed one, reported a
		// scheme token that `resolvedFill` and `hasOwnFill` both denied.
		const cell = await cellFrom((xml) =>
			xml.replace('<a:tc>', '<a:tc><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>')
		)
		assertEqual(cell.hasOwnFill, false, 'no a:tcPr fill')
		assertEqual(cell.fillSchemeColor, null, 'and no scheme token either')
	})

	test('anchorCtr reads every xsd:boolean lexical form', async () => {
		// `xsd:boolean` has four forms and PowerPoint writes only `1`/`0`, so a hand-rolled test
		// passes every deck this library and PowerPoint produce and misreads the rest.
		for (const [written, expected] of [
			['1', true],
			['true', true],
			['0', false],
			['false', false],
		]) {
			const cell = await cellFrom((xml) => xml.replace('<a:tcPr', `<a:tcPr anchorCtr="${written}"`))
			assertEqual(cell.anchorCtr, expected, `anchorCtr="${written}"`)
		}
	})
})
