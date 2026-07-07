import {
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

// 1x1 transparent PNG (data URI). Used to exercise the picture-fill (`<a:blipFill>`) path
// for shapes and text boxes.
const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

async function captureWarnings(fn) {
	const orig = console.warn
	const warnings = []
	console.warn = (...args) => warnings.push(args.join(' '))
	try {
		await fn()
	} finally {
		console.warn = orig
	}
	return warnings
}

defineRegressionSuite('Image (blip) fills', [
	{
		name: 'shape image fill emits blipFill referencing an embedded media rel, before line props',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.ShapeType.rect, {
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
				s.addShape(p.ShapeType.custGeom, {
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
				s.addShape(p.ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { image: { data: PNG_1x1 } } })
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
				s.addShape(p.ShapeType.rect, {
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
				s.addShape(p.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { image: { data: PNG_1x1 } } })
				s.addShape(p.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { image: { data: PNG_1x1 } } })
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
					s.addShape(p.ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { type: 'image', image: {} } })
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
					s.addShape(p.ShapeType.rect, {
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
