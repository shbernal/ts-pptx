import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

defineRegressionSuite('Image shape clipping', [
	{
		name: 'addImage({ rounding: true }) clips to an ellipse',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 2, h: 2, rounding: true })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="ellipse"/.test(xml), 'expected prstGeom prst="ellipse"; got: ' + xml)
		},
	},
	{
		name: 'addImage with no shape/rounding stays rect',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 2, h: 2 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="rect"/.test(xml), 'expected prstGeom prst="rect"; got: ' + xml)
		},
	},
	{
		name: 'addImage({ shape: "hexagon" }) clips to that preset',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 2, h: 2, shape: 'hexagon' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="hexagon"/.test(xml), 'expected prstGeom prst="hexagon"; got: ' + xml)
		},
	},
	{
		name: 'addImage shape takes precedence over rounding',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 2, h: 2, shape: 'roundRect', rounding: true, rectRadius: 0.25 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="roundRect"/.test(xml), 'expected prstGeom prst="roundRect"; got: ' + xml)
			assert(!/<a:prstGeom\s+prst="ellipse"/.test(xml), 'rounding should not override shape; got: ' + xml)
			assert(/<a:gd\s+name="adj"\s+fmla="val \d+"/.test(xml), 'expected rectRadius adjust value; got: ' + xml)
		},
	},
	{
		name: 'addImage({ shape, sizing:cover }) emits both srcRect and the preset geometry',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 2, h: 2, shape: 'ellipse', sizing: { type: 'cover', w: 2, h: 2 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:srcRect\b/.test(xml), 'expected srcRect from sizing:cover; got: ' + xml)
			assert(/<a:prstGeom\s+prst="ellipse"/.test(xml), 'expected prstGeom prst="ellipse"; got: ' + xml)
		},
	},
	{
		name: 'addImage({ points }) clips to a freeform custGeom path',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: PNG_DATA,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					points: [{ x: 1, y: 0 }, { x: 2, y: 2 }, { x: 0, y: 2 }, { close: true }],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:custGeom>/.test(xml), 'expected custGeom on the picture; got: ' + xml)
			assert(/<a:moveTo>/.test(xml), 'expected a moveTo; got: ' + xml)
			assert(/<a:lnTo>/.test(xml), 'expected a lnTo; got: ' + xml)
			assert(/<a:close \/>/.test(xml), 'expected a close; got: ' + xml)
			assert(!/<a:prstGeom\b/.test(xml), 'points should suppress prstGeom; got: ' + xml)
		},
	},
	{
		// The "image embedded in a shape" composition: a freeform custGeom clip (in spPr)
		// AND a source crop (srcRect in blipFill) on the same picture — exactly what a
		// PowerPoint picture placeholder produces. The two are written into different
		// children of <p:pic>, so they must both appear and in OOXML document order
		// (CT_Picture: blipFill before spPr).
		name: 'addImage({ points, sizing:cover }) emits srcRect (blipFill) + custGeom (spPr) in order',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: PNG_DATA,
					x: 1,
					y: 1,
					w: 2,
					h: 3,
					points: [{ x: 1, y: 0 }, { x: 2, y: 0 }, { x: 2, y: 3 }, { x: 0, y: 3 }, { close: true }],
					sizing: { type: 'cover', w: 2, h: 3 },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:srcRect\b/.test(xml), 'expected srcRect from sizing:cover; got: ' + xml)
			assert(/<a:custGeom>/.test(xml), 'expected custGeom from points; got: ' + xml)
			// cover/contain/crop now emit an explicit <a:fillRect/> inside <a:stretch> (PowerPoint's own form,
			// ECMA-376 L.4.8.4.3) so a custGeom-clipped, source-cropped pic has no empty-stretch ambiguity.
			assert(
				/<a:stretch><a:fillRect\/><\/a:stretch>/.test(xml),
				'expected explicit fillRect inside stretch; got: ' + xml
			)
			assert(!/<a:prstGeom\b/.test(xml), 'points should suppress prstGeom; got: ' + xml)
			const iSrcRect = xml.indexOf('<a:srcRect')
			const iCustGeom = xml.indexOf('<a:custGeom>')
			assert(
				iSrcRect !== -1 && iCustGeom !== -1 && iSrcRect < iCustGeom,
				'srcRect (blipFill) must precede custGeom (spPr); got: ' + xml
			)
		},
	},
	{
		// The half-disc ("D") cover is most naturally expressed with an arcTo for the curved
		// edge rather than two cubicBezTo approximations. Verify genXmlCustGeom emits a valid
		// <a:arcTo> with all four guide attributes when the path uses an arc segment.
		name: 'addImage({ points: [...arcTo] }) emits a valid arcTo segment',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: PNG_DATA,
					x: 1,
					y: 1,
					w: 2,
					h: 3,
					points: [
						{ x: 0.64, y: 0 },
						{ x: 2, y: 0 },
						{ x: 2, y: 3 },
						{ x: 0.64, y: 3 },
						{ x: 0, y: 1.5, curve: { type: 'arc', hR: 1.5, wR: 0.64, stAng: 90, swAng: 180 } },
						{ close: true },
					],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:custGeom>/.test(xml), 'expected custGeom; got: ' + xml)
			assert(
				/<a:arcTo hR="\d+" wR="\d+" stAng="-?\d+" swAng="-?\d+" \/>/.test(xml),
				'expected a valid arcTo with hR/wR/stAng/swAng; got: ' + xml
			)
		},
	},
	{
		name: 'addImage points wins over shape/rounding',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: PNG_DATA,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					shape: 'hexagon',
					rounding: true,
					points: [{ x: 0, y: 0 }, { x: 2, y: 0 }, { x: 1, y: 2 }, { close: true }],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:custGeom>/.test(xml), 'expected custGeom; got: ' + xml)
			assert(!/<a:prstGeom\b/.test(xml), 'points should win over shape/rounding; got: ' + xml)
		},
	},
])
