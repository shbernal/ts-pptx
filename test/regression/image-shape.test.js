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
])
