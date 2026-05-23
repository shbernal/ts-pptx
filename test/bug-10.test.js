'use strict'

const { build, readEntry, assert } = require('./helpers')

module.exports = [
	{
		name: 'B10: addShape("oval", ...) emits prst="ellipse" (not invalid "oval")',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addShape('oval', { x: 1, y: 1, w: 0.4, h: 0.4, fill: { color: '00B0B9' } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="ellipse"/.test(xml),
				'expected prstGeom prst="ellipse" in slide1.xml; got: ' + xml)
			assert(!/<a:prstGeom\s+prst="oval"/.test(xml),
				'invalid prst="oval" still present in slide1.xml')
		}
	},
	{
		name: 'B11: addShape("roundedRectangle", ...) emits prst="roundRect" (not invalid "roundedRectangle")',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addShape('roundedRectangle', { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="roundRect"/.test(xml),
				'expected prstGeom prst="roundRect" in slide1.xml; got: ' + xml)
			assert(!/<a:prstGeom\s+prst="roundedRectangle"/.test(xml),
				'invalid prst="roundedRectangle" still present in slide1.xml')
		}
	},
	{
		name: 'B11: addShape("rectangle", ...) emits prst="rect" (not invalid "rectangle")',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addShape('rectangle', { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="rect"/.test(xml),
				'expected prstGeom prst="rect" in slide1.xml; got: ' + xml)
			assert(!/<a:prstGeom\s+prst="rectangle"/.test(xml),
				'invalid prst="rectangle" still present in slide1.xml')
		}
	},
	{
		name: 'B10: enum-constant API still works (pres.shapes.OVAL -> ellipse)',
		fn: async () => {
			const { zip, pres } = await build(p => {
				const s = p.addSlide()
				s.addShape(p.shapes.OVAL, { x: 1, y: 1, w: 0.4, h: 0.4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="ellipse"/.test(xml),
				'expected prstGeom prst="ellipse" via pres.shapes.OVAL; got: ' + xml)
			void pres
		}
	}
]
