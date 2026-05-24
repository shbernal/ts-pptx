'use strict'

const { build, readEntry, assert } = require('./helpers')

// Acceptance: emitted <a:tcPr> must never carry NaN in marL/R/T/B even when
// the user supplies a non-numeric/non-array `margin` (string, plain object,
// undefined slot, etc).  Falls back to defaults.

module.exports = [
	{
		name: 'cell margin set to a string falls back to defaults (no NaN)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addTable([[{ text: 'a', options: { margin: 'foo' } }]], { x: 1, y: 1, w: 4, colW: [4] })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('NaN') === -1, 'NaN must not appear in slide1.xml; got: ' + xml)
			// And the cell should still get numeric mar attributes.
			assert(/<a:tcPr[^>]*marL="\d+"[^>]*marR="\d+"[^>]*marT="\d+"[^>]*marB="\d+"/.test(xml),
				'expected numeric marL/R/T/B on <a:tcPr>; got: ' + xml)
		}
	},
	{
		name: 'cell margin set to a plain object falls back to defaults (no NaN)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addTable([[{ text: 'b', options: { margin: { top: 5 } } }]], { x: 1, y: 1, w: 4, colW: [4] })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('NaN') === -1, 'NaN must not appear in slide1.xml; got: ' + xml)
			assert(/<a:tcPr[^>]*marL="\d+"[^>]*marR="\d+"[^>]*marT="\d+"[^>]*marB="\d+"/.test(xml),
				'expected numeric marL/R/T/B on <a:tcPr>; got: ' + xml)
		}
	},
	{
		name: 'table-level margin set to a string falls back to defaults (no NaN)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addTable([[{ text: 'c' }]], { x: 1, y: 1, w: 4, colW: [4], margin: 'bad' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('NaN') === -1, 'NaN must not appear in slide1.xml; got: ' + xml)
			assert(/<a:tcPr[^>]*marL="\d+"[^>]*marR="\d+"[^>]*marT="\d+"[^>]*marB="\d+"/.test(xml),
				'expected numeric marL/R/T/B on <a:tcPr>; got: ' + xml)
		}
	},
	{
		name: 'numeric margin still emits valid mar attributes (regression guard)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addTable([[{ text: 'd', options: { margin: 0.05 } }]], { x: 1, y: 1, w: 4, colW: [4] })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('NaN') === -1, 'NaN must not appear in slide1.xml; got: ' + xml)
			assert(/<a:tcPr[^>]*marL="\d+"[^>]*marR="\d+"[^>]*marT="\d+"[^>]*marB="\d+"/.test(xml),
				'expected numeric marL/R/T/B on <a:tcPr>; got: ' + xml)
		}
	}
]
