'use strict'

const JSZip = require('jszip')
const PptxGenJS = require('../src/bld/pptxgen.cjs.js')
const { assert } = require('./helpers')

async function buildSlide1(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

module.exports = [
	{
		name: 'B7: shape fill with "#FF0000" emits val="FF0000" (no leading hash)',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1, y: 1, w: 2, h: 1,
				fill: { color: '#FF0000' }
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="#') === -1,
				'expected no <a:srgbClr val="#..."> in slide XML; emitted XML still contains a hash-prefixed color.\n' + xml)
			assert(xml.indexOf('<a:srgbClr val="FF0000"') !== -1,
				'expected <a:srgbClr val="FF0000"> in slide XML; got:\n' + xml)
		}
	},
	{
		name: 'B7: text color "#00FF00" emits val="00FF00"',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addText('hello', { x: 1, y: 1, w: 4, h: 0.5, color: '#00FF00' })

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="#') === -1,
				'expected no <a:srgbClr val="#..."> in slide XML; got:\n' + xml)
			assert(xml.indexOf('<a:srgbClr val="00FF00"') !== -1,
				'expected <a:srgbClr val="00FF00"> in slide XML; got:\n' + xml)
		}
	},
	{
		name: 'B7: shape line color "#0000FF" emits val="0000FF"',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1, y: 1, w: 2, h: 1,
				line: { color: '#0000FF', width: 2 }
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="#') === -1,
				'expected no <a:srgbClr val="#..."> in slide XML; got:\n' + xml)
			assert(xml.indexOf('<a:srgbClr val="0000FF"') !== -1,
				'expected <a:srgbClr val="0000FF"> in slide XML; got:\n' + xml)
		}
	},
	{
		name: 'B7: shadow color "#888888" emits val="888888"',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1, y: 1, w: 2, h: 1,
				fill: { color: 'CCCCCC' },
				shadow: { type: 'outer', color: '#888888', blur: 6, offset: 2, opacity: 0.5 }
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="#') === -1,
				'expected no <a:srgbClr val="#..."> in slide XML (shadow); got:\n' + xml)
			// shadow path emits <a:srgbClr val="..."> as part of <a:effectLst>
			const shadowIdx = xml.indexOf('<a:effectLst>')
			assert(shadowIdx !== -1, 'expected <a:effectLst> in slide XML; got:\n' + xml)
			const shadowBlock = xml.substring(shadowIdx, shadowIdx + 400)
			assert(shadowBlock.indexOf('<a:srgbClr val="888888"') !== -1,
				'expected shadow color "888888" inside <a:effectLst>; got:\n' + shadowBlock)
		}
	}
]
