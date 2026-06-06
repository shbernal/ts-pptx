import JSZip from 'jszip'
import PptxGenJS from '../dist/node.js'
import { assert } from './helpers.js'

async function buildSlide1(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// Compute expected emitted alpha integer for a 2-char alpha hex (PowerPoint scale: 100000).
function alphaPct(hex) {
	return Math.round((parseInt(hex, 16) / 255) * 100000)
}

export default [
	{
		name: 'fill color "00000020" splits to val="000000" + <a:alpha val="12549"/>',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				fill: { color: '00000020' },
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="00000020"') === -1, 'expected no 8-char val="00000020" in slide XML; got:\n' + xml)
			const expectedAlpha = alphaPct('20') // 12549
			const expectedFragment = `<a:srgbClr val="000000"><a:alpha val="${expectedAlpha}"/></a:srgbClr>`
			assert(xml.indexOf(expectedFragment) !== -1, `expected ${expectedFragment} in slide XML; got:\n` + xml)
		},
	},
	{
		name: 'text color "00FF0080" splits to val="00FF00" + <a:alpha val="50196"/>',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addText('hello', { x: 1, y: 1, w: 4, h: 0.5, color: '00FF0080' })

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="00FF0080"') === -1, 'expected no 8-char val="00FF0080" in slide XML; got:\n' + xml)
			const expectedAlpha = alphaPct('80') // 50196
			const expectedFragment = `<a:srgbClr val="00FF00"><a:alpha val="${expectedAlpha}"/></a:srgbClr>`
			assert(xml.indexOf(expectedFragment) !== -1, `expected ${expectedFragment} in slide XML; got:\n` + xml)
		},
	},
	{
		name: 'shape line color "0000FF40" splits to val="0000FF" + <a:alpha val="..."/>',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				line: { color: '0000FF40', width: 2 },
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="0000FF40"') === -1, 'expected no 8-char val="0000FF40" in slide XML; got:\n' + xml)
			const expectedAlpha = alphaPct('40') // 25098
			const expectedFragment = `<a:srgbClr val="0000FF"><a:alpha val="${expectedAlpha}"/></a:srgbClr>`
			assert(xml.indexOf(expectedFragment) !== -1, `expected ${expectedFragment} in slide XML; got:\n` + xml)
		},
	},
	{
		name: 'shadow color "00000020" without explicit opacity emits val="000000" + derived alpha',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				fill: { color: 'CCCCCC' },
				shadow: { type: 'outer', color: '00000020', blur: 6, offset: 2 },
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="00000020"') === -1, 'expected no 8-char val="00000020" in slide XML; got:\n' + xml)

			const effectIdx = xml.indexOf('<a:effectLst>')
			assert(effectIdx !== -1, 'expected <a:effectLst> in slide XML; got:\n' + xml)
			const effectClose = xml.indexOf('</a:effectLst>', effectIdx)
			assert(effectClose !== -1, 'expected </a:effectLst> in slide XML; got:\n' + xml)
			const effectBlock = xml.slice(effectIdx, effectClose)

			assert(
				effectBlock.indexOf('<a:srgbClr val="000000">') !== -1,
				'expected <a:srgbClr val="000000"> inside shadow effectLst; got:\n' + effectBlock
			)
			// Derived opacity: 32/255 ≈ 0.12549 → ×100000 ≈ 12549
			const expectedAlpha = alphaPct('20') // 12549
			assert(
				effectBlock.indexOf(`<a:alpha val="${expectedAlpha}"/>`) !== -1,
				`expected <a:alpha val="${expectedAlpha}"/> inside shadow effectLst; got:\n` + effectBlock
			)
		},
	},
	{
		name: 'shadow color "88888880" with explicit opacity=0.5 — explicit opacity wins, color stripped',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				fill: { color: 'CCCCCC' },
				shadow: { type: 'outer', color: '88888880', blur: 6, offset: 2, opacity: 0.5 },
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="88888880"') === -1, 'expected no 8-char val="88888880" in slide XML; got:\n' + xml)

			const effectIdx = xml.indexOf('<a:effectLst>')
			assert(effectIdx !== -1, 'expected <a:effectLst> in slide XML; got:\n' + xml)
			const effectClose = xml.indexOf('</a:effectLst>', effectIdx)
			const effectBlock = xml.slice(effectIdx, effectClose)

			assert(
				effectBlock.indexOf('<a:srgbClr val="888888">') !== -1,
				'expected <a:srgbClr val="888888"> inside shadow effectLst; got:\n' + effectBlock
			)
			// Explicit opacity wins: 0.5 → 50000 (NOT 50196 derived from 0x80/255)
			assert(
				effectBlock.indexOf('<a:alpha val="50000"/>') !== -1,
				'expected <a:alpha val="50000"/> (explicit) inside shadow effectLst; got:\n' + effectBlock
			)
		},
	},
	{
		name: 'hash-prefixed 8-char "#FF0000FF" splits to val="FF0000" + <a:alpha val="100000"/>',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				fill: { color: '#FF0000FF' },
			})

			const xml = await buildSlide1(pres)
			assert(xml.indexOf('val="#') === -1, 'expected no <a:srgbClr val="#..."> in slide XML; got:\n' + xml)
			assert(xml.indexOf('val="FF0000FF"') === -1, 'expected no 8-char val="FF0000FF" in slide XML; got:\n' + xml)
			const expectedFragment = '<a:srgbClr val="FF0000"><a:alpha val="100000"/></a:srgbClr>'
			assert(xml.indexOf(expectedFragment) !== -1, `expected ${expectedFragment} in slide XML; got:\n` + xml)
		},
	},
]
