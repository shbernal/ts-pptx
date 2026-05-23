'use strict'

const { build, readEntry, assert } = require('./helpers')

async function getSlide1(zip) {
	return readEntry(zip, 'ppt/slides/slide1.xml')
}

// Extract the first <a:t>...</a:t> block (text run text) from slide1.xml
function firstATText(xml) {
	const m = xml.match(/<a:t>([\s\S]*?)<\/a:t>/)
	if (!m) throw new Error('no <a:t> found in slide1.xml; xml=' + xml)
	return m[1]
}

// Extract the first <a:pPr ...>...</a:pPr> (or self-closing) block
function firstPPr(xml) {
	const m = xml.match(/<a:pPr[^>]*\/?>(?:[\s\S]*?<\/a:pPr>)?/)
	if (!m) throw new Error('no <a:pPr> found in slide1.xml; xml=' + xml)
	return m[0]
}

module.exports = [
	{
		name: 'B9: addText("• item",{bullet:true}) strips leading bullet glyph from <a:t> while keeping <a:buChar/>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('\u2022 item', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const xml = await getSlide1(zip)
			const ppr = firstPPr(xml)
			const tText = firstATText(xml)
			assert(/<a:buChar char="&#x2022;"\/>/.test(ppr),
				'expected <a:buChar char="&#x2022;"/> on <a:pPr>; got: ' + ppr)
			assert(tText === 'item',
				'expected first <a:t> to be "item" (bullet glyph stripped); got: ' + JSON.stringify(tText))
		}
	},
	{
		name: 'B9: addText("hello",{bullet:true}) leaves text unchanged',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const xml = await getSlide1(zip)
			const tText = firstATText(xml)
			assert(tText === 'hello',
				'expected first <a:t> to be "hello"; got: ' + JSON.stringify(tText))
		}
	},
	{
		name: 'B9: mid-text bullet glyph "a • b" is preserved when bullet:true',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('a \u2022 b', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const xml = await getSlide1(zip)
			const tText = firstATText(xml)
			assert(tText === 'a \u2022 b',
				'expected mid-text bullet preserved; got: ' + JSON.stringify(tText))
		}
	},
	{
		name: 'B9: bullet:false preserves leading bullet glyph (user opted out of bullet markup)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('\u2022 item', { x: 1, y: 1, w: 4, h: 1, bullet: false })
			})
			const xml = await getSlide1(zip)
			const tText = firstATText(xml)
			assert(tText === '\u2022 item',
				'expected bullet:false preserves leading glyph; got: ' + JSON.stringify(tText))
		}
	},
	{
		name: 'B9: bullet:{type:"bullet"} also strips leading bullet glyph',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('\u2022 hello', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet' } })
			})
			const xml = await getSlide1(zip)
			const ppr = firstPPr(xml)
			const tText = firstATText(xml)
			assert(/<a:buChar char="&#x2022;"\/>/.test(ppr),
				'expected default <a:buChar/> on <a:pPr>; got: ' + ppr)
			assert(tText === 'hello',
				'expected first <a:t> to be "hello"; got: ' + JSON.stringify(tText))
		}
	},
	{
		name: 'B9: variant glyphs (◦, ▪) are also stripped when bullet:true',
		fn: async () => {
			const { zip: zipHollow } = await build(p => {
				const s = p.addSlide()
				s.addText('\u25E6 a', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const { zip: zipSquare } = await build(p => {
				const s = p.addSlide()
				s.addText('\u25AA b', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const tHollow = firstATText(await getSlide1(zipHollow))
			const tSquare = firstATText(await getSlide1(zipSquare))
			assert(tHollow === 'a', 'expected hollow-circle glyph stripped; got: ' + JSON.stringify(tHollow))
			assert(tSquare === 'b', 'expected black-small-square glyph stripped; got: ' + JSON.stringify(tSquare))
		}
	}
]
