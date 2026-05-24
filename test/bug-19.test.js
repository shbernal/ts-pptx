'use strict'

const { build, readEntry, assert } = require('./helpers')

// Extract the first slide's <a:pPr ...>...</a:pPr> block (paragraph properties)
async function getPPr(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const m = xml.match(/<a:pPr[^>]*\/?>(?:[\s\S]*?<\/a:pPr>)?/)
	if (!m) throw new Error('no <a:pPr> found in slide1.xml; xml=' + xml)
	return { xml, ppr: m[0] }
}

module.exports = [
	{
		name: 'bullet:{type:"bullet"} emits default <a:buChar/> with marL/indent (was silently swallowed)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet' } })
			})
			const { ppr, xml } = await getPPr(zip)
			assert(/marL="\d+"/.test(ppr), 'expected marL=".." attr on <a:pPr>; got: ' + ppr + '\nxml: ' + xml)
			assert(/indent="-\d+"/.test(ppr), 'expected indent="-.." attr on <a:pPr>; got: ' + ppr)
			assert(/<a:buChar char="&#x2022;"\/>/.test(ppr),
				'expected default bullet <a:buChar char="&#x2022;"/> emitted for bullet:{type:"bullet"}; got: ' + ppr)
			assert(!/<a:buAutoNum/.test(ppr), 'must NOT emit <a:buAutoNum/> for bullet:{type:"bullet"}; got: ' + ppr)
			assert(!/<a:buNone\/>/.test(ppr), 'must NOT emit <a:buNone/> for bullet:{type:"bullet"}; got: ' + ppr)
		}
	},
	{
		name: 'bullet:{type:"bullet"} produces same bullet markup as bullet:true',
		fn: async () => {
			const { zip: zipObj } = await build(p => {
				const s = p.addSlide()
				s.addText('x', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet' } })
			})
			const { zip: zipTrue } = await build(p => {
				const s = p.addSlide()
				s.addText('x', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const objPpr = (await getPPr(zipObj)).ppr
			const truePpr = (await getPPr(zipTrue)).ppr
			assert(objPpr === truePpr,
				'bullet:{type:"bullet"} <a:pPr> should equal bullet:true <a:pPr>:\n  obj : ' + objPpr + '\n  true: ' + truePpr)
		}
	},
	{
		name: 'regression: bullet:{type:"number"} still emits <a:buAutoNum/>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('one', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'number' } })
			})
			const { ppr } = await getPPr(zip)
			assert(/<a:buAutoNum type="arabicPeriod" startAt="1"\/>/.test(ppr),
				'expected <a:buAutoNum type="arabicPeriod" startAt="1"/> for bullet:{type:"number"}; got: ' + ppr)
			assert(!/<a:buChar/.test(ppr), 'must NOT emit <a:buChar/> for bullet:{type:"number"}; got: ' + ppr)
		}
	},
	{
		name: 'bullet:{type:"bullet", characterCode:"2713"} now reachable → emits <a:buChar char="&#x2713;"/>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('check', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet', characterCode: '2713' } })
			})
			const { ppr } = await getPPr(zip)
			assert(/<a:buChar char="&#x2713;"\/>/.test(ppr),
				'expected <a:buChar char="&#x2713;"/> for bullet:{type:"bullet", characterCode:"2713"}; got: ' + ppr)
			assert(!/<a:buAutoNum/.test(ppr),
				'must NOT emit <a:buAutoNum/> for bullet:{type:"bullet", characterCode:"2713"}; got: ' + ppr)
		}
	}
]
