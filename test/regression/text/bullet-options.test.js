import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

// Extract the first slide's <a:pPr ...>...</a:pPr> block (paragraph properties)
async function getPPr(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const m = xml.match(/<a:pPr[^>]*\/?>(?:[\s\S]*?<\/a:pPr>)?/)
	if (!m) throw new Error('no <a:pPr> found in slide1.xml; xml=' + xml)
	return { xml, ppr: m[0] }
}

defineRegressionSuite('Bullet option serialization [legacy bug-19]', [
	{
		name: 'bullet:{type:"bullet"} emits default <a:buChar/> with marL/indent (was silently swallowed)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet' } })
			})
			const { ppr, xml } = await getPPr(zip)
			assert(/marL="\d+"/.test(ppr), 'expected marL=".." attr on <a:pPr>; got: ' + ppr + '\nxml: ' + xml)
			assert(/indent="-\d+"/.test(ppr), 'expected indent="-.." attr on <a:pPr>; got: ' + ppr)
			assert(
				/<a:buChar char="&#x2022;"\/>/.test(ppr),
				'expected default bullet <a:buChar char="&#x2022;"/> emitted for bullet:{type:"bullet"}; got: ' + ppr
			)
			assert(!/<a:buAutoNum/.test(ppr), 'must NOT emit <a:buAutoNum/> for bullet:{type:"bullet"}; got: ' + ppr)
			assert(!/<a:buNone\/>/.test(ppr), 'must NOT emit <a:buNone/> for bullet:{type:"bullet"}; got: ' + ppr)
		},
	},
	{
		name: 'bullet:{type:"bullet"} produces same bullet markup as bullet:true',
		fn: async () => {
			const { zip: zipObj } = await build((p) => {
				const s = p.addSlide()
				s.addText('x', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet' } })
			})
			const { zip: zipTrue } = await build((p) => {
				const s = p.addSlide()
				s.addText('x', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const objPpr = (await getPPr(zipObj)).ppr
			const truePpr = (await getPPr(zipTrue)).ppr
			assert(
				objPpr === truePpr,
				'bullet:{type:"bullet"} <a:pPr> should equal bullet:true <a:pPr>:\n  obj : ' + objPpr + '\n  true: ' + truePpr
			)
		},
	},
	{
		name: 'regression: bullet:{type:"number"} still emits <a:buAutoNum/>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('one', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'number' } })
			})
			const { ppr } = await getPPr(zip)
			assert(
				/<a:buAutoNum type="arabicPeriod" startAt="1"\/>/.test(ppr),
				'expected <a:buAutoNum type="arabicPeriod" startAt="1"/> for bullet:{type:"number"}; got: ' + ppr
			)
			assert(!/<a:buChar/.test(ppr), 'must NOT emit <a:buChar/> for bullet:{type:"number"}; got: ' + ppr)
		},
	},
	{
		name: 'bullet:{type:"bullet", characterCode:"2713"} now reachable → emits <a:buChar char="&#x2713;"/>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('check', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet', characterCode: '2713' } })
			})
			const { ppr } = await getPPr(zip)
			assert(
				/<a:buChar char="&#x2713;"\/>/.test(ppr),
				'expected <a:buChar char="&#x2713;"/> for bullet:{type:"bullet", characterCode:"2713"}; got: ' + ppr
			)
			assert(
				!/<a:buAutoNum/.test(ppr),
				'must NOT emit <a:buAutoNum/> for bullet:{type:"bullet", characterCode:"2713"}; got: ' + ppr
			)
		},
	},
	{
		// The pair is asserted together, because either half alone passes against the bug:
		// before `'inherit'` existed both spellings produced the identical `a:buNone` block.
		name: "bullet:'inherit' states nothing, where an omitted bullet states <a:buNone/>",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'omitted', options: { breakLine: true } },
						{ text: 'inherits', options: { bullet: 'inherit' } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paras = [...xml.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((m) => m[0])
			assert(paras.length === 2, 'expected two paragraphs; got ' + paras.length + '\nxml: ' + xml)
			assert(
				/<a:pPr indent="0" marL="0"><a:buNone\/><\/a:pPr>/.test(paras[0]),
				'an omitted bullet must keep emitting the explicit off; got: ' + paras[0]
			)
			assert(!/<a:pPr/.test(paras[1]), "bullet:'inherit' must emit no <a:pPr> at all here; got: " + paras[1])
			assert(!/<a:buNone/.test(paras[1]), "bullet:'inherit' must not emit <a:buNone/>; got: " + paras[1])
			assert(!/marL=|indent=/.test(paras[1]), "bullet:'inherit' must not flatten marL/indent; got: " + paras[1])
		},
	},
	{
		name: "bullet:'inherit' keeps the paragraph's other properties",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('aligned', { x: 1, y: 1, w: 4, h: 1, bullet: 'inherit', align: 'center', indentLevel: 2 })
			})
			const { ppr } = await getPPr(zip)
			assert(/algn="ctr"/.test(ppr), 'expected algn="ctr" to survive; got: ' + ppr)
			assert(/lvl="2"/.test(ppr), 'expected lvl="2" to survive; got: ' + ppr)
			assert(!/<a:buNone/.test(ppr), 'must NOT emit <a:buNone/>; got: ' + ppr)
			assert(!/marL=|indent=/.test(ppr), 'must NOT emit marL/indent; got: ' + ppr)
		},
	},
	{
		// `'inherit'` is a TRUTHY member of the `bullet` union, and two sites downstream of the
		// emitter used to test it for truthiness. This is the line-grouping one: a bullet starts
		// a new paragraph, and a bullet that draws nothing must not.
		name: "bullet:'inherit' on consecutive runs does not split them into paragraphs",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'run one ', options: { bullet: 'inherit' } },
						{ text: 'run two', options: { bullet: 'inherit' } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paras = [...xml.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
			assert(paras.length === 1, 'expected ONE paragraph; got ' + paras.length + '\nxml: ' + xml)
			assert(/run one <\/a:t>/.test(xml) && /run two<\/a:t>/.test(xml), 'both runs should survive; xml: ' + xml)
		},
	},
	{
		// The other truthiness site: a leading glyph is stripped so it is not drawn twice beside
		// the `a:buChar` the paragraph emits. Under `'inherit'` there is no `a:buChar` to double,
		// so the character is the author's text and must survive.
		name: "a leading bullet glyph survives bullet:'inherit' and is still stripped for bullet:true",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: '• kept', options: { bullet: 'inherit', breakLine: true } },
						{ text: '• stripped', options: { bullet: true } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:t>• kept<\/a:t>/.test(xml), "the glyph must survive bullet:'inherit'; xml: " + xml)
			assert(/<a:t>stripped<\/a:t>/.test(xml), 'the glyph must still be stripped for bullet:true; xml: ' + xml)
		},
	},
	{
		// A paragraph whose FIRST run inherits and whose continuation runs state no bullet — the
		// exact shape the pptx-to-script converter emits. The paragraph-property emitter used to
		// retry on each run until one produced non-empty XML, so the continuation run supplied a
		// `a:buNone` the first run had asked to leave out, appended after its `<a:r>`.
		name: "bullet:'inherit' on the first run of a multi-run paragraph is not undone by the rest",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'first ', options: { bullet: 'inherit' } },
						{ text: 'second', options: { bold: true } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/<a:buNone/.test(xml), 'no <a:buNone/> may reach the slide; xml: ' + xml)
			assert(!/<a:r>[\s\S]*<a:pPr/.test(xml), '<a:pPr> may never follow a run; xml: ' + xml)
		},
	},
])
