import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

// Every `<a:pPr>` on slide 1, in document order — these cases are mostly about the *pairing*
// of two paragraphs, since either half alone passes against the bug they cover.
async function paragraphProps(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const paras = [...xml.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((m) => m[0])
	return { xml, paras, pPrs: paras.map((p) => (p.match(/<a:pPr[^>]*(?:\/>|>[\s\S]*?<\/a:pPr>)/) ?? [''])[0]) }
}

defineRegressionSuite('Paragraph margins (a:pPr/@marL, @indent)', 'issue-15-followup', [
	{
		// 27pt is DEF_BULLET_MARGIN, the hang every bulleted paragraph used to be pinned to.
		name: 'paraMarginLeft/paraIndent override the bullet-derived default',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hanging', { x: 1, y: 1, w: 4, h: 1, bullet: true, paraMarginLeft: 36, paraIndent: -18 })
			})
			const { pPrs } = await paragraphProps(zip)
			assert(/marL="457200"/.test(pPrs[0]), 'expected marL="457200" (36pt); got: ' + pPrs[0])
			assert(/indent="-228600"/.test(pPrs[0]), 'expected indent="-228600" (-18pt); got: ' + pPrs[0])
			assert(!/marL="342900"/.test(pPrs[0]), 'the bullet default must not survive an override; got: ' + pPrs[0])
			assert(/<a:buChar/.test(pPrs[0]), 'the bullet itself must still be drawn; got: ' + pPrs[0])
		},
	},
	{
		// The pair is asserted together: before the option existed, both spellings produced the
		// identical `marL="0" indent="0"`, so either half alone passes against the bug.
		name: "bullet:false with paraMarginLeft/paraIndent 'inherit' suppresses the bullet without flattening the margins",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'off, margins inherited', options: { bullet: false, paraMarginLeft: 'inherit', paraIndent: 'inherit', breakLine: true } }, // prettier-ignore
						{ text: 'off, margins flattened', options: { bullet: false } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const { pPrs } = await paragraphProps(zip)
			assert(/<a:buNone\/>/.test(pPrs[0]), 'the explicit off must still be stated; got: ' + pPrs[0])
			assert(!/marL=|indent=/.test(pPrs[0]), "'inherit' must state no margin at all; got: " + pPrs[0])
			assert(
				/<a:pPr indent="0" marL="0"><a:buNone\/><\/a:pPr>/.test(pPrs[1]),
				'an omitted margin option must keep writing the historical zeros; got: ' + pPrs[1]
			)
		},
	},
	{
		name: "bullet:'inherit' can still state margins of its own",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('indented', { x: 1, y: 1, w: 4, h: 1, bullet: 'inherit', paraMarginLeft: 0, paraIndent: 18 })
			})
			const { pPrs } = await paragraphProps(zip)
			assert(/marL="0"/.test(pPrs[0]), 'expected an explicit marL="0"; got: ' + pPrs[0])
			assert(/indent="228600"/.test(pPrs[0]), 'expected a positive first-line indent; got: ' + pPrs[0])
			assert(!/<a:buNone/.test(pPrs[0]), "bullet:'inherit' must not gain a bullet child; got: " + pPrs[0])
		},
	},
	{
		// `0` is a real margin and `||` would swap it for the shape's value — the reason the
		// shape-level inheritance of these two uses `??` where the options beside it use `||`.
		name: 'a run-level paraMarginLeft of 0 beats the shape-level value',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'zero', options: { bullet: 'inherit', paraMarginLeft: 0, breakLine: true } },
						{ text: 'shape', options: { bullet: 'inherit' } },
					],
					{ x: 1, y: 1, w: 4, h: 1, paraMarginLeft: 45 }
				)
			})
			const { pPrs } = await paragraphProps(zip)
			assert(/marL="0"/.test(pPrs[0]), 'the run must keep its explicit zero; got: ' + pPrs[0])
			assert(/marL="571500"/.test(pPrs[1]), 'the second paragraph must take the shape value; got: ' + pPrs[1])
		},
	},
	{
		// ST_TextMargin is unsigned and ST_TextIndent tops out at 4032pt; out of range makes
		// PowerPoint report the package as needing repair, so both are clamped rather than passed.
		name: 'out-of-range margins are clamped into their ST_* ranges',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('clamped', { x: 1, y: 1, w: 4, h: 1, bullet: 'inherit', paraMarginLeft: -5, paraIndent: 99999 })
			})
			const { pPrs } = await paragraphProps(zip)
			assert(/marL="0"/.test(pPrs[0]), 'a negative margin must clamp to 0; got: ' + pPrs[0])
			assert(/indent="51206400"/.test(pPrs[0]), 'indent must clamp to 4032pt; got: ' + pPrs[0])
		},
	},
	{
		// The byte-identity half: a deck that never mentions the options must emit exactly what
		// it emitted before they existed, in the same attribute order.
		name: 'an unset option changes nothing about the historical output',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'bulleted', options: { bullet: true, breakLine: true } },
						{ text: 'not bulleted', options: { bullet: false } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const { pPrs } = await paragraphProps(zip)
			assert(
				pPrs[0].startsWith('<a:pPr marL="342900" indent="-342900">'),
				'a drawn bullet keeps marL-then-indent at the 27pt default; got: ' + pPrs[0]
			)
			assert(
				pPrs[1].startsWith('<a:pPr indent="0" marL="0">'),
				'the no-bullet arm keeps its indent-then-marL order; got: ' + pPrs[1]
			)
		},
	},
])
