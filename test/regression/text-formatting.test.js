import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

defineRegressionSuite('Text formatting', 'legacy bug-01', [
	{
		name: 'mixed-formatting paragraph emits at most one <a:pPr>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText([{ text: 'Bold ', options: { bold: true } }, { text: 'and regular' }], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
					fontSize: 24,
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paragraphs = xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			assert(paragraphs.length > 0, 'expected at least one <a:p>')
			for (const p of paragraphs) {
				const pPrCount = (p.match(/<a:pPr[\s>]/g) || []).length
				assert(pPrCount <= 1, 'paragraph has ' + pPrCount + ' <a:pPr> tags but OOXML allows at most 1: ' + p)
			}
		},
	},
	{
		// breakLine: false on the last piece of a CRLF-containing run must not force a
		// paragraph break between that piece and whatever follows it (upstream-issue-1138).
		// Before the fix, the CRLF split mutated the shared options object to breakLine:true,
		// so the run after the split always landed in its own paragraph regardless of user intent.
		name: 'breakLine: false on CRLF run keeps following run in same paragraph',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'first\nsecond', options: { breakLine: false } },
						{ text: ' tail', options: {} },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paragraphs = xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			// 'first' → para 1 (break from \n); 'second tail' → para 2 (no trailing break)
			assert(paragraphs.length === 2, 'expected 2 paragraphs, got ' + paragraphs.length)
			assert(paragraphs[1].includes('second'), 'expected "second" in second paragraph')
			assert(paragraphs[1].includes('tail'), 'expected "tail" in same paragraph as "second"')
		},
	},
	{
		name: 'paragraph with mixed-align runs still emits 1 <a:pPr> per <a:p>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'left', options: { align: 'left' } },
						{ text: 'right', options: { align: 'right' } },
					],
					{ x: 1, y: 1, w: 6, h: 1, fontSize: 18 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paragraphs = xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			for (const p of paragraphs) {
				const pPrCount = (p.match(/<a:pPr[\s>]/g) || []).length
				assert(pPrCount <= 1, 'paragraph has ' + pPrCount + ' <a:pPr> tags: ' + p)
			}
		},
	},
])
