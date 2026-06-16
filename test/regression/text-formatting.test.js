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
		// Mixed RTL/LTR runs with embedded newlines (upstream-issue-1349). Segments differ only by
		// `lang` (AR vs EN), share `align: 'right'`/`rtlMode`, and some begin or contain "\n". The
		// runs that share a visual line must stay in one paragraph (a lang change is NOT a paragraph
		// break), each newline must start exactly one new paragraph, every <a:p> must carry at most
		// one <a:pPr rtl="1">, and the leading-"\n" split must not leave a junk empty <a:t></a:t> run.
		name: 'mixed RTL/LTR runs with newlines split into clean paragraphs',
		fn: async () => {
			const ar = { align: 'right', fontSize: 12, rtlMode: true, lang: 'AR' }
			const en = { align: 'right', fontSize: 12, rtlMode: true, lang: 'EN' }
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'A\nB ', options: ar },
						{ text: 'text', options: en },
						{ text: ' C ', options: ar },
						{ text: 'eng', options: en },
						{ text: '\nD ', options: ar },
						{ text: 'num', options: en },
						{ text: ' 3', options: ar },
					],
					{ x: 0, y: 1, w: 5, h: 2 }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const body = xml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/)[0]
			const paragraphs = body.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			// 3 newline-delimited lines => exactly 3 paragraphs
			assert(paragraphs.length === 3, 'expected 3 paragraphs, got ' + paragraphs.length)
			// No empty-string run artifact survives the leading-"\n" split
			assert(!body.includes('<a:t></a:t>'), 'expected no empty <a:t></a:t> runs: ' + body)
			const text = (p) => (p.match(/<a:t>[^<]*<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('')
			assert(text(paragraphs[0]) === 'A', 'p1 text should be "A", got "' + text(paragraphs[0]) + '"')
			// Lang change alone must NOT break the line: all four runs stay together
			assert(
				text(paragraphs[1]) === 'B text C eng',
				'p2 should keep mixed-lang runs together, got "' + text(paragraphs[1]) + '"'
			)
			assert(text(paragraphs[2]) === 'D num 3', 'p3 text should be "D num 3", got "' + text(paragraphs[2]) + '"')
			for (const p of paragraphs) {
				const pPrCount = (p.match(/<a:pPr[\s>]/g) || []).length
				assert(pPrCount <= 1, 'paragraph has ' + pPrCount + ' <a:pPr> tags but OOXML allows at most 1: ' + p)
				assert(p.includes('<a:pPr rtl="1"'), 'paragraph should carry rtl="1": ' + p)
			}
			// Per-run direction metadata is preserved (AR and EN runs both present)
			assert(body.includes('lang="AR"') && body.includes('lang="EN"'), 'expected both AR and EN run langs preserved')
		},
	},
	{
		// A *lone* empty run is an intentional blank paragraph and must be preserved (eg: "a\n\nb").
		name: 'blank line keeps its empty paragraph',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('line1\n\nline3', { x: 1, y: 1, w: 4, h: 2 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const body = xml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/)[0]
			const paragraphs = body.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			assert(paragraphs.length === 3, 'expected 3 paragraphs (incl. blank middle), got ' + paragraphs.length)
			const text = (p) => (p.match(/<a:t>[^<]*<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('')
			assert(text(paragraphs[0]) === 'line1', 'p1 should be "line1"')
			assert(text(paragraphs[1]) === '', 'p2 should be the blank line')
			assert(text(paragraphs[2]) === 'line3', 'p3 should be "line3"')
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
