import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// Regression: parseTextToLines grouped each styled run as a separate inputLines2
// entry, so the word-wrap column counter reset between runs.  Two runs that together
// fit on one line were split into two independent lines, doubling the estimated cell
// height and causing a spurious autoPage slide break.
//
// Reproduces upstream-pr-1237.

defineRegressionSuite('Table autoPage rich-text line wrapping', 'upstream-pr-1237', [
	{
		name: 'two styled runs that fit on one line must not create a spurious slide break',
		fn: async () => {
			// Layout (1 column, 1 inch wide):
			//   CPL ≈ Math.floor(72) / (12 / 2.3) ≈ 13 chars per line
			//   Run 1: "aaaa" (4 chars, bold)
			//   Run 2: "bb"   (2 chars, normal)
			//   Combined: "aaaa" + "bb" = 6 chars < CPL → fits on one line
			//
			// Height parameters (slideMargin:0, margin:0):
			//   lineHeight ≈ inch2Emu(12 * 1.67 / 100) ≈ 0.2004 in
			//   emuSlideTabH = h(0.5in) - y(0.25in) = 0.25 in
			//   0.25 > 0.2004 → 1 line fits on slide 1
			//   0.25 < 0.4008 → 2 lines would overflow to slide 2
			//
			// Before the fix: each run got a separate parsedLines entry → 2 lines → 2 slides.
			// After the fix:  both runs share line-tracking state   → 1 line → 1 slide.

			const rows = [[{ text: [{ text: 'aaaa', options: { bold: true } }, { text: 'bb' }] }]]

			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(rows, {
					x: 0.25,
					y: 0.25,
					w: 1,
					h: 0.5,
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			const slideFiles = listEntries(zip).filter((f) => /ppt\/slides\/slide\d+\.xml$/.test(f))
			assert(
				slideFiles.length === 1,
				`expected 1 slide (both styled runs fit on one line); got ${slideFiles.length} — ` +
					`the runs may have been line-wrapped independently (upstream-pr-1237 bug)`
			)

			const slide1Xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(slide1Xml.includes('>aaaa<'), 'expected "aaaa" text on slide 1')
			assert(slide1Xml.includes('>bb<'), 'expected "bb" text on slide 1')
		},
	},
	{
		name: 'newline in a rich-text run still creates a paragraph break',
		fn: async () => {
			// A cell with a \n-containing run must still produce two separate paragraphs.
			// This guards that the step-2 fix does not collapse \n-separated lines.

			const rows = [[{ text: [{ text: 'line1\nline2' }] }]]

			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(rows, {
					x: 0.25,
					y: 0.25,
					w: 3,
					h: 2,
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			// Both paragraphs must appear on the same single slide (not split across slides).
			const slideFiles = listEntries(zip).filter((f) => /ppt\/slides\/slide\d+\.xml$/.test(f))
			assert(slideFiles.length === 1, `expected 1 slide for newline test; got ${slideFiles.length}`)

			const slide1Xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(slide1Xml.includes('>line1<'), 'expected "line1" in slide XML')
			assert(slide1Xml.includes('>line2<'), 'expected "line2" in slide XML')
		},
	},
	{
		name: 'explicit breakLine:true still separates lines across slides when overflow',
		fn: async () => {
			// A cell with breakLine:true between two long runs must honour the break and
			// can still overflow when the total height exceeds the slide area.
			// Parameters chosen so that run1 fits on slide 1 but run2 overflows to slide 2.

			const rows = [
				[
					{
						text: [{ text: 'first', options: { breakLine: true } }, { text: 'second' }],
					},
				],
			]

			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(rows, {
					x: 0.25,
					y: 0.25,
					w: 3,
					h: 0.5,
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			// breakLine:true forces 2 logical lines → 2 slides given the tight height.
			const slideFiles = listEntries(zip).filter((f) => /ppt\/slides\/slide\d+\.xml$/.test(f))
			assert(
				slideFiles.length === 2,
				`expected 2 slides when breakLine:true forces two logical lines; got ${slideFiles.length}`
			)

			const slide1Xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const slide2Xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			assert(slide1Xml.includes('>first<'), 'expected "first" on slide 1')
			assert(slide2Xml.includes('>second<'), 'expected "second" on slide 2')
		},
	},
])
