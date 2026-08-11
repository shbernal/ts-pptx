import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../../helpers.js'

// Regression: parseTextToLines grouped each styled run as a separate inputLines2
// entry, so the word-wrap column counter reset between runs.  Two runs that together
// fit on one line were split into two independent lines, doubling the estimated cell
// height and causing a spurious autoPage slide break.
//
// Reproduces upstream-pr-1237.

defineRegressionSuite('Table autoPage rich-text line wrapping [upstream-pr-1237]', [
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
			// `h` is the table's height (an extent), so `y` does not shrink it:
			//   emuSlideTabH = h(0.3 in); one line ≈ 0.2004 in fits, two lines ≈ 0.4008 in overflow.

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
					h: 0.3,
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
	// --- cell-text shapes the wrapper has to survive -------------------------------
	// `parseTextToLines` documents four input shapes for `cell.text` in its own header
	// (string, number, single object, object[]). `addTable` normalizes some of them and
	// passes others straight through, so the wrapper still meets shapes no assertion had
	// pinned. These cases pin what each one *renders*, not merely that it survives.
	{
		name: 'a run whose text is a number is dropped by the wrapper, not rendered',
		fn: async () => {
			// The wrapper groups runs into lines by scanning for "\n"/`breakLine`, and skips
			// any run whose text is not a string. `addTable` stringifies a whole-cell numeric
			// `text` but leaves a numeric run inside an array alone, so the run reaches the
			// wrapper as a number and never makes it into a line — the paged cell keeps only
			// its string runs. Pinned because "silently drops a run" is worth being loud about
			// if it ever changes.
			const rows = [[{ text: [{ text: 'keep' }, { text: 2024 }, { text: 'also' }] }]]

			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
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

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('>keep<'), 'expected the string runs to render')
			assert(xml.includes('>also<'), 'expected the string run after the numeric one to render')
			assert(!xml.includes('2024'), 'a numeric run does not survive the auto-page wrapper')
		},
	},
	{
		name: 'a single-object `text` (not an array) pages as an empty cell',
		fn: async () => {
			// `{ text: { text: 'x' } }` is the "object" shape in the wrapper's own header
			// comment, and `addTable` passes it through untouched. The wrapper only unwraps
			// the string/number and array shapes, so a lone object matches none of them and
			// the cell pages out empty rather than throwing.
			const rows = [[{ text: { text: 'lonely' } }, { text: 'sibling' }]]

			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.25,
					y: 0.25,
					w: 6,
					h: 2,
					colW: [3, 3],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('>sibling<'), 'the neighbouring cell must still render')
			assert(!xml.includes('lonely'), 'a lone object `text` carries no string the wrapper can read')
		},
	},
	{
		name: 'a trailing breakLine run leaves no dangling line buffer',
		fn: async () => {
			// `breakLine` on the *last* run flushes the line buffer inside the loop, so the
			// post-loop flush has nothing left to add. Getting that wrong appends an empty
			// trailing line, which costs a line of measured height and can push a one-line
			// cell onto a second page. `h` here fits exactly one line.
			const rows = [[{ text: [{ text: 'only', options: { breakLine: true } }] }]]

			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.25,
					y: 0.25,
					w: 3,
					h: 0.3,
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			const slides = listEntries(zip).filter((f) => /ppt\/slides\/slide\d+\.xml$/.test(f))
			assert(slides.length === 1, `a trailing breakLine must not add a second line/page; got ${slides.length}`)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('>only<'), 'expected the run to render')
		},
	},
])
