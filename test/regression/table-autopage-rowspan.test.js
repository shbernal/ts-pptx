import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// Regression: when a table is auto-paged and a page break would fall inside a
// rowspan group, the rows under the span must NOT be placed on a new slide —
// doing so produces a table with the wrong column count (1 instead of 2) on
// the continuation slide.
//
// Reproduces upstream-pr-1391 / upstream-issue-1231.

defineRegressionSuite('Table autoPage rowspan', 'upstream-pr-1391', [
	{
		name: 'rows under an active rowspan are kept on the same slide as the span anchor',
		fn: async () => {
			// Layout: 2 columns.
			//   Row 0: [A (rowspan=2), B0]
			//   Row 1: [B1]              ← column 0 occupied by A
			//   Row 2: [C0, C1]
			//   Row 3: [D0, D1]
			//   Row 4: [E0, E1]
			//
			// Height parameters are chosen so that:
			//   - Row 0 fits on slide 1 (height 0.3 in, line ≈ 0.2 in)
			//   - Row 1 would trigger the break condition BUT it is under A's span
			//     → the break must be suppressed and row 1 stays on slide 1
			//   - Row 2 then triggers a real break (span is done) → slide 2
			// Expected: 2 slides total (not 3).

			const rows = [
				[{ text: 'A', options: { rowspan: 2 } }, { text: 'B0' }],
				[{ text: 'B1' }],
				[{ text: 'C0' }, { text: 'C1' }],
				[{ text: 'D0' }, { text: 'D1' }],
				[{ text: 'E0' }, { text: 'E1' }],
			]

			const { zip } = await build((p) => {
				const s = p.addSlide()
				// margin:0 / slideMargin:0 eliminate per-row margin overhead so the math is
				// deterministic.  With fontSize:12 each row line is ~0.2004 in tall.
				// emuSlideTabH = h(0.7) - y(0.4) - bottom_margin(0) ≈ 0.3 in.
				// Row 0 (0.20 in) fits.  Row 1 (under rowspan) would push to 0.40 in > 0.30 —
				// without the fix the break fires here; with the fix it is suppressed.
				// Row 2 (0.60 in > 0.30) triggers the real break on the now-span-free row.
				s.addTable(rows, {
					x: 0.5,
					y: 0.4,
					w: 9,
					h: 0.7,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			// 1. Should produce exactly 2 slides (not 3).
			const slideFiles = listEntries(zip).filter((f) => /ppt\/slides\/slide\d+\.xml$/.test(f))
			assert(
				slideFiles.length === 2,
				`expected 2 slides after rowspan-aware page break; got ${slideFiles.length} — ` +
					`the rows under the rowspan may have been split onto a new slide (upstream-pr-1391 bug)`
			)

			// 2. Slide 1 must contain "B1" (row 1 belongs on slide 1 alongside its span anchor).
			const slide1Xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				slide1Xml.includes('>B1<'),
				'expected "B1" text on slide 1 (row under rowspan should stay with span anchor); ' +
					'got slide1.xml: ' +
					slide1Xml.slice(0, 400)
			)

			// 3. Slide 2 must be a 2-column table (row 2 has 2 normal cells).
			const slide2Xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			const gridColCount = (slide2Xml.match(/<a:gridCol\b/g) || []).length
			assert(
				gridColCount === 2,
				`expected 2 <a:gridCol> elements on slide 2; got ${gridColCount} — ` +
					`the slide 2 table has the wrong column count`
			)

			// 4. Slide 2 must contain "C0" (first row after the span boundary).
			assert(slide2Xml.includes('>C0<'), 'expected "C0" text on slide 2 as the first post-span row')
		},
	},
])
