import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Regression: when a table with autoPage overflows and the originating slide is
// NOT in the last-defined section, continuation slides must land in the same
// section as the originating slide — not in a new Default-N section.
//
// Reproduces upstream-issue-1405.

defineRegressionSuite('Table autoPage section preservation', 'upstream-issue-1405', [
	{
		name: 'overflow slides stay in the originating slide section, not Default-N',
		fn: async () => {
			// Layout:
			//   Section A  — Slide 1 (has the autoPage table)
			//   Section B  — (empty; defined after A to make it the "last section")
			//
			// With the old heuristic, addNewSlide checked whether the last slide in
			// the deck was in the *last section* (Section B).  s1 is NOT in Section B,
			// so sectAlreadyInUse = false → sectionTitle = null → addSlide sees no
			// sectionTitle with a non-default last section → creates Default-1.
			//
			// With the fix, addNewSlide finds the section that *actually contains*
			// the last slide (s1 → Section A) and uses that title.

			const rows = Array.from({ length: 10 }, (_, i) => [{ text: `Row ${i} col A` }, { text: `Row ${i} col B` }])

			const { zip } = await build((p) => {
				p.addSection({ title: 'Section A' })
				p.addSection({ title: 'Section B' }) // becomes the "last section"
				const s1 = p.addSlide({ sectionTitle: 'Section A' })
				// Tight box (0.3 in) + fontSize:12 (~0.2 in/row) forces overflow after ~1 row.
				s1.addTable(rows, {
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

			const presXml = await readEntry(zip, 'ppt/presentation.xml')

			// 1. No Default-N section should exist.
			assert(
				!presXml.includes('name="Default-'),
				'expected no Default-N section; autoPage overflow should stay in Section A. ' +
					'presentation.xml excerpt: ' +
					presXml.slice(presXml.indexOf('sectionLst') - 5, presXml.indexOf('sectionLst') + 500)
			)

			// 2. Section A must own more than one slide (the source + at least one overflow).
			const sectionAMatch = presXml.match(/<p14:section name="Section A"[^>]*>([\s\S]*?)<\/p14:section>/)
			assert(sectionAMatch, 'expected a <p14:section name="Section A"> in presentation.xml')
			const sectionASlideCount = (sectionAMatch[1].match(/<p14:sldId\b/g) || []).length
			assert(
				sectionASlideCount >= 2,
				`expected Section A to contain ≥2 slides (source + overflow); got ${sectionASlideCount}`
			)

			// 3. Section B must remain empty.
			const sectionBMatch = presXml.match(/<p14:section name="Section B"[^>]*>([\s\S]*?)<\/p14:section>/)
			assert(sectionBMatch, 'expected a <p14:section name="Section B"> in presentation.xml')
			const sectionBSlideCount = (sectionBMatch[1].match(/<p14:sldId\b/g) || []).length
			assert(sectionBSlideCount === 0, `expected Section B to remain empty; got ${sectionBSlideCount} slides`)
		},
	},
])
