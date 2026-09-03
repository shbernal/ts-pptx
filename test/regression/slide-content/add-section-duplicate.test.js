import { defineRegressionSuite, build, assertEqual, captureDiagnostics } from '../../helpers.js'

// Regression: addSection() must not create a second section with a title that
// already exists. Duplicate section titles confuse section-by-title lookups
// (e.g. addSlide({ sectionTitle }) and autoPage continuation), which silently
// resolve to the first match.
//
// Reproduces upstream-issue-1152.

defineRegressionSuite('addSection duplicate-title guard [upstream-issue-1152]', [
	{
		name: 'a duplicate section title is ignored, not appended',
		fn: async () => {
			const { pres } = await build((p) => {
				p.addSection({ title: 'Charts' })
				p.addSection({ title: 'Charts' }) // duplicate — should be ignored
				p.addSection({ title: 'Tables' })
			})

			assertEqual(pres.sections.length, 2, 'duplicate "Charts" section should not be added')
			assertEqual(pres.sections[0].title, 'Charts', 'first section')
			assertEqual(pres.sections[1].title, 'Tables', 'second section')
		},
	},
	{
		name: 'a section without a title is ignored',
		fn: async () => {
			const { pres } = await build((p) => {
				p.addSection({ title: 'Intro' })
				p.addSection({}) // missing title — should be ignored, not pushed titleless
			})

			assertEqual(pres.sections.length, 1, 'titleless section should not be added')
			assertEqual(pres.sections[0].title, 'Intro', 'only valid section remains')
		},
	},
	{
		name: 'order counts from 1, so `order: 1` puts the section first',
		fn: async () => {
			// The option documents `1-n` and the index was used raw, so `order: 1` inserted
			// *second* -- off by one against its own documentation, and silently.
			const { pres } = await build((p) => {
				p.addSection({ title: 'A' })
				p.addSection({ title: 'B' })
				p.addSection({ title: 'C', order: 1 })
			})
			assertEqual(pres.sections.map((s) => s.title).join(','), 'C,A,B', 'the section asked to be first is first')
		},
	},
	{
		name: 'an order past the end appends',
		fn: async () => {
			const { pres } = await build((p) => {
				p.addSection({ title: 'A' })
				p.addSection({ title: 'B', order: 12 })
			})
			assertEqual(pres.sections.map((s) => s.title).join(','), 'A,B', 'nothing to insert before')
		},
	},
	{
		name: 'an order that names no position warns and appends',
		fn: async () => {
			// `0` was the sharp edge: falsy, so it took the append branch with no word said,
			// while every other unusable value spliced from somewhere unpredictable.
			const { codes } = await captureDiagnostics(async () => {
				const { pres } = await build((p) => {
					p.addSection({ title: 'A' })
					p.addSection({ title: 'B', order: 0 })
					p.addSection({ title: 'C', order: -2 })
					p.addSection({ title: 'D', order: 1.5 })
				})
				assertEqual(pres.sections.map((s) => s.title).join(','), 'A,B,C,D', 'each unusable order appends')
			})
			assertEqual(
				codes.filter((c) => c === 'section/invalid-order').length,
				3,
				'and each one says so: got ' + JSON.stringify(codes)
			)
		},
	},
])
