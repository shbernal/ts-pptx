import { defineRegressionSuite, build, assertEqual } from '../helpers.js'

// Regression: addSection() must not create a second section with a title that
// already exists. Duplicate section titles confuse section-by-title lookups
// (e.g. addSlide({ sectionTitle }) and autoPage continuation), which silently
// resolve to the first match.
//
// Reproduces upstream-issue-1152.

defineRegressionSuite('addSection duplicate-title guard', 'upstream-issue-1152', [
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
])
