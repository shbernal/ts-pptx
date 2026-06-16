import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Upstream PR #1458 (issue #1443) proposed stripping every placeholder <p:sp> from
// notesMaster1.xml down to an empty <p:spTree>, on the theory that PowerPoint's repair
// pass removes them. That was a misdiagnosis: the real repair trigger was notesMaster1's
// theme relationship pointing at the slideMaster's theme1.xml (fixed by giving it theme2.xml,
// guarded in theme-relationships.test.js). Our notesMaster emits the verbatim
// PowerPoint-authored placeholder set, and makeXmlNotesSlide placeholders inherit their
// geometry/style from these master placeholders — so stripping them would orphan those
// inheritances and re-trigger repair. This suite locks in that we KEEP them.
defineRegressionSuite('Notes master placeholders (#1443, #1458)', 'upstream-pr-1458', [
	{
		name: 'notesMaster1.xml retains its placeholder shapes (not stripped to an empty spTree)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
				s.addNotes('speaker notes')
			})
			const xml = await readEntry(zip, 'ppt/notesMasters/notesMaster1.xml')

			// The authored notesMaster ships six placeholders; require the full set so a
			// strip-everything change cannot pass unnoticed.
			const phMatches = xml.match(/<p:ph\b[^>]*>/g) || []
			assert(phMatches.length === 6, 'expected 6 <p:ph> placeholders in notesMaster1.xml; got ' + phMatches.length)

			// The notesSlide placeholders (sldImg, body, sldNum) inherit from these; assert each
			// inheritance source is present by type so a partial removal also fails.
			for (const phType of ['hdr', 'dt', 'sldImg', 'body', 'ftr', 'sldNum']) {
				assert(xml.includes('type="' + phType + '"'), 'notesMaster1.xml missing placeholder type="' + phType + '"')
			}
		},
	},
	{
		name: 'notesMaster1.xml retains <p:notesStyle> and <p:clrMap>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/notesMasters/notesMaster1.xml')
			assert(xml.includes('<p:notesStyle>'), 'notesMaster1.xml missing <p:notesStyle>')
			assert(xml.includes('<p:clrMap '), 'notesMaster1.xml missing <p:clrMap>')
		},
	},
])
