import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Regression: a custom `fontFace` must be written the way PowerPoint writes a font picked from the
// UI — into the Latin (<a:latin>) and complex-script (<a:cs>) slots only. Forcing a Latin-only face
// into the East Asian (<a:ea>) slot, especially with the bogus charset values PowerPoint never emits
// on ea/cs (charset="-122"/"-120"), duplicates/ghosts text in Office 365 (upstream gitbrent/PptxGenJS#1301).
// `fontFaceEA` opts back into an explicit <a:ea> face for CJK text (preserves upstream #174).

defineRegressionSuite('Text fontFace latin/ea/cs slots (upstream #1301)', [
	{
		name: 'fontFace fills latin + cs only, leaving ea to inherit the theme',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('Hello', { x: 1, y: 1, fontFace: 'Jost Light' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.includes('<a:latin typeface="Jost Light" pitchFamily="34" charset="0"/><a:cs typeface="Jost Light"/>'),
				'fontFace should emit latin + cs with the face and no ea/cs charset'
			)
			// The ghost trigger: a Latin face in <a:ea> and the non-conformant negative charsets.
			assert(!xml.includes('<a:ea typeface="Jost Light"'), 'fontFace must not force a Latin face into the <a:ea> slot')
			assert(
				!xml.includes('charset="-122"') && !xml.includes('charset="-120"'),
				'ea/cs must not emit bogus negative charsets'
			)
		},
	},
	{
		name: 'fontFaceEA emits an explicit East Asian face in document order (latin, ea, cs)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('你好', { x: 1, y: 1, fontFace: 'Jost Light', fontFaceEA: '微軟正黑體' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.includes(
					'<a:latin typeface="Jost Light" pitchFamily="34" charset="0"/><a:ea typeface="微軟正黑體"/><a:cs typeface="Jost Light"/>'
				),
				'fontFaceEA should emit <a:ea> between latin and cs with no charset'
			)
		},
	},
	{
		name: 'table cell endParaRPr follows the same latin/ea/cs rule',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: '', options: { fontFace: 'Jost Light', fontSize: 18 } }]], { x: 1, y: 1, w: 4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.includes('<a:latin typeface="Jost Light" charset="0"/><a:cs typeface="Jost Light"/>'),
				'empty table cell should emit latin + cs only (no forced ea)'
			)
			assert(!xml.includes('<a:ea typeface="Jost Light"'), 'table cell must not force a Latin face into <a:ea>')
		},
	},
])
