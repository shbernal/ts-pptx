import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Regression: ThemeProps must let callers set the theme's East Asian (<a:ea>) and complex-script
// (<a:cs>) font slots for both the major and minor fonts. PowerPoint emits these empty by default
// and resolves per-script via the <a:font script="..."> list; setting them lets CJK / complex-script
// runs fall back to a caller-chosen theme font (upstream gitbrent/PptxGenJS#1288).

defineRegressionSuite('Theme East Asian / complex-script fonts (upstream #1288)', [
	{
		name: 'EA and CS faces populate <a:ea>/<a:cs> for major and minor fonts',
		fn: async () => {
			const { zip } = await build((p) => {
				p.theme = {
					headFontFace: 'Arial Narrow',
					bodyFontFace: 'Arial',
					headFontFaceEA: 'Yu Gothic',
					bodyFontFaceEA: 'Yu Mincho',
					headFontFaceCS: 'Arial',
					bodyFontFaceCS: 'Times New Roman',
				}
				p.addSlide().addText('テーマ', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/theme/theme1.xml')
			// Major font (headings).
			assert(
				xml.includes(
					'<a:majorFont><a:latin typeface="Arial Narrow"/><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/>'
				),
				'major font ea/cs not set from headFontFaceEA/headFontFaceCS'
			)
			// Minor font (body).
			assert(
				xml.includes(
					'<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Yu Mincho"/><a:cs typeface="Times New Roman"/>'
				),
				'minor font ea/cs not set from bodyFontFaceEA/bodyFontFaceCS'
			)
		},
	},
	{
		name: 'unset EA/CS faces keep PowerPoint empty defaults',
		fn: async () => {
			const { zip } = await build((p) => {
				p.theme = { headFontFace: 'Arial Narrow', bodyFontFace: 'Arial' }
				p.addSlide()
			})
			const xml = await readEntry(zip, 'ppt/theme/theme1.xml')
			assert(
				xml.includes('<a:majorFont><a:latin typeface="Arial Narrow"/><a:ea typeface=""/><a:cs typeface=""/>'),
				'major font ea/cs should stay empty when unset'
			)
			assert(
				xml.includes('<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/>'),
				'minor font ea/cs should stay empty when unset'
			)
		},
	},
])
