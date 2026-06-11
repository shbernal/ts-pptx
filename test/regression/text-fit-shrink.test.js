import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// upstream-issue-1199: `fit: 'shrink'` historically emitted a bare <a:normAutofit/>,
// so PowerPoint only shrank text after an edit/resize. The object form bakes explicit
// fontScale/lnSpcReduction (authored as percent, stored in 1000ths of a percent) into
// the file so the text renders pre-shrunk.
defineRegressionSuite('Text fit shrink (normAutofit fontScale/lnSpcReduction)', 'upstream-issue-1199', [
	{
		name: "object 'shrink' emits fontScale + lnSpcReduction in 1000ths of a percent",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('shrink me', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fit: { type: 'shrink', fontScale: 85, lnSpcReduction: 20 },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.indexOf('<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>') !== -1,
				'expected <a:normAutofit fontScale="85000" lnSpcReduction="20000"/>; got: ' + xml
			)
		},
	},
	{
		name: 'only the supplied attribute is emitted',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('scale only', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fit: { type: 'shrink', fontScale: 62.5 },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.indexOf('<a:normAutofit fontScale="62500"/>') !== -1,
				'expected <a:normAutofit fontScale="62500"/> with no lnSpcReduction; got: ' + xml
			)
		},
	},
	{
		name: "bare 'shrink' string still emits attribute-less normAutofit",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('bare', { x: 1, y: 1, w: 4, h: 1, fit: 'shrink' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('<a:normAutofit/>') !== -1, 'expected bare <a:normAutofit/>; got: ' + xml)
		},
	},
	{
		name: 'out-of-range values are dropped (no degenerate attribute)',
		fn: async () => {
			const warnings = []
			const orig = console.warn
			console.warn = (msg) => warnings.push(msg)
			let xml
			try {
				const { zip } = await build((p) => {
					p.addSlide().addText('bad', {
						x: 1,
						y: 1,
						w: 4,
						h: 1,
						fit: { type: 'shrink', fontScale: 150, lnSpcReduction: NaN },
					})
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = orig
			}
			assert(
				xml.indexOf('<a:normAutofit/>') !== -1,
				'expected bare <a:normAutofit/> when values are invalid; got: ' + xml
			)
			assert(warnings.length === 2, 'expected a warning per invalid attribute; got: ' + JSON.stringify(warnings))
		},
	},
])
