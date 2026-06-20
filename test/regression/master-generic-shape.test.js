import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

async function layoutXmlMatching(zip, re) {
	const layouts = listEntries(zip).filter((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p))
	for (const path of layouts) {
		const xml = await readEntry(zip, path)
		if (re.test(xml)) return xml
	}
	return null
}

defineRegressionSuite('Slide master generic shape object', 'issue-776', [
	{
		name: 'master { shape: { type:"ellipse", options } } emits <a:prstGeom prst="ellipse"> in a layout',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_ELLIPSE',
					objects: [{ shape: { type: 'ellipse', options: { x: 1, y: 1, w: 2, h: 2, fill: { color: 'FF0000' } } } }],
				})
			})
			const xml = await layoutXmlMatching(zip, /<a:prstGeom prst="ellipse"/)
			assert(xml, 'expected a layout containing <a:prstGeom prst="ellipse">')
			assert(/<a:srgbClr val="FF0000"/.test(xml), 'expected the shape fill color in the layout; got: ' + xml)
		},
	},
	{
		name: 'generic shape route handles a preset with no dedicated shortcut (chevron)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_CHEVRON',
					objects: [{ shape: { type: 'chevron', options: { x: 0.5, y: 0.5, w: 3, h: 1 } } }],
				})
			})
			const xml = await layoutXmlMatching(zip, /<a:prstGeom prst="chevron"/)
			assert(xml, 'expected a layout containing <a:prstGeom prst="chevron">')
		},
	},
])
