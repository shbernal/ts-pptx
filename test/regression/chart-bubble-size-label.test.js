import { defineRegressionSuite, build, readEntry, listEntries, assert, assertIncludes } from '../helpers.js'

// Upstream gitbrent/PptxGenJS#744: bubble charts could not show each bubble's size as a data label.
// The bubble `sizes` data already flowed into <c:bubbleSize>, but the data-label block hard-coded
// <c:showBubbleSize val="0"/>. A new `showBubbleSize` chart option now toggles that flag.

function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

const BUBBLE_DATA = [
	{ name: 'X-Axis', values: [1, 2, 3, 4] },
	{ name: 'Y-Values 1', values: [13, 20, 21, 25], sizes: [10, 5, 20, 15] },
]

defineRegressionSuite('Chart bubble size data label (upstream #744)', [
	{
		name: 'bubble chart: showBubbleSize true emits <c:showBubbleSize val="1"/>',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bubble, BUBBLE_DATA, {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showBubbleSize: true,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:showBubbleSize val="1"/>', 'bubble size label enabled')
		},
	},
	{
		name: 'bubble3D chart: showBubbleSize true emits <c:showBubbleSize val="1"/>',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bubble3d, BUBBLE_DATA, {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showBubbleSize: true,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:showBubbleSize val="1"/>', 'bubble3D size label enabled')
		},
	},
	{
		name: 'bubble chart: default (omitted) keeps <c:showBubbleSize val="0"/>',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bubble, BUBBLE_DATA, { x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:showBubbleSize val="0"/>', 'bubble size label off by default')
			assert(!xml.includes('<c:showBubbleSize val="1"/>'), 'expected no enabled bubble size flag by default')
		},
	},
])
