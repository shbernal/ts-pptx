import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertIncludes } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// bubble charts could not show each bubble's size as a data label.
// The bubble `sizes` data already flowed into <c:bubbleSize>, but the data-label block hard-coded
// <c:showBubbleSize val="0"/>. A new `showBubbleSize` chart option now toggles that flag.

const BUBBLE_DATA = [
	{ name: 'X-Axis', values: [1, 2, 3, 4] },
	{ name: 'Y-Values 1', values: [13, 20, 21, 25], sizes: [10, 5, 20, 15] },
]

defineRegressionSuite('Chart bubble size data label', [
	{
		name: 'bubble chart: showBubbleSize true emits <c:showBubbleSize val="1"/>',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(BUBBLE_DATA, {
					type: ChartType.bubble,
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
				p.addSlide().addChart(BUBBLE_DATA, {
					type: ChartType.bubble3d,
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
				p.addSlide().addChart(BUBBLE_DATA, { type: ChartType.bubble, x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:showBubbleSize val="0"/>', 'bubble size label off by default')
			assert(!xml.includes('<c:showBubbleSize val="1"/>'), 'expected no enabled bubble size flag by default')
		},
	},
	{
		// The bubble data-label font size must preserve fractional points (sz is hundredths of a
		// point), matching every other chart data-label site. It previously rounded the size to a
		// whole point first (10.5pt -> sz="1100"); now it converts directly (10.5pt -> sz="1050").
		name: 'bubble chart: fractional dataLabelFontSize keeps half-point precision',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(BUBBLE_DATA, {
					type: ChartType.bubble,
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showLabel: true,
					dataLabelFontSize: 10.5,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, 'sz="1050"', '10.5pt data-label font emits sz="1050"')
			assert(!xml.includes('sz="1100"'), 'expected no whole-point rounding of the data-label size')
		},
	},
])
