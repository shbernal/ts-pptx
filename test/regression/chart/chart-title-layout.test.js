import { ChartType } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertIncludes,
	assertNotIncludes,
} from '../../helpers.js'

// Chart title layout options (`genXmlTitle`): alignment, rotation, and manual
// positioning (`titlePos`). `titlePos` emits a `<c:manualLayout>` whose x and y
// axes are independent — supplying only one of x/y leaves the other on automatic
// layout (only the provided axis gets an `edge` mode entry).

function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

function titleBlock(xml) {
	const match = xml.match(/<c:title>[\s\S]*?<\/c:title>/)
	assert(match, 'expected a <c:title> block; got: ' + xml)
	return match[0]
}

const DATA = [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }]

defineRegressionSuite('Chart title layout', [
	{
		name: 'titleAlign and titleRotate flow into the title paragraph/body properties',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					showTitle: true,
					title: 'Quarterly',
					titleAlign: 'left',
					titleRotate: 45,
				})
			})
			const title = titleBlock(await chartXml(zip))
			assertIncludes(title, '<a:pPr algn="l">', "titleAlign 'left' sets paragraph alignment l")
			assert(/<a:bodyPr rot="\d+"\/>/.test(title), 'titleRotate emits a bodyPr rotation; got: ' + title)
		},
	},
	{
		name: 'titlePos with x and y emits a manualLayout with both edge modes',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					showTitle: true,
					title: 'T',
					titlePos: { x: 0.5, y: 0.5 },
				})
			})
			const title = titleBlock(await chartXml(zip))
			assertIncludes(title, '<c:manualLayout>', 'titlePos emits a manualLayout')
			assertIncludes(title, '<c:xMode val="edge"/>', 'x axis is placed by edge')
			assertIncludes(title, '<c:yMode val="edge"/>', 'y axis is placed by edge')
		},
	},
	{
		name: 'titlePos with only x leaves the y axis on automatic layout',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					showTitle: true,
					title: 'T',
					titlePos: { x: 0.5 },
				})
			})
			const title = titleBlock(await chartXml(zip))
			assertIncludes(title, '<c:manualLayout>', 'a single-axis titlePos still emits manualLayout')
			assertIncludes(title, '<c:xMode val="edge"/>', 'x axis is placed by edge')
			assertNotIncludes(title, '<c:yMode', 'y axis is left on automatic layout (no yMode)')
		},
	},
])
