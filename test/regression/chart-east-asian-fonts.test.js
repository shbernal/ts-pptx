import { defineRegressionSuite, build, readEntry, listEntries, assert, assertIncludes } from '../helpers.js'

// Upstream gitbrent/PptxGenJS#1420: chart title, legend, and axis/category label font settings did not
// take effect for Chinese (and other East Asian) text, most visibly on PowerPoint for Mac.
// Root cause: chart text runs emitted only `<a:latin typeface="...">`. In DrawingML a typeface applies
// only to the script class of its element — `<a:latin>` covers Latin/ASCII, `<a:ea>` East Asian, `<a:cs>`
// complex scripts — so East Asian glyphs fell back to the theme font. The fix stamps the requested
// typeface onto all three (`<a:latin>/<a:ea>/<a:cs>`), which is what choosing a font in PowerPoint does.

function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

// Assert the latin/ea/cs trio is present for a given typeface.
function assertFontTrio(xml, face, label) {
	assertIncludes(xml, `<a:latin typeface="${face}"/>`, `${label}: latin`)
	assertIncludes(xml, `<a:ea typeface="${face}"/>`, `${label}: ea`)
	assertIncludes(xml, `<a:cs typeface="${face}"/>`, `${label}: cs`)
}

defineRegressionSuite('Chart East Asian fonts (upstream #1420)', [
	{
		name: 'chart title font emits latin/ea/cs trio',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showTitle: true,
					title: '图表标题',
					titleFontFace: 'Microsoft YaHei',
				})
			})
			const xml = await chartXml(zip)
			assertFontTrio(xml, 'Microsoft YaHei', 'title')
		},
	},
	{
		name: 'legend font emits latin/ea/cs trio (no orphaned latin-only run)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showLegend: true,
					legendFontFace: 'SimSun',
				})
			})
			const xml = await chartXml(zip)
			assertFontTrio(xml, 'SimSun', 'legend')
		},
	},
	{
		name: 'category and value axis label fonts emit latin/ea/cs trio',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					catAxisLabelFontFace: 'KaiTi',
					valAxisLabelFontFace: 'FangSong',
				})
			})
			const xml = await chartXml(zip)
			assertFontTrio(xml, 'KaiTi', 'catAxis')
			assertFontTrio(xml, 'FangSong', 'valAxis')
		},
	},
	{
		name: 'data label font emits latin/ea/cs trio',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showValue: true,
					dataLabelFontFace: 'NSimSun',
				})
			})
			const xml = await chartXml(zip)
			assertFontTrio(xml, 'NSimSun', 'dataLabel')
		},
	},
])
