import { ChartType } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	assert,
	assertIncludes,
	assertNotIncludes,
	firstXmlBlock,
} from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// a chart's `dataLabelFormatCode` / `valLabelFormatCode` renders in
// LibreOffice but is ignored by PowerPoint and Google Slides (e.g. `0.1` shows instead of `10%`).
// Root cause: PowerPoint and Google Slides display values using the cached *source* number format in
// each series' `<c:val><c:numRef><c:numCache><c:formatCode>` — which was hard-coded to "General" — not
// the `<c:dLbls><c:numFmt>` mask. The fix stamps the resolved value format onto every value numCache.

// Pull the <c:formatCode> from the first value series cache (c:val for bar/line/pie, c:yVal for scatter).
function valCacheFormatCode(xml, valTag) {
	const valBlock = firstXmlBlock(xml, valTag)
	const cacheMatch = valBlock.match(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/)
	assert(cacheMatch, `expected a <c:formatCode> inside <${valTag}>; got: ${valBlock}`)
	return cacheMatch[1]
}

defineRegressionSuite('Chart value format code', [
	{
		// A source-linked axis paints its labels in its cache's format, so an X axis format written
		// linked lost to the value format the X cache carries (probe-scatter-x-cache-format.mjs).
		name: 'scatter chart: a stated X axis format is written unlinked, over the value format the X cache keeps',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'X', values: [0.25, 0.5, 0.75] },
						{ name: 'Y', values: [1, 2, 3] },
					],
					{ type: ChartType.scatter, x: 1, y: 1, w: 6, h: 3, valLabelFormatCode: '0%', catAxisLabelFormatCode: '0.00' }
				)
			})
			const xml = await chartXml(zip)
			// A scatter's first c:valAx is its X axis.
			assertIncludes(firstXmlBlock(xml, 'c:valAx'), '<c:numFmt formatCode="0.00" sourceLinked="0"/>', 'X axis numFmt')
			assert(valCacheFormatCode(xml, 'c:xVal') === '0%', 'the X cache keeps the value format')
		},
	},
	{
		name: 'scatter chart with no X axis format keeps the axis source-linked',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'X', values: [1, 2, 3] },
						{ name: 'Y', values: [1, 2, 3] },
					],
					{ type: ChartType.scatter, x: 1, y: 1, w: 6, h: 3 }
				)
			})
			const xml = await chartXml(zip)
			assertIncludes(
				firstXmlBlock(xml, 'c:valAx'),
				'<c:numFmt formatCode="General" sourceLinked="1"/>',
				'X axis numFmt'
			)
		},
	},
	{
		name: 'bar chart: dataLabelFormatCode flows into the c:val numCache formatCode',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B', 'C'], values: [0.1, 0.2, 0.3] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showValue: true,
					dataLabelFormatCode: '0%',
				})
			})
			const xml = await chartXml(zip)
			assert(
				valCacheFormatCode(xml, 'c:val') === '0%',
				'expected c:val numCache formatCode "0%"; got: ' + valCacheFormatCode(xml, 'c:val')
			)
			// The chart-level data-label mask is still emitted (LibreOffice path) and must agree.
			assertIncludes(xml, '<c:numFmt formatCode="0%" sourceLinked="0"/>', 'dLbls numFmt')
		},
	},
	{
		name: 'bar chart: explicit valLabelFormatCode wins over dataLabelFormatCode for the value cache',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showValue: true,
					dataLabelFormatCode: '0%',
					valLabelFormatCode: '$0.00',
				})
			})
			const xml = await chartXml(zip)
			assert(
				valCacheFormatCode(xml, 'c:val') === '$0.00',
				'expected value cache "$0.00"; got: ' + valCacheFormatCode(xml, 'c:val')
			)
		},
	},
	{
		name: 'bar chart with no explicit format: value cache mirrors the default data-label format (#,##0)',
		fn: async () => {
			// `dataLabelFormatCode` defaults to '#,##0' (gen-objects). Mirroring it into the value cache
			// makes PowerPoint/Google Slides agree with LibreOffice, which already honored that default.
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 3,
				})
			})
			const xml = await chartXml(zip)
			assert(
				valCacheFormatCode(xml, 'c:val') === '#,##0',
				'expected default "#,##0"; got: ' + valCacheFormatCode(xml, 'c:val')
			)
			assertIncludes(xml, '<c:numFmt formatCode="#,##0" sourceLinked="0"/>', 'dLbls numFmt matches cache')
		},
	},
	{
		name: 'pie chart: dataLabelFormatCode is stamped onto the (previously format-less) c:val numCache',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B', 'C'], values: [0.5, 0.3, 0.2] }], {
					type: ChartType.pie,
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showPercent: true,
					dataLabelFormatCode: '0%',
				})
			})
			const xml = await chartXml(zip)
			assert(
				valCacheFormatCode(xml, 'c:val') === '0%',
				'expected pie c:val cache "0%"; got: ' + valCacheFormatCode(xml, 'c:val')
			)
		},
	},
	{
		name: 'scatter chart: dataLabelFormatCode flows into the X/Y value caches (no longer hard-coded General)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'X-Axis', values: [1, 2, 3] },
						{ name: 'Y1', values: [0.1, 0.2, 0.3] },
					],
					{ type: ChartType.scatter, x: 1, y: 1, w: 6, h: 3, showValue: true, dataLabelFormatCode: '0%' }
				)
			})
			const xml = await chartXml(zip)
			assert(
				valCacheFormatCode(xml, 'c:yVal') === '0%',
				'expected scatter c:yVal cache "0%"; got: ' + valCacheFormatCode(xml, 'c:yVal')
			)
			assert(
				valCacheFormatCode(xml, 'c:xVal') === '0%',
				'expected scatter c:xVal cache "0%"; got: ' + valCacheFormatCode(xml, 'c:xVal')
			)
			assertNotIncludes(xml, '<c:formatCode>General</c:formatCode>', 'no leftover hard-coded General value cache')
		},
	},
])
