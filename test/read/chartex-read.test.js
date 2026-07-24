// Write→read fidelity for the chartEx (`cx:`) reader added in src/read/api/chartex.ts.
//
// chartEx is the Office-2016 chart family (waterfall/funnel/treemap/sunburst/
// histogram/pareto/boxWhisker/regionMap) — a BLIND SPOT the deep model had no
// branch for: the frame is wrapped in `mc:AlternateContent` (so the shape reader
// never even enumerated it), and its data lives in a `cx:chartData` block series
// point at by `cx:dataId`, not the classic `c:cat`/`c:val` caches. The writer
// already emits every piece (src/gen/chart/chartex-*.ts), so each getter is proven
// by authoring a chart with the write API, reading it back through the deep model,
// and asserting the extracted values. The writer's bytes are the fixture.

import { ChartType } from '../../dist/node.js'
import { describe, test } from 'vitest'
import { authorRead, firstChartEx, firstShape, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

describe('ChartEx — write→read fidelity', () => {
	test('a chartEx frame is surfaced through its mc:AlternateContent wrapper', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'Cash Flow', labels: ['Start', 'Q1', 'End'], values: [100, 40, 190] }], {
				type: ChartType.waterfall,
				x: 1,
				y: 1,
				w: 6,
				h: 4,
			})
		})
		// The frame lives inside mc:Choice; the reader must unwrap it (not skip the AlternateContent).
		const frame = firstShape(presentation, (s) => s.shapeType === 'graphicFrame')
		assert(frame !== null, 'the chartEx graphicFrame is enumerated as a shape')
		assertEqual(frame.hasChartEx, true, 'the frame reports a chartEx host')
		assertEqual(frame.hasChart, false, 'and is NOT a classic chart')
		assert(frame.chart === null, 'the classic chart getter is null for a chartEx frame')
		assert(frame.chartEx !== null, 'the chartEx getter resolves the cx: part')
		assert(frame.chartEx.partName.endsWith('.xml'), `part name resolves, got ${frame.chartEx.partName}`)
	})

	test('a waterfall reads its layout, title, legend, categories, and series values', async () => {
		const { presentation } = await authorRead((pres) => {
			pres
				.addSlide()
				.addChart([{ name: 'Cash Flow', labels: ['Start', 'Q1', 'Q2', 'End'], values: [100, 40, -30, 190] }], {
					type: ChartType.waterfall,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					showTitle: true,
					title: 'Annual Cash Flow',
					showValue: true,
					showLegend: true,
					legendPos: 'b',
					subtotals: [0, 3],
				})
		})
		const cx = firstChartEx(presentation)
		assert(cx !== null, 'the chartEx chart is located')
		assertEqual(cx.layoutId, 'waterfall', 'the raw cx: layout token')
		assert(Array.isArray(cx.layoutIds) && cx.layoutIds.length === 1, 'a single-series waterfall has one layoutId')
		assertEqual(cx.title, 'Annual Cash Flow', 'the rich title text flattens')
		assertEqual(cx.legend?.position, 'b', 'the legend position round-trips')
		assertEqual(cx.categories.join(','), 'Start,Q1,Q2,End', 'category labels read from the cx:strDim leaf level')
		const ser = cx.series[0]
		assert(ser != null, 'the series is read')
		assertEqual(ser.name, 'Cash Flow', 'the series name reads from cx:tx/cx:v')
		assertEqual(ser.dataId, 0, 'the series points at data block 0')
		assertEqual(ser.values.join(','), '100,40,-30,190', 'the numeric values resolve through the dataId')
	})

	test('showValue emits data labels the reader decodes as a visibility toggle set', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
				type: ChartType.waterfall,
				x: 1,
				y: 1,
				w: 5,
				h: 3,
				showValue: true,
			})
		})
		const labels = firstChartEx(presentation).series[0].dataLabels
		assert(labels !== null, 'showValue produced a data-label block')
		assertEqual(labels.value, true, 'the value toggle is on')
		assertEqual(labels.seriesName, false, 'the series-name toggle is off')
		assertEqual(labels.categoryName, false, 'the category-name toggle is off')
	})

	test('waterfall axes read id + kind + gap/gridlines off the scaling child', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
				type: ChartType.waterfall,
				x: 1,
				y: 1,
				w: 5,
				h: 3,
			})
		})
		const axes = firstChartEx(presentation).axes
		assertEqual(axes.length, 2, 'a waterfall carries a category + value axis')
		const cat = axes.find((a) => a.kind === 'cat')
		const val = axes.find((a) => a.kind === 'val')
		assert(cat != null && val != null, 'both axis kinds are distinguished by their scaling child')
		assertEqual(cat.id, 0, 'the category axis is id 0')
		assertEqual(cat.gapWidth, 0.5, 'the fractional gap width reads as a float (0.5, not the classic integer percent)')
		assertEqual(val.id, 1, 'the value axis is id 1')
		assertEqual(val.majorGridlines, true, 'the value axis carries major gridlines')
	})

	test('pareto surfaces both series with the derived line binding its owner and a secondary axis', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'Defects', labels: ['A', 'B', 'C'], values: [50, 30, 10] }], {
				type: ChartType.pareto,
				x: 1,
				y: 1,
				w: 5,
				h: 3,
			})
		})
		const cx = firstChartEx(presentation)
		assertEqual(cx.layoutIds.join(','), 'clusteredColumn,paretoLine', 'pareto emits a column + a cumulative line')
		assertEqual(cx.series.length, 2, 'both series are read')
		assertEqual(cx.series[0].name, 'Defects', 'the column series is named')
		assertEqual(cx.series[1].ownerIndex, 0, 'the paretoLine derives its data from series 0 (ownerIdx)')
		assertEqual(cx.axes.length, 3, 'a primary cat + value axis plus the secondary percentage axis')
	})

	test('treemap reads its hierarchical leaf categories and carries no axes', async () => {
		const { presentation } = await authorRead((pres) => {
			// labels is string[][]: one array per hierarchy level, leaf-first — so the
			// leaf regions are level 0 and their parent continents are level 1.
			pres.addSlide().addChart(
				[
					{
						name: 'Sales',
						labels: [
							['US', 'CA', 'DE'],
							['NA', 'NA', 'EU'],
						],
						values: [10, 20, 30],
					},
				],
				{ type: ChartType.treemap, x: 1, y: 1, w: 5, h: 3 }
			)
		})
		const cx = firstChartEx(presentation)
		assertEqual(cx.layoutId, 'treemap', 'the treemap layout token')
		// The writer emits levels leaf-first, so the first cx:lvl is the leaf labels.
		assertEqual(cx.categories.join(','), 'US,CA,DE', 'the leaf level of the category hierarchy reads')
		assertEqual(cx.series[0].values.join(','), '10,20,30', 'the size values resolve')
		assertEqual(cx.axes.length, 0, 'a hierarchical treemap is genuinely axis-free')
	})

	test('a classic chart is unaffected — hasChartEx is false and chartEx is null', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
				type: ChartType.bar,
				x: 1,
				y: 1,
				w: 5,
				h: 3,
			})
		})
		const frame = firstShape(presentation, (s) => s.shapeType === 'graphicFrame')
		assertEqual(frame.hasChart, true, 'a classic chart still reports hasChart')
		assertEqual(frame.hasChartEx, false, 'and is not mistaken for a chartEx chart')
		assert(frame.chartEx === null, 'the chartEx getter is null for a classic chart')
		assert(frame.chart !== null, 'the classic chart getter still resolves')
	})

	test.skipIf(!validatorInstalled)('an authored chartEx deck is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'Cash Flow', labels: ['Start', 'End'], values: [100, 190] }], {
				type: ChartType.waterfall,
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				showTitle: true,
				title: 'Cash',
				subtotals: [0, 1],
			})
		})
		// chartEx has one expected, whitelisted validator complaint: `cx:axisId/@val`
		// (PowerPoint requires the attribute form the SDK schema doesn't declare — see
		// src/gen/chart/chartex-xml.ts). Everything else must be clean.
		const errors = await schemaErrors(buf)
		const unexpected = errors.filter((e) => !/axisId/.test(JSON.stringify(e)))
		assertEqual(unexpected.length, 0, `no unexpected schema violations, got ${JSON.stringify(unexpected)}`)
	})
})
