// Unit tests for src/read/api/chart.ts driven by synthetic `c:chartSpace` XML.
//
// The fixture-based suite (chart.test.js) exercises the happy path against
// mixed.pptx. This file feeds hand-authored chart parts straight into the read
// model via the exported `Part` + `Chart` classes, covering the edge branches a
// single real deck rarely contains all at once: empty plot areas, missing /
// auto titles, unnamed series, inline numLit/strLit caches, sparse `c:pt/@idx`
// points, and non-numeric cached values. No fixture .pptx is needed — a `Part`
// is just (partName, contentType, bytes), and `new Chart(part)` reads its DOM.

import { describe, test } from 'vitest'
import { Chart, Part } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const NS =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

/** Wrap chart-space inner XML in a `Part` and hand it to a read-model `Chart`. */
function chart(inner) {
	const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace ${NS}>${inner}</c:chartSpace>`
	return new Chart(new Part('/ppt/charts/chart1.xml', CHART_CONTENT_TYPE, new TextEncoder().encode(xml)))
}

/** A `c:pt` cache point; pass `null` for a point with no `c:v` child. */
function pt(idx, v) {
	return v === null ? `<c:pt idx="${idx}"/>` : `<c:pt idx="${idx}"><c:v>${v}</c:v></c:pt>`
}

describe('Chart read model — plot-area / chart-type edges', () => {
	test('combo chart lists group types in document order; interleaving text nodes are skipped', () => {
		// Newlines between the groups become text nodes in the plot area — the
		// group scan must skip non-element nodes (nodeType !== 1).
		const c = chart(`<c:chart><c:plotArea>
			<c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
			<c:lineChart><c:ser><c:idx val="1"/></c:ser></c:lineChart>
		</c:plotArea></c:chart>`)
		assertEqual(c.chartTypes.join(','), 'bar,line', 'combo chart types in order')
		assertEqual(c.chartType, 'bar', 'first group is the chartType')
		assertEqual(c.series.length, 2, 'series flattened across both groups')
	})

	test('empty plot area yields no types, a null chartType, and no categories', () => {
		const c = chart(`<c:chart><c:plotArea/></c:chart>`)
		assertEqual(c.chartTypes.length, 0, 'no chart groups')
		assertEqual(c.chartType, null, 'chartType is null on an empty plot area')
		assertEqual(c.series.length, 0, 'no series')
		assertEqual(c.categories.length, 0, 'categories empty with no first series')
	})

	test('missing c:chart and missing c:plotArea both degrade to empty, not throw', () => {
		const noChart = chart(`<c:spPr/>`)
		assertEqual(noChart.chartType, null, 'no c:chart → null chartType')
		assertEqual(noChart.title, null, 'no c:chart → null title')
		assertEqual(noChart.series.length, 0, 'no c:chart → no series')

		const noPlot = chart(`<c:chart/>`)
		assertEqual(noPlot.chartType, null, 'c:chart but no c:plotArea → null chartType')
		assertEqual(noPlot.series.length, 0, 'no plot area → no series')
	})

	test('element_ exposes the chartSpace document element', () => {
		const c = chart(`<c:chart><c:plotArea/></c:chart>`)
		assertEqual(c.element_.localName, 'chartSpace', 'element_ is the c:chartSpace root')
	})
})

describe('Chart read model — title edges', () => {
	test('rich title text is concatenated across a:t runs', () => {
		const c = chart(
			`<c:chart><c:title><c:tx><c:rich>` +
				`<a:p><a:r><a:t>Break</a:t></a:r><a:r><a:t>even</a:t></a:r></a:p>` +
				`</c:rich></c:tx></c:title><c:plotArea/></c:chart>`
		)
		assertEqual(c.title, 'Breakeven', 'title concatenates run text')
	})

	test('title with no c:tx is null (auto title)', () => {
		const c = chart(`<c:chart><c:title/><c:plotArea/></c:chart>`)
		assertEqual(c.title, null, 'c:title without c:tx/c:rich → null')
	})

	test('title with c:tx but no c:rich is null', () => {
		const c = chart(`<c:chart><c:title><c:tx/></c:title><c:plotArea/></c:chart>`)
		assertEqual(c.title, null, 'c:tx without c:rich → null')
	})

	test('rich title that is empty text collapses to null', () => {
		const c = chart(
			`<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t></a:t></a:r></a:p></c:rich></c:tx></c:title>` +
				`<c:plotArea/></c:chart>`
		)
		assertEqual(c.title, null, 'empty rich text → null, not ""')
	})
})

describe('ChartSeries read model — name / index / caches', () => {
	/** A one-group chart whose single `c:ser` is `serInner`; returns that series. */
	function series(serInner) {
		return chart(`<c:chart><c:plotArea><c:barChart><c:ser>${serInner}</c:ser></c:barChart></c:plotArea></c:chart>`)
			.series[0]
	}

	test('index reads c:idx/@val and is null when absent', () => {
		assertEqual(series(`<c:idx val="4"/>`).index, 4, 'index from c:idx')
		assertEqual(series(`<c:order val="0"/>`).index, null, 'no c:idx → null index')
	})

	test('name: null without c:tx, direct c:v wins, else falls back to the cached first point', () => {
		assertEqual(series(``).name, null, 'no c:tx → null name')
		assertEqual(series(`<c:tx><c:v>Direct</c:v></c:tx>`).name, 'Direct', 'direct c:tx/c:v name')
		assertEqual(
			series(`<c:tx><c:strRef><c:strCache><c:ptCount val="1"/>${pt(0, 'Cached')}</c:strCache></c:strRef></c:tx>`).name,
			'Cached',
			'name falls back to the strCache first point'
		)
	})

	test('values parse numRef caches; empty and non-numeric points become null', () => {
		const s = series(
			`<c:val><c:numRef><c:numCache><c:ptCount val="4"/>` +
				`${pt(0, '360000')}${pt(1, 'abc')}${pt(2, '')}${pt(3, null)}` +
				`</c:numCache></c:numRef></c:val>`
		)
		assertEqual(s.values.length, 4, 'four values sized by ptCount')
		assertEqual(s.values[0], 360000, 'numeric point parsed')
		assertEqual(s.values[1], null, 'non-numeric point → null')
		assertEqual(s.values[2], null, 'empty point → null')
		assertEqual(s.values[3], null, 'point with no c:v → null')
	})

	test('categories read a strRef cache, ordered by idx', () => {
		const s = series(
			`<c:cat><c:strRef><c:strCache><c:ptCount val="2"/>${pt(0, 'Q1')}${pt(1, 'Q2')}</c:strCache></c:strRef></c:cat>`
		)
		assertEqual(s.categories.join(','), 'Q1,Q2', 'category labels as written')
	})

	test('inline numLit / strLit caches (no workbook ref) are honored', () => {
		const vals = series(`<c:val><c:numLit><c:ptCount val="2"/>${pt(0, '10')}${pt(1, '20')}</c:numLit></c:val>`).values
		assertEqual(vals.join(','), '10,20', 'inline numLit values')
		const cats = series(`<c:cat><c:strLit><c:ptCount val="1"/>${pt(0, 'Only')}</c:strLit></c:cat>`).categories
		assertEqual(cats.join(','), 'Only', 'inline strLit categories')
	})

	test('a c:val with no cache/ref/lit yields an empty value list', () => {
		assertEqual(series(`<c:val/>`).values.length, 0, 'no cache → empty values')
	})

	test('sparse points without ptCount grow the array to the highest idx+1', () => {
		// No c:ptCount; a single point at idx 3 must produce a length-4 array
		// with the intervening slots null.
		const s = series(`<c:val><c:numRef><c:numCache>${pt(3, '99')}</c:numCache></c:numRef></c:val>`)
		assertEqual(s.values.length, 4, 'array grown to highest idx + 1')
		assertEqual(s.values[3], 99, 'the sparse point lands at its idx')
		assertEqual(s.values[0], null, 'unfilled leading slots are null')
	})

	test('element_ exposes the underlying c:ser element', () => {
		assertEqual(series(`<c:idx val="0"/>`).element_.localName, 'ser', 'series element_ is c:ser')
	})
})

describe('ChartSeries read model — whether a fill is there', () => {
	/** A one-group chart whose single `c:ser` carries `spPrInner` as its `c:spPr`; returns that series. */
	function seriesWithSpPr(spPrInner) {
		return chart(
			`<c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>` +
				`<c:spPr>${spPrInner}</c:spPr>` +
				`</c:ser></c:barChart></c:plotArea></c:chart>`
		).series[0]
	}

	// A fill is there when the element is there, whatever colour form it holds. The check used to
	// ask whether the decoded `colorRef` carried an `srgb` or a `scheme` value -- two of the three
	// literal forms `readColorRef` reads, and none of the forms it does not -- so a series filled
	// with an `a:prstClr` reported no fill at all and re-authoring it dropped a fill the file states.
	//
	// Synthetic rather than a PowerPoint fixture, and deliberately so: `a:prstClr` and `a:sysClr`
	// are legal in `a:solidFill` but are not what PowerPoint's colour picker writes for a series.
	// Every chart across chart-series-shapes.pptx and mixed.pptx uses `a:srgbClr` or `a:schemeClr`.
	// The ground truth here is the schema, not PowerPoint's authoring behaviour, so there is no
	// fixture to author and this file's own header already covers hand-authored chart parts.
	test('a preset colour is a fill, and its name is reported', () => {
		const s = seriesWithSpPr('<a:solidFill><a:prstClr val="black"/></a:solidFill>')
		assert(s.fill, 'a prstClr fill is a fill')
		assertEqual(s.fill.colorRef.preset, 'black', 'the preset name is reported')
		assertEqual(s.fill.colorRef.srgb, null, 'and it is not an srgb colour')
		assertEqual(s.fill.noFill, false, 'a coloured series is not noFill')
	})

	test('a colour form the model does not decode is still a fill', () => {
		// `a:sysClr` has no field on `ColorRef`, so every colour field stays null -- but the file
		// states a fill, and reporting `null` would claim it states none.
		const s = seriesWithSpPr('<a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill>')
		assert(s.fill, 'a sysClr fill is still a fill')
		assertEqual(s.fill.colorRef.srgb, null, 'sysClr is not decoded into a literal')
		assertEqual(s.fill.colorRef.scheme, null, 'nor into a scheme token')
	})

	test('the forms that already worked still do', () => {
		const srgb = seriesWithSpPr('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
		assertEqual(srgb.fill.colorRef.srgb, 'FF0000', 'an srgb fill')
		const scheme = seriesWithSpPr('<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>')
		assertEqual(scheme.fill.colorRef.scheme, 'accent1', 'a scheme fill')
		const none = seriesWithSpPr('<a:noFill/>')
		assert(none.fill, 'an explicit noFill is reported')
		assertEqual(none.fill.noFill, true, 'as noFill')
	})

	test('a c:spPr that declares no fill choice still reads null', () => {
		// The regression guard: it is easy to widen the test into reporting a fill for everything.
		assertEqual(seriesWithSpPr('<a:ln w="12700"/>').fill, null, 'a line-only c:spPr inherits the theme')
		assertEqual(
			chart(`<c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea></c:chart>`)
				.series[0].fill,
			null,
			'a series with no c:spPr at all'
		)
	})
})
