import TsPptx, { ChartType, InvalidOptionError } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
	captureDiagnostics,
} from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// addChart normalizes/validates several numeric and enum options before emitting.
// These branches (clamping an out-of-range line-marker size; dropping invalid
// gridLine size/style/cap so PowerPoint-invalid values never reach the XML) are
// exercised here through the public API so the generated chart part proves the
// scrub actually happened.

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
const BASE = { x: 1, y: 1, w: 6, h: 3 }

defineRegressionSuite('Chart option validation', [
	{
		name: 'lineDataSymbolSize above the 2-72 range is clamped to 72',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.line,
					lineDataSymbol: 'circle',
					lineDataSymbolSize: 999,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:size val="72"/>', 'oversize marker clamps to the 72 max')
			assertNotIncludes(xml, '<c:size val="999"/>', 'the out-of-range value must not reach the XML')
		},
	},
	{
		name: 'lineDataSymbolSize below the range is clamped to 2',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.line,
					lineDataSymbol: 'circle',
					lineDataSymbolSize: 1,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:size val="2"/>', 'undersize marker clamps to the 2 min')
		},
	},
	{
		name: 'invalid gridLine cap is dropped (never emitted)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.bar,
					valGridLine: { size: 2, style: 'solid', cap: 'INVALID' },
				})
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, 'INVALID', 'an unrecognized gridLine cap must be scrubbed before emit')
		},
	},
	{
		name: 'invalid gridLine style is dropped',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, valGridLine: { size: 2, style: 'wavy' } })
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, 'wavy', 'an unrecognized gridLine style must be scrubbed before emit')
		},
	},
	{
		// The two chart emitters partition `ChartType` between them and each treats the other's
		// members as not its own, so a `type` outside the catalog matches no arm in either and
		// cannot produce a plot. It has to be refused at the boundary, where the caller can see
		// which call was wrong — reaching an emitter with it used to yield a `<c:plotArea>` with
		// axes and nothing in it, i.e. a chart-shaped hole the deck opens and shows empty.
		name: 'a chart type outside the catalog is refused at addChart, not emitted as an empty plot',
		fn: () => {
			let thrown = null
			try {
				// The cast is the point: `CHART_NAME` keeps TypeScript callers out, and this guard exists
				// for the JavaScript ones it cannot reach.
				new TsPptx().addSlide().addChart(SERIES, /** @type {never} */ ({ ...BASE, type: 'nonsense' }))
			} catch (err) {
				thrown = err
			}
			assert(thrown instanceof InvalidOptionError, `an unknown chart type throws InvalidOptionError (got ${thrown})`)
			assertEqual(thrown.code, 'chart/unknown-type', 'the condition carries its own code')
			assertEqual(thrown.detail.type, 'nonsense', 'the offending type is carried as structured detail')
		},
	},
	{
		name: 'a combo subchart type outside the catalog is refused the same way',
		fn: () => {
			// The combo form routes each entry's `type` through the same emitters, so the guard has
			// to see the `ChartMulti[]` entries too, not just the single-type `options.type`.
			let thrown = null
			try {
				new TsPptx().addSlide().addChart(
					/** @type {never} */ ([
						{ type: ChartType.bar, data: SERIES },
						{ type: 'nonsense', data: SERIES },
					]),
					BASE
				)
			} catch (err) {
				thrown = err
			}
			assert(thrown instanceof InvalidOptionError, `an unknown subchart type throws InvalidOptionError (got ${thrown})`)
			assertEqual(thrown.code, 'chart/unknown-type', 'the condition carries its own code')
		},
	},
	{
		name: 'every catalog type is accepted and reaches an emitter that can build it',
		fn: async () => {
			// Two guarantees, both about drift between the `ChartType` enum and the code around it.
			//
			// Against the boundary guard: a `ChartType` member `isChartType` does not know would be
			// refused by `addChart` even though both emitters can build it.
			//
			// Against emitter routing: `makeChartType` and `chartExLayoutId` partition the catalog
			// between them by `switch`, so neither is exhaustive over `ChartType` and neither can be
			// made compiler-enforced with a `never` arm. Each throws `chart/type-not-routed` on a
			// member it has no case for, so building the whole catalog here is what catches an added
			// member nobody routed — verified by adding a temporary enum member, which fails this
			// test from `makeChartType`, and from `chartExLayoutId` once it is added to CHARTEX_TYPES.
			// This is the only gate on that; `typescript/switch-exhaustiveness-check` is off, for the
			// reasons recorded in `.oxlintrc.jsonc`.
			for (const type of Object.values(ChartType)) {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, { ...BASE, type })
				})
				assert(zip, `${type} builds through addChart`)
			}
		},
	},
	{
		name: 'legendFontSize reaches the legend txPr in hundredths',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, showLegend: true, legendFontSize: 14 })
			})
			assertIncludes(await chartXml(zip), '<a:defRPr sz="1400">', 'the legend font size is emitted in hundredths')
		},
	},
	{
		// It used to be the one font-size option whose emitter wrapped the value in `Number()`,
		// so a string worked here while the same string threw at every other spelling of the
		// same option. Pinned from untyped JS, which is the only place it can now arrive.
		name: 'a string legendFontSize is refused rather than coerced',
		fn: async () => {
			let thrown = null
			try {
				await build((p) => {
					p.addSlide().addChart(SERIES, {
						...BASE,
						type: ChartType.bar,
						showLegend: true,
						legendFontSize: /** @type {never} */ ('14'),
					})
				})
			} catch (err) {
				thrown = err
			}
			assert(thrown instanceof InvalidOptionError, `a non-number font size throws InvalidOptionError (got ${thrown})`)
			assertEqual(thrown.code, 'coord/non-finite', 'the converter refuses it with its own code')
		},
	},
	{
		name: 'non-positive gridLine size is dropped so defaults apply',
		fn: async () => {
			// A negative size with a real style takes the `size <= 0` branch (the
			// early `style === 'none'` return would otherwise skip validation).
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, valGridLine: { size: -5, style: 'dash' } })
			})
			const xml = await chartXml(zip)
			assert(xml.length > 0, 'chart still builds after dropping the invalid gridLine size')
		},
	},
	{
		// `typeof x === 'number'` is the one numeric guard `NaN` passes, and it was the guard on
		// both axis-crossing decisions while every other numeric axis option used truthiness.
		name: 'a non-finite axis crossing falls back to the rule instead of emitting NaN',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, {
						...BASE,
						type: ChartType.bar,
						valAxisCrossesAt: NaN,
						catAxisCrossesAt: NaN,
					})
				})
				return chartXml(zip)
			})
			assertNotIncludes(xml, 'val="NaN"', 'ST_Double has no NaN')
			assertIncludes(xml, '<c:crosses val="autoZero"/>', 'the axis falls back to its default rule')
			assertEqual(
				codes.filter((c) => c === 'chart/option-out-of-range').length,
				2,
				'both axes say so; got ' + JSON.stringify(codes)
			)
		},
	},
	{
		name: 'a finite axis crossing still emits crossesAt',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, valAxisCrossesAt: 2 })
			})
			assertIncludes(await chartXml(zip), '<c:crossesAt val="2"/>', 'an explicit position is honoured')
		},
	},
	{
		// `ChartOpts.x` is a `Coord`, and the title builder took it through an `as number` cast,
		// so a string reached the layout arithmetic and `+` concatenated instead of adding.
		name: 'a non-numeric chart `x` does not put NaN in the title layout',
		fn: async () => {
			// A percentage needs the slide axis and the chart part is built without a layout, so
			// the chart's own offset is left out of the fold and the caller is told. What must not
			// happen is the old outcome: string concatenation, then `<c:x val="NaN"/>`.
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, {
						...BASE,
						x: '10%',
						type: ChartType.bar,
						showTitle: true,
						title: 'T',
						titlePos: { x: 0.5, y: 0.5 },
					})
				})
				return chartXml(zip)
			})
			assertNotIncludes(xml, 'val="NaN"', 'no axis of the manual layout is NaN')
			assertIncludes(xml, '<c:xMode val="edge"/>', 'the caller still gets the manual layout they asked for')
			assert(codes.includes('chart/option-out-of-range'), 'and is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a unit-bearing chart `x` folds into the title layout',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					x: '2in',
					type: ChartType.bar,
					showTitle: true,
					title: 'T',
					titlePos: { x: 0.5, y: 0.5 },
				})
			})
			const inches = await chartXml(zip)
			const { zip: zip2 } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					x: 2,
					type: ChartType.bar,
					showTitle: true,
					title: 'T',
					titlePos: { x: 0.5, y: 0.5 },
				})
			})
			assertEqual(
				/<c:x val="([^"]+)"/.exec(inches)?.[1],
				/<c:x val="([^"]+)"/.exec(await chartXml(zip2))?.[1],
				'"2in" and 2 are the same coordinate'
			)
		},
	},
	{
		// `ST_Skip` is an `xsd:unsignedInt` of at least 1; the option was typed as a free-form
		// string and emitted verbatim, and the type also rejected the natural `2`.
		name: 'a tick-label frequency that is not a positive integer is dropped with a warning',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, catAxisLabelFrequency: 'every other' })
				})
				return chartXml(zip)
			})
			assertNotIncludes(xml, '<c:tickLblSkip', 'nothing outside ST_Skip reaches the attribute')
			assert(codes.includes('chart/option-out-of-range'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a numeric tick-label frequency is emitted',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, catAxisLabelFrequency: 2 })
			})
			assertIncludes(await chartXml(zip), '<c:tickLblSkip val="2"/>', 'every other label')
		},
	},
	{
		// `majorUnit`/`minorUnit` belong to `CT_ValAx` and `CT_DateAx`. `CT_CatAx` has no slot for
		// either, so a plain category axis takes neither -- and says so rather than writing an
		// element PowerPoint then refuses to open the deck over.
		name: 'a plain category axis has no unit slot, and the caller is told',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, {
						...BASE,
						type: ChartType.bar3d,
						catAxisMajorUnit: 3,
						catAxisMinorUnit: 1,
						valAxisMajorUnit: 4,
						valAxisMinorUnit: 2,
					})
				})
				return chartXml(zip)
			})
			const catAx = xml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)
			assert(catAx, 'expected a <c:catAx> block; got: ' + xml)
			assertNotIncludes(catAx[0], '<c:majorUnit', 'CT_CatAx has no majorUnit')
			assertNotIncludes(catAx[0], '<c:minorUnit', 'CT_CatAx has no minorUnit')
			const units = [...xml.matchAll(/<c:(major|minor)Unit val="(\d+)"\/>/g)].map((m) => `${m[1]}:${m[2]}`)
			assertEqual(
				JSON.stringify(units.sort()),
				JSON.stringify(['major:4', 'minor:2'].sort()),
				'only the value axis carries units; got ' + JSON.stringify(units)
			)
			assert(codes.includes('chart/option-not-on-axis'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a date axis interleaves its numeric and time units in schema order',
		fn: async () => {
			// `CT_DateAx` orders them baseTimeUnit, majorUnit, majorTimeUnit, minorUnit,
			// minorTimeUnit. Emitting the three time units and then the two numeric ones is a
			// content-model violation even though every element is legal on the type.
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.line,
					catLabelFormatCode: 'yyyy-mm-dd',
					catAxisBaseTimeUnit: 'days',
					catAxisMajorTimeUnit: 'months',
					catAxisMinorTimeUnit: 'years',
					catAxisMajorUnit: 3,
					catAxisMinorUnit: 1,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(
				xml,
				'<c:baseTimeUnit val="days"/><c:majorUnit val="3"/><c:majorTimeUnit val="months"/>' +
					'<c:minorUnit val="1"/><c:minorTimeUnit val="years"/>',
				'the five units come out in CT_DateAx order'
			)
		},
	},
	{
		name: 'the time units stay behind their format code',
		fn: async () => {
			// Their gate is real: PowerPoint auto-adjusts them once it has the date bounds, and
			// they belong to a `c:dateAx`. Without a format code there is no date axis to hold one.
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, catAxisMajorTimeUnit: 'months' })
				})
				return chartXml(zip)
			})
			assertNotIncludes(xml, '<c:majorTimeUnit', 'no format code, no date axis, no time unit')
			assert(codes.includes('chart/option-not-on-axis'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a scatter X axis is a value axis and carries the numeric units',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', values: [1, 2, 3] }], {
					...BASE,
					type: ChartType.scatter,
					catAxisMajorUnit: 2,
					catAxisMinorUnit: 1,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:majorUnit val="2"/>', 'the X value axis takes catAxisMajorUnit')
			assertIncludes(xml, '<c:minorUnit val="1"/>', 'and catAxisMinorUnit')
		},
	},
	{
		// `axisPos` was declared on `ChartPropsBase` and read by nothing: the only `axisPos` in
		// `src/` is a local in `makeValAxis` computed from `barDir` and the axis id. Per-axis
		// placement wants `catAxisLabelPos`-style naming, not one key shared across three axes.
		name: 'axisPos placed nothing, which is why it could be removed',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, axisPos: 't' })
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:axPos val="b"/>', 'the category axis is still placed from barDir')
			assertIncludes(xml, '<c:axPos val="l"/>', 'and so is the value axis')
			assertNotIncludes(xml, '<c:axPos val="t"/>', 'the option never placed anything')
		},
	},
	{
		name: 'a chart option that is not a number throws instead of silently taking the default',
		fn: async () => {
			// One policy for an out-of-range number, stated on `clampRangedInput`: a finite value
			// has a nearest legal neighbour, so it clamps and warns; a value that is not a number
			// has none, so the request is discarded and that throws. The chart clamp answered
			// `undefined` instead -- discarding the request and reporting nothing -- so
			// `holeSize: NaN` silently took the default while `holeSize: 200` warned. Same option,
			// same class of mistake, two behaviours.
			for (const [option, value] of [
				['holeSize', NaN],
				['barGapWidthPct', NaN],
				['firstSliceAng', /** @type {never} */ ('90')],
				['lineDataSymbolSize', NaN],
			]) {
				let thrown = null
				try {
					const pres = new TsPptx()
					pres.addSlide().addChart(SERIES, { ...BASE, type: ChartType.doughnut, [option]: value })
				} catch (err) {
					thrown = err
				}
				assert(thrown instanceof InvalidOptionError, `${option} must throw an InvalidOptionError`)
				assertEqual(thrown.code, 'chart/option-non-finite', `${option} carries the shared code`)
			}
		},
	},
	{
		name: 'a fractional chart option is rounded, and the caller is told',
		fn: async () => {
			// These are integer schema types, so `holeSize: 42.5` is as much a correction as
			// `holeSize: 200` -- and it is in range, so it is the case a bounds check alone misses.
			const { codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.doughnut, holeSize: 42.5 })
				})
			)
			assertIncludes(codes, 'chart/option-out-of-range')
		},
	},
	{
		name: 'seriesOptions on a chart type that colours points, not series, warns rather than doing nothing',
		fn: async () => {
			// A pie colours *points* and a surface colours bands, so a per-series override has no
			// referent on either even in principle. "The caller said it and nothing happened" is the
			// state the option rules forbid.
			const { codes, messages } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.pie, seriesOptions: [{ color: 'FF0000' }] })
				})
			)
			assertIncludes(codes, 'chart/option-not-supported')
			assert(
				messages.some((m) => m.includes('`color`')),
				'the warning names the field that will be dropped; got ' + JSON.stringify(messages)
			)
		},
	},
	{
		name: 'and stays quiet on the plots that do read it',
		fn: async () => {
			for (const type of [
				ChartType.bar,
				ChartType.line,
				ChartType.radar,
				ChartType.area,
				ChartType.scatter,
				ChartType.bubble,
			]) {
				const { codes } = await captureDiagnostics(() =>
					build((p) => {
						p.addSlide().addChart(SERIES, { ...BASE, type, seriesOptions: [{ color: 'FF0000' }] })
					})
				)
				assert(
					!codes.includes('chart/option-not-supported'),
					`${type} reads seriesOptions.color; got ` + JSON.stringify(codes)
				)
			}
		},
	},
	{
		// The warning is per FIELD, not per chart type: a type that reads `color` can still drop
		// `lineSize`, because a bar series takes its outline from `dataBorder` and never reaches
		// `seriesStroke`. Type-level checking called this supported and dropped it in silence.
		name: 'a field the plot cannot resolve warns even when the type reads seriesOptions',
		fn: async () => {
			const { codes, messages } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart(SERIES, {
						...BASE,
						type: ChartType.bar,
						seriesOptions: [{ color: 'FF0000', lineSize: 3 }],
					})
				})
			)
			assertIncludes(codes, 'chart/option-not-supported')
			assert(
				messages.some((m) => m.includes('`lineSize`') && !m.includes('`color`')),
				'only `lineSize` should be reported as dropped; got ' + JSON.stringify(messages)
			)
		},
	},
	{
		// A stock chart's price series draw no line by design and its `<c:dLbls>` is a constant, so
		// `color` is the one field with a referent -- and it has two: the volume bar and the close
		// marker.
		name: 'a stock chart reads only the colour of a series override',
		fn: async () => {
			const { codes, messages } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart(SERIES, {
						...BASE,
						type: ChartType.stock,
						seriesOptions: [{ color: 'FF0000', dataLabelFontSize: 9 }],
					})
				})
			)
			assertIncludes(codes, 'chart/option-not-supported')
			assert(
				messages.some((m) => m.includes('`dataLabelFontSize`') && !m.includes('`color`')),
				'only the data-label field should be reported as dropped; got ' + JSON.stringify(messages)
			)
		},
	},
])
