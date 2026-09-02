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
		// Three sibling axes had three rules for the same pair of options: the value axis emitted
		// them unconditionally, the category axis only behind a format code or an XY chart, and the
		// series axis only behind a format code. So this deck used to emit exactly one element.
		name: 'all three axes emit their numeric major/minor units',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.bar3d,
					catAxisMajorUnit: 3,
					catAxisMinorUnit: 1,
					serAxisMajorUnit: 2,
					valAxisMajorUnit: 4,
					valAxisMinorUnit: 2,
				})
			})
			const xml = await chartXml(zip)
			const units = [...xml.matchAll(/<c:(major|minor)Unit val="(\d+)"\/>/g)].map((m) => `${m[1]}:${m[2]}`)
			assertEqual(
				JSON.stringify(units.sort()),
				JSON.stringify(['major:2', 'major:3', 'major:4', 'minor:1', 'minor:2'].sort()),
				'every axis unit reaches the part; got ' + JSON.stringify(units)
			)
		},
	},
	{
		name: 'the time units stay behind their format code',
		fn: async () => {
			// Their gate is real: PowerPoint auto-adjusts them once it has the date bounds, and
			// they belong to a `c:dateAx`. Only the numeric siblings came out from behind it.
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, catAxisMajorTimeUnit: 'months' })
			})
			assertNotIncludes(await chartXml(zip), '<c:majorTimeUnit', 'no format code, no date axis, no time unit')
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
])
