/**
 * ts-pptx: Chart DrawingML Assembly
 *
 * Builds a chart's `ppt/charts/chartN.xml` -- the `<c:chartSpace>` DrawingML that
 * PowerPoint renders. `makeXmlCharts` assembles the top-level envelope (header, plot
 * area, axes region, legend, metadata) and `makeChartType` dispatches the plot itself to
 * the per-family builder: {@link ./plot-cat-axis} for area/bar/bar3D/line/radar,
 * {@link ./plot-scatter}, {@link ./plot-bubble}, {@link ./plot-pie}. Axes come from
 * {@link ./chart-axes} and the shared leaf fragments from {@link ./chart-parts}.
 *
 * Every function in this directory is a pure string builder -- no I/O, no mutation of
 * the presentation model. The `<c:f>` series formulas here point back at the cells
 * written by the embedded workbook ({@link ./embed-xlsx}); that mapping lives in
 * {@link ./data-refs}.
 */

import { asChartType, ChartType } from '../../enums.js'
import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_SERIES_PRIMARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	DEF_FONT_SIZE,
	DEF_FONT_TITLE_SIZE,
	XML_DECL,
} from '../../constants-internal.js'
import type {
	ChartOptsInternal,
	ChartOptsOverrides,
	OptsChartDataInternal,
	SlideRelChart,
} from '../../types/internal.js'
import type { BorderProps, ShapeFillProps } from '../../types/index.js'
import { warn } from '../../diagnostics.js'
import { genXmlColorSelection, solidPaint } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { lineWidthToEmu } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { createChartTextFonts, dimmedTextFill, dimmedTextLine, genXmlTitle } from './chart-parts.js'
import { makeCatAxis, makeSerAxis, makeValAxis } from './chart-axes.js'
import { makeCatAxisPlot } from './plot-cat-axis.js'
import { makeScatterPlot } from './plot-scatter.js'
import { makeBubblePlot } from './plot-bubble.js'
import { makePiePlot } from './plot-pie.js'
import { isVolumeStockStyle, makeStockPlot } from './plot-stock.js'
import { makeSurfacePlot, makeSurfaceScene } from './plot-surface.js'
import { InternalError, InvalidOptionError } from '../../errors.js'
import { isScatterChart, isXyChart } from './chart-kind.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'

/** The three namespaces every `<c:chartSpace>` declares, in the order PowerPoint writes them. */
const CHART_SPACE_NS = {
	'xmlns:c': OOXML_NS.c,
	'xmlns:a': OOXML_NS.a,
	'xmlns:r': OOXML_NS.r,
}

/**
 * Whether a chart-area/plot-area `fill` says anything the fill dispatch can act on.
 *
 * `c:spPr` is `a:CT_ShapeProperties` — the same optional `EG_FillProperties` group the
 * shape path writes into — so gradient, pattern and an omitted fill child are all legal
 * here. The gate used to be `fill?.color`, which meant every spelling carrying no colour
 * (`gradient`, `pattern`, `image`, `inherit`) fell to the no-fill arm and did nothing.
 *
 * The gate cannot simply be `fill != null`. `normalizeChartOptions` defaults
 * `plotArea.fill` to `{}`, so every chart in existence arrives here with a fill object,
 * and a presence check would paint each one a default grey. The same is true of the one
 * input a caller can write by hand: `{ transparency: 50 }` alone is not a fill, since
 * there is no colour for the alpha to apply to. Both stay no-fill.
 *
 * A payload with no `type` (`{ gradient }`, `{ pattern }`) is likewise not stated, which
 * matches the shape path: only `line` infers its type from a bare `gradient`, fills have
 * always wanted `type: 'gradient'` spelled out.
 */
function isStatedFill(fill?: ShapeFillProps): fill is ShapeFillProps {
	return Boolean(fill?.color || fill?.type)
}

/**
 * Resolve one combo subchart's overrides against the chart-level options.
 *
 * The override bag carries a third state the rest of the generator does not: a *present*
 * `undefined`, meaning "this subchart's own value was rejected, so emit nothing here", as against
 * an absent key, which inherits the chart-level value (see {@link ChartOptsOverrides}). The spread
 * is what applies that distinction — and once it has, the keys still holding `undefined` have said
 * everything they had to say, so they are dropped. What the plot builders receive is an ordinary
 * options bag with one spelling of absent.
 * @param chartOptions - the chart-level options, already normalized
 * @param overrides - this subchart's vetted overrides
 */
function resolveSubchartOptions(chartOptions: ChartOptsInternal, overrides: ChartOptsOverrides): ChartOptsInternal {
	const merged: Record<string, unknown> = { ...chartOptions, ...overrides }
	for (const key of Object.keys(merged)) {
		if (merged[key] === undefined) delete merged[key]
	}
	return merged
}

/**
 * The chart title (or the `autoTitleDeleted` flag that suppresses PowerPoint's default one)
 * followed by the 3-D scene, for the two families that plot into one.
 */
function makeChartHeaderXml(rel: SlideRelChart): string {
	// NOTE: `autoTitleDeleted` is emitted in both arms — without it PowerPoint creates a default
	// chart title even when `showTitle` is false.
	const title = rel.opts.showTitle
		? genXmlTitle(
				{
					title: rel.opts.title || 'Chart Title',
					color: rel.opts.titleColor,
					fontFace: rel.opts.titleFontFace,
					fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
					titleAlign: rel.opts.titleAlign,
					titleBold: rel.opts.titleBold,
					titleItalic: rel.opts.titleItalic,
					titleUnderline: rel.opts.titleUnderline,
					titlePos: rel.opts.titlePos,
					titleRotate: rel.opts.titleRotate,
				},
				rel.opts.x as number,
				rel.opts.y as number
			) + voidEl('c:autoTitleDeleted', { val: 0 })
		: voidEl('c:autoTitleDeleted', { val: 1 })
	// 3-D view: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_perspective_topic_ID0E6BUQB.html
	if (rel.opts._type === ChartType.bar3d) {
		return (
			title +
			el('c:view3D', null, [
				raw(voidEl('c:rotX', { val: rel.opts.v3DRotX })),
				raw(voidEl('c:rotY', { val: rel.opts.v3DRotY })),
				raw(voidEl('c:rAngAx', { val: !rel.opts.v3DRAngAx ? 0 : 1 })),
				raw(voidEl('c:perspective', { val: rel.opts.v3DPerspective })),
			])
		)
	}
	// A surface chart is a 3-D scene: view3D + floor/side/back walls precede the plotArea.
	if (rel.opts._type === ChartType.surface) return title + makeSurfaceScene(rel.opts)
	return title
}

/**
 * The plot area's `<c:layout>`. Manual layout is emitted only when the caller asked for one:
 * an empty `<c:layout/>` lets PowerPoint auto-fit, which it does well across all four TRBL
 * legend positions.
 */
function makePlotAreaLayoutXml(rel: SlideRelChart): string {
	const layout = rel.opts.layout
	if (!layout) return voidEl('c:layout')
	const manualLayout = el('c:manualLayout', null, [
		// Every space here sits before `/>`, INSIDE the tag rather than between elements, so
		// it is out of scope for the flatten (docs/chart-whitespace-flatten.md) and stays.
		// `prove-whitespace` freezes intra-tag whitespace, and reported these when the
		// codemod first took them — which is the whole reason it looks at them.
		raw(voidEl('c:layoutTarget', { val: 'inner' }, { closePrefix: ' ' })),
		raw(voidEl('c:xMode', { val: 'edge' }, { closePrefix: ' ' })),
		raw(voidEl('c:yMode', { val: 'edge' }, { closePrefix: ' ' })),
		raw(voidEl('c:x', { val: layout.x || 0 }, { closePrefix: ' ' })),
		raw(voidEl('c:y', { val: layout.y || 0 }, { closePrefix: ' ' })),
		raw(voidEl('c:w', { val: layout.w || 1 }, { closePrefix: ' ' })),
		raw(voidEl('c:h', { val: layout.h || 1 }, { closePrefix: ' ' })),
	])
	return el('c:layout', null, raw(manualLayout))
}

/**
 * Build the category/value/series axis XML (empty for pie/doughnut). Resolves combo-chart
 * category axes to val axes when owned by a scatter/bubble subchart, per the tracked flags.
 */
function makeChartAxesXml(
	rel: SlideRelChart,
	usesSecondaryValAxis: boolean,
	usesSecondaryCatAxis: boolean,
	primaryCatAxisValType: ChartType | null,
	secondaryCatAxisValType: ChartType | null,
	primaryCatAxisHasCategoryChart: boolean,
	secondaryCatAxisHasCategoryChart: boolean
): string {
	let strXml = ''
	// Stock charts drive their own (date/category)+value axis pair. The non-volume styles use just
	// the primary pair (the generic path below would also produce it), but the volume styles add a
	// secondary pair for the price series (the volume bar owns the primary pair) — so route all stock
	// axes here, reusing the shared cat/val emitters. `barDir` defaults to 'col' in the define layer,
	// which already gives the correct axis positions (category at bottom, value at left).
	if (rel.opts._type === ChartType.stock) {
		let stockXml = makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY)
		stockXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY)
		if (isVolumeStockStyle(rel.opts.stockStyle)) {
			// Secondary value axis (right) for the price series, then a hidden secondary category axis.
			stockXml += makeValAxis(rel.opts, AXIS_ID_VALUE_SECONDARY)
			stockXml += makeCatAxis({ ...rel.opts, catAxisHidden: true }, AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_SECONDARY)
		}
		return stockXml
	}
	if (rel.opts._type !== ChartType.pie && rel.opts._type !== ChartType.doughnut) {
		// Param check
		if (rel.opts.valAxes && rel.opts.valAxes.length > 1 && !usesSecondaryValAxis) {
			throw new InvalidOptionError(
				'chart/secondary-axis-unused',
				'Secondary axis must be used by one of the multiple charts'
			)
		}

		// Resolve the effective `_type` for a combo category axis so scatter/bubble
		// subcharts get a `<c:valAx>` X axis. Returns the scatter/bubble type when
		// that axis is owned only by such a subchart, else null (category axis).
		const comboCatAxisType = (isSecondary: boolean): { _type: ChartType } | Record<string, never> => {
			const valType = isSecondary ? secondaryCatAxisValType : primaryCatAxisValType
			const hasCategoryChart = isSecondary ? secondaryCatAxisHasCategoryChart : primaryCatAxisHasCategoryChart
			if (!valType) return {}
			if (hasCategoryChart) {
				// A category-based chart and a scatter/bubble chart cannot share one
				// axis (one needs <c:catAx>, the other <c:valAx>). Keep the category
				// axis and warn rather than silently emit a repair-triggering file.
				warn(
					'chart/axis-type-conflict',
					`A category-based chart and a scatter/bubble chart cannot share the same ${isSecondary ? 'secondary' : 'primary'} category axis; emitting a category axis. Put the scatter/bubble series on a separate axis.`
				)
				return {}
			}
			return { _type: valType }
		}

		if (rel.opts.catAxes) {
			if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
				throw new InvalidOptionError(
					'chart/axis-count-mismatch',
					'There must be the same number of value and category axes.'
				)
			}
			strXml += makeCatAxis(
				{ ...rel.opts, ...rel.opts.catAxes[0], ...comboCatAxisType(false) },
				AXIS_ID_CATEGORY_PRIMARY,
				AXIS_ID_VALUE_PRIMARY
			)
		} else {
			strXml += makeCatAxis(
				{ ...rel.opts, ...comboCatAxisType(false) },
				AXIS_ID_CATEGORY_PRIMARY,
				AXIS_ID_VALUE_PRIMARY
			)
		}

		if (rel.opts.valAxes) {
			strXml += makeValAxis({ ...rel.opts, ...rel.opts.valAxes[0] }, AXIS_ID_VALUE_PRIMARY)
			if (rel.opts.valAxes[1]) {
				strXml += makeValAxis({ ...rel.opts, ...rel.opts.valAxes[1] }, AXIS_ID_VALUE_SECONDARY)
			}
		} else {
			strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY)

			// Add series axis for 3D bar and surface (both plot over a category × series grid)
			if (rel.opts._type === ChartType.bar3d || rel.opts._type === ChartType.surface) {
				strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY)
			}

			// For combo charts referencing a secondary value axis via the
			// `secondaryValAxis: true` flag (without a `valAxes` array),
			// auto-synthesise the missing secondary value axis def so that
			// the axId references in <c:plotArea> all resolve.
			if (usesSecondaryValAxis) {
				strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_SECONDARY)
			}
		}

		// Combo Charts: Add secondary axes after all vals
		if (rel.opts?.catAxes && rel.opts?.catAxes[1]) {
			strXml += makeCatAxis(
				{ ...rel.opts, ...rel.opts.catAxes[1], ...comboCatAxisType(true) },
				AXIS_ID_CATEGORY_SECONDARY,
				AXIS_ID_VALUE_SECONDARY
			)
		} else if (usesSecondaryCatAxis && (!rel.opts.catAxes || !rel.opts.catAxes[1])) {
			// Same as above for the secondary category axis.
			strXml += makeCatAxis(
				{ ...rel.opts, ...comboCatAxisType(true) },
				AXIS_ID_CATEGORY_SECONDARY,
				AXIS_ID_VALUE_SECONDARY
			)
		}
	}
	return strXml
}

/**
 * The `<c:spPr>` shared by the plot area and the chart space: a stated fill or none, a border
 * or an explicit no-line, and an empty effect list.
 */
function chartShapeProps(fill: ShapeFillProps | undefined, border: BorderProps | undefined): string {
	return el('c:spPr', null, [
		raw(isStatedFill(fill) ? genXmlColorSelection(fill) : voidEl('a:noFill')),
		raw(
			border
				? el(
						'a:ln',
						{ w: lineWidthToEmu(resolveBorderWidth(border, 1)), cap: 'flat' },
						raw(genXmlColorSelection(solidPaint(border.color ?? '363636', border.transparency)))
					)
				: el('a:ln', null, raw(voidEl('a:noFill')))
		),
		raw(voidEl('a:effectLst', null)),
	])
}

/**
 * The `<c:dTable>` grid drawn under the plot. Sits between the last `</c:valAx>` and the plot
 * area's own `<c:spPr>`.
 */
function makeDataTableXml(rel: SlideRelChart): string {
	const spPr = el('c:spPr', null, [
		raw(voidEl('a:noFill', null)),
		raw(dimmedTextLine(15000, 85000)),
		raw(voidEl('a:effectLst', null)),
	])
	const defRPr = el(
		'a:defRPr',
		{
			sz: ptToHundredths(rel.opts.dataTableFontSize || DEF_FONT_SIZE),
			b: 0,
			i: 0,
			u: 'none',
			strike: 'noStrike',
			kern: 1200,
			baseline: 0,
		},
		[
			raw(dimmedTextFill(65000, 35000)),
			raw(voidEl('a:latin', { typeface: '+mn-lt' })),
			raw(voidEl('a:ea', { typeface: '+mn-ea' })),
			raw(voidEl('a:cs', { typeface: '+mn-cs' })),
		]
	)
	const txPr = el('c:txPr', null, [
		raw(
			voidEl('a:bodyPr', {
				rot: 0,
				spcFirstLastPara: 1,
				vertOverflow: 'ellipsis',
				vert: 'horz',
				wrap: 'square',
				anchor: 'ctr',
				anchorCtr: 1,
			})
		),
		raw(voidEl('a:lstStyle', null)),
		raw(el('a:p', null, [raw(el('a:pPr', { rtl: 0 }, raw(defRPr))), raw(voidEl('a:endParaRPr', { lang: 'en-US' }))])),
	])
	return el('c:dTable', null, [
		raw(voidEl('c:showHorzBorder', { val: !rel.opts.showDataTableHorzBorder ? 0 : 1 })),
		raw(voidEl('c:showVertBorder', { val: !rel.opts.showDataTableVertBorder ? 0 : 1 })),
		// These two `val` attributes were written padded into a column with the two above, and that
		// padding is emitted bytes. `el()` writes exactly one space before an attribute, by design,
		// so the aligned pair stays hand-written rather than bending the builder around a cosmetic
		// quirk. Removing the padding is an output change, not a cleanup; the ratchet header says why.
		raw(`  <c:showOutline    val="${!rel.opts.showDataTableOutline ? 0 : 1}"/>`),
		raw(`  <c:showKeys       val="${!rel.opts.showDataTableKeys ? 0 : 1}"/>`),
		raw(spPr),
		raw(txPr),
	])
}

/**
 * The `<c:legend>`: position, the entries a combo chart deletes, an optional manual placement,
 * and the text properties when any legend font option is set.
 */
function makeLegendXml(rel: SlideRelChart): string {
	// For combo charts: suppress the series belonging to subcharts that set `showLegend: false`.
	let entries = ''
	if (Array.isArray(rel.opts._type)) {
		let seriesIdx = 0
		for (const type of rel.opts._type) {
			if (type.options?.showLegend === false) {
				for (let i = 0; i < type.data.length; i++) {
					entries += el('c:legendEntry', null, [
						raw(voidEl('c:idx', { val: seriesIdx + i })),
						raw(voidEl('c:delete', { val: 1 })),
					])
				}
			}
			seriesIdx += type.data.length
		}
	}

	// Manual legend placement. Each axis of CT_ManualLayout is independent: omitting xMode/x (or
	// yMode/y, etc.) leaves that axis on automatic layout. x/y use edge mode so they are absolute
	// fractions of the chart; w/h are fractions of the chart size. Schema order: xMode, yMode, x, y, w, h.
	const legendLayout = rel.opts.legendLayout
	const placed = (key: 'x' | 'y' | 'w' | 'h'): boolean => typeof legendLayout?.[key] === 'number'
	let layout = ''
	if (placed('x') || placed('y') || placed('w') || placed('h')) {
		const modes =
			(placed('x') ? voidEl('c:xMode', { val: 'edge' }) : '') + (placed('y') ? voidEl('c:yMode', { val: 'edge' }) : '')
		const vals = (['x', 'y', 'w', 'h'] as const)
			.filter(placed)
			.map((key) => voidEl(`c:${key}`, { val: legendLayout?.[key] }))
			.join('')
		layout = el('c:layout', null, raw(el('c:manualLayout', null, raw(modes + vals))))
	}

	let txPr = ''
	if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
		// No `Number()` here, and none at the ten sibling font-size options either: the
		// option is typed `number`, and a caller from untyped JS who passes a string now
		// gets `coord/non-finite` from the converter rather than a silent coercion.
		const defRPr = el(
			'a:defRPr',
			{ sz: rel.opts.legendFontSize ? ptToHundredths(rel.opts.legendFontSize) : undefined },
			[
				raw(rel.opts.legendColor ? genXmlColorSelection(rel.opts.legendColor) : ''),
				raw(rel.opts.legendFontFace ? createChartTextFonts(rel.opts.legendFontFace) : ''),
			]
		)
		txPr = el('c:txPr', null, [
			raw(voidEl('a:bodyPr', null)),
			raw(voidEl('a:lstStyle', null)),
			raw(el('a:p', null, [raw(el('a:pPr', null, raw(defRPr))), raw(voidEl('a:endParaRPr', { lang: 'en-US' }))])),
		])
	}

	return el('c:legend', null, [
		raw(voidEl('c:legendPos', { val: rel.opts.legendPos })),
		raw(entries),
		raw(layout),
		raw(voidEl('c:overlay', { val: 0 })),
		raw(txPr),
	])
}

/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {SlideRelChart} rel - chart object
 * @return {string} XML
 */
export function makeXmlCharts(rel: SlideRelChart): string {
	// `chartArea`/`plotArea` are always populated by addChartDefinition() but stay optional on the type.
	const chartArea = rel.opts.chartArea ?? {}
	const plotAreaOpts = rel.opts.plotArea ?? {}
	let usesSecondaryValAxis = false
	let usesSecondaryCatAxis = false
	// Combo charts: a scatter/bubble subchart draws numbers on its category (X)
	// axis, so that axis must be emitted as a `<c:valAx>` rather than a `<c:catAx>`
	// or PowerPoint flags the file for repair. Track, per category axis,
	// the scatter/bubble subchart type that owns it (if any) and whether a
	// category-based subchart also references it (an unsatisfiable conflict).
	let primaryCatAxisValType: ChartType | null = null
	let secondaryCatAxisValType: ChartType | null = null
	let primaryCatAxisHasCategoryChart = false
	let secondaryCatAxisHasCategoryChart = false

	// STEP 1: the plot for every subchart's series and points.
	let plots = ''
	if (Array.isArray(rel.opts._type)) {
		for (const type of rel.opts._type) {
			const options = resolveSubchartOptions(rel.opts, type.options)
			const valAxisId = options.secondaryValAxis ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY
			const catAxisId = options.secondaryCatAxis ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY
			usesSecondaryValAxis = usesSecondaryValAxis || (options.secondaryValAxis ?? false)
			usesSecondaryCatAxis = usesSecondaryCatAxis || (options.secondaryCatAxis ?? false)
			const subType = asChartType(type.type)
			// Record whether this subchart needs a value-based X axis (scatter/bubble)
			// or a category-based X axis, keyed to the primary/secondary cat axis it uses.
			const usesValueXAxis = isXyChart(subType)
			if (options.secondaryCatAxis) {
				if (usesValueXAxis) secondaryCatAxisValType = subType
				else secondaryCatAxisHasCategoryChart = true
			} else {
				if (usesValueXAxis) primaryCatAxisValType = subType
				else primaryCatAxisHasCategoryChart = true
			}
			plots += makeChartType(subType, type.data, options, valAxisId, catAxisId)
		}
	} else if (rel.opts._type) {
		plots = makeChartType(rel.opts._type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY)
	}

	// STEP 2: the axes the plots just referenced (the flags above decide how many).
	const axes = makeChartAxesXml(
		rel,
		usesSecondaryValAxis,
		usesSecondaryCatAxis,
		primaryCatAxisValType,
		secondaryCatAxisValType,
		primaryCatAxisHasCategoryChart,
		secondaryCatAxisHasCategoryChart
	)

	// STEP 3: the plot area, then the chart it sits in.
	const plotArea = el('c:plotArea', null, [
		raw(makePlotAreaLayoutXml(rel)),
		raw(plots),
		raw(axes),
		// NOTE: the data table goes between `</c:valAx>` and `<c:spPr>`.
		rel.opts.showDataTable ? raw(makeDataTableXml(rel)) : null,
		raw(chartShapeProps(plotAreaOpts.fill, plotAreaOpts.border)),
	])
	const chart = el('c:chart', null, [
		raw(makeChartHeaderXml(rel)),
		raw(plotArea),
		rel.opts.showLegend ? raw(makeLegendXml(rel)) : null,
		raw(voidEl('c:plotVisOnly', { val: 1 })),
		raw(voidEl('c:dispBlanksAs', { val: rel.opts.displayBlanksAs })),
		isScatterChart(rel.opts._type) ? raw(voidEl('c:showDLblsOverMax', { val: 1 })) : null,
	])

	// STEP 4: the chart space around it — shape props, the embedded workbook relationship, and
	// the metadata extension (CT_ChartSpace order: externalData → printSettings → userShapes → extLst).
	return (
		XML_DECL +
		el('c:chartSpace', CHART_SPACE_NS, [
			// PowerPoint defaults to 1904 dates, Excel to 1900.
			raw(voidEl('c:date1904', { val: 0 })),
			raw(voidEl('c:roundedCorners', { val: chartArea.roundedCorners ? 1 : 0 })),
			raw(chart),
			raw(chartShapeProps(chartArea.fill, chartArea.border)),
			raw(el('c:externalData', { 'r:id': 'rId1' }, raw(voidEl('c:autoUpdate', { val: 0 })))),
			raw(genXmlChartMetadata(rel.opts.metadata)),
		])
	)
}

/**
 * Stable ts-pptx vendor GUID identifying the chart-metadata extension on `c:chartSpace/c:extLst`.
 * Custom data rides under this URI in a foreign namespace so PowerPoint preserves it (the extLst
 * mechanism) instead of stripping/repairing it as it would an unrecognised sibling element.
 */
const CHART_METADATA_EXT_URI = '{094A432E-1F6C-499B-95B8-B57DC9536949}'

/** Foreign namespace for the chart-metadata extension payload. */
const CHART_METADATA_NS = 'http://ts-pptx.com/schema/chart/metadata'

/**
 * Create the chart-space extension-list XML carrying custom `metadata` key/value annotations.
 * Emits nothing when metadata is absent or contains no valid entries. Keys must be non-empty
 * strings and values must be strings; invalid entries are dropped with a warning rather than
 * emitting degenerate XML (per the "no silent coercion" policy).
 * @param {Record<string, string>} [metadata] custom chart-level metadata
 * @return {string} `<c:extLst>…</c:extLst>` XML, or '' when there is nothing valid to emit
 */
function genXmlChartMetadata(metadata?: Record<string, string>): string {
	if (metadata == null) return ''
	if (typeof metadata !== 'object' || Array.isArray(metadata)) {
		warn('chart/invalid-metadata', 'chart `metadata` must be a plain object of string key/value pairs; ignored.')
		return ''
	}

	let items = ''
	for (const [key, value] of Object.entries(metadata)) {
		if (typeof key !== 'string' || key.length === 0) {
			warn(
				'chart/invalid-metadata-key',
				`chart metadata key "${String(key)}" is not a non-empty string; entry skipped.`
			)
			continue
		}
		if (typeof value !== 'string') {
			warn('chart/invalid-metadata-value', `chart metadata value for key "${key}" is not a string; entry skipped.`)
			continue
		}
		items += voidEl('pgm:item', { key, value })
	}
	if (items === '') return ''

	return el(
		'c:extLst',
		null,
		raw(
			el(
				'c:ext',
				{ uri: CHART_METADATA_EXT_URI },
				raw(el('pgm:metadata', { 'xmlns:pgm': CHART_METADATA_NS }, raw(items)))
			)
		)
	)
}

/**
 * Create XML string for any given chart type.
 * Dispatches to a per-family plot helper; resolves the effective value number format first.
 * @param {ChartType} chartType chart type name
 * @param {OptsChartDataInternal[]} data chart data
 * @param {ChartOptsInternal} opts chart options
 * @param {string} valAxisId chart val axis id
 * @param {string} catAxisId chart cat axis id
 * @example 'bubble' returns <c:bubbleChart></c>
 * @example '<c:lineChart>'
 * @return {string} XML chart
 */

function makeChartType(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string
): string {
	// NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
	// ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
	//
	// PowerPoint and Google Slides render values using the cached *source* number format carried in
	// each series' `<c:numCache><c:formatCode>` (mirroring the embedded workbook cell format), NOT the
	// `<c:dLbls><c:numFmt>` mask — so when the value cache is left as "General" those engines display
	// raw values (e.g. `0.1` instead of `10%`) even though LibreOffice honors the dLbls mask. Resolve a
	// single effective value format and stamp it onto every value cache below so all three engines agree.
	// Precedence keeps the historical `valLabelFormatCode` winner,
	// then the data-table format, and finally falls back to `dataLabelFormatCode` (the most common knob).
	const valFmtCode = opts.valLabelFormatCode || opts.dataTableFormatCode || opts.dataLabelFormatCode || 'General'

	switch (chartType) {
		case ChartType.area:
		case ChartType.bar:
		case ChartType.bar3d:
		case ChartType.line:
		case ChartType.radar:
			return makeCatAxisPlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.scatter:
			return makeScatterPlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.bubble:
		case ChartType.bubble3d:
			return makeBubblePlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.stock:
			return makeStockPlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.surface:
			return makeSurfacePlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.doughnut:
		case ChartType.pie:
			return makePiePlot(chartType, data, opts, valFmtCode)
		default:
			// The unmatched members are exactly the chartEx catalog (`CHARTEX_TYPES` in `enums.ts`),
			// which `chartExLayoutId` in `./chartex-xml` owns; callers pick between the two builders
			// with `isChartExType`, so arriving here means the chart was routed to the wrong emitter.
			// Returning `''` here used to hide that: it emitted a `<c:plotArea>` with axes and no plot
			// at all, i.e. a chart-shaped hole PowerPoint opens and shows empty.
			throw new InternalError(
				'chart/type-not-routed',
				`makeChartType: "${String(chartType)}" has no <c:...Chart> plot — the chartEx catalog is built by makeXmlChartEx, and a newly added classic type needs an arm here`,
				{ detail: { chartType } }
			)
	}
}
