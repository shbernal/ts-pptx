/**
 * Gate deck: the chart-emitter matrix.
 *
 * Corpus for `scripts/byte-identity.mjs`, built for the `[chart-el-migration]` work —
 * moving `src/gen/chart/` off hand-concatenated XML and onto `src/gen/oxml/el.ts`. That
 * migration is gated on byte-identity, and the showcase corpus reaches three chart types
 * (bar, doughnut, line) out of the nine emitter files. Seven of them had no evidence at
 * all, which makes a green gate on those files a statement about nothing. See
 * `./README.md` for why a gate deck is a separate thing from a showcase.
 *
 * One chart per slide, one case per entry in `CASES`. Every case is reached on every run —
 * there is no sampling and no randomness — so a diff names the case that moved.
 *
 * What each group is here to reach:
 *
 *   plot-cat-axis   bar/bar3D in both directions and all three groupings, line with every
 *                   marker and dash knob, area, radar. The single densest emitter.
 *   plot-pie        pie and doughnut, label permutations, hole size, first-slice angle.
 *   plot-scatter    all three `dataLabelFormatScatter` shapes — they emit structurally
 *                   different `<c:dLbls>`, so one case cannot stand for the others.
 *   plot-bubble     bubble and bubble3D, with and without the size data label.
 *   plot-stock      all four `stockStyle`s. `vhlc`/`vohlc` are not cosmetic variants: they
 *                   put the price series on a SECONDARY axis pair and lead with a volume
 *                   `<c:barChart>`, so they are the only path to the four-axis layout.
 *   plot-surface    3-D surface, 2-D contour, and wireframe — three branches, all cheap.
 *   chart-axes      the axis knobs the plots never touch on their own: titles, gridlines,
 *                   min/max, log scale, inverted orientation, tick marks, label rotation,
 *                   display units, crossing, and the `dateAx` branch that `catLabelFormatCode`
 *                   selects.
 *   chart-xml       title, legend, data table, 3-D view, explicit layout, area fills, and a
 *                   combo chart with secondary axes.
 *   embed-xlsx      multi-level categories, blank labels, and long strings — the workbook is
 *                   written from the same data every chart above already carries.
 *
 * ESCAPING. Two cases carry `&`, `<`, `>` and quotes in series names and category labels.
 * Those strings flow through `el()`-built blocks today (`strRefBlock`, `catRefBlock`) and
 * through the workbook's shared-string table, so they are valid now and must stay
 * byte-identical through the migration. They are deliberately NOT put anywhere that is
 * still interpolated raw — a format code, say — because today's emitter would write invalid
 * XML there and the baseline would freeze a broken part as the reference. That gap is a
 * real bug and gets a unit test during the migration of the file that owns it, not a
 * fixture here.
 */
import TsPptx, { ChartType } from '../../dist/node.js'

const LABELS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
/** One series over the shared weekday labels. @param {string} name @param {number[]} values */
const S = (name, values) => ({ name, labels: LABELS, values })

const TWO = [S('North', [12, 18, 9, 22, 15]), S('South', [7, 14, 19, 11, 20])]
const THREE = [...TWO, S('East', [4, 9, 13, 6, 17])]

// Adversarial text. Valid through today's emitters; see the ESCAPING note above.
const SPICY = [
	S('R&D "core" <alpha>', [12, 18, 9, 22, 15]),
	{ name: 'Ops & Support', labels: ['A&B', '<x>', '"q"', "o'n", 'p>q'], values: [7, 14, 19, 11, 20] },
]

const ML_LABELS = [
	['Gear', 'Berg', 'Cam', 'Gear', 'Berg', 'Cam'],
	['Mech', '', '', 'Elec', '', ''],
]
const MULTILEVEL = [
	{ name: 'West', labels: ML_LABELS, values: [11, 8, 3, 0, 11, 3] },
	{ name: 'East', labels: ML_LABELS, values: [0, 3, 2, 0, 4, 1] },
]

const XY_PLAIN = [
	{ name: 'X-Axis', values: [1, 2, 3, 4] },
	{ name: 'Y1', values: [13, 20, 21, 25] },
]
// Scatter point labels live in `labels`, one group of per-point strings. The blank is
// load-bearing: 'customXY' appends XVALUE/YVALUE fields only for non-blank labels.
const XY_LABELLED = [
	{ name: 'X-Axis', values: [1, 2, 3, 4] },
	{ name: 'Y1', values: [13, 20, 21, 25], labels: [['Alpha', 'Beta', '', 'Delta']] },
]
const BUBBLE = [
	{ name: 'X-Axis', values: [1, 2, 3, 4] },
	{ name: 'Y1', values: [13, 20, 21, 25], sizes: [10, 5, 20, 15] },
]

const HIGH = S('High', [55, 57, 57, 58, 58])
const LOW = S('Low', [11, 12, 13, 11, 35])
const CLOSE = S('Close', [32, 35, 34, 35, 43])
const OPEN = S('Open', [20, 33, 30, 33, 37])
const VOL = S('Volume', [1200, 1500, 900, 1700, 1400])
const STOCK = {
	hlc: [HIGH, LOW, CLOSE],
	ohlc: [OPEN, HIGH, LOW, CLOSE],
	vhlc: [VOL, HIGH, LOW, CLOSE],
	vohlc: [VOL, OPEN, HIGH, LOW, CLOSE],
}

const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5']

/** Every case, reached on every run. @type {{name: string, data: any, opts: any}[]} */
const CASES = [
	// ---------------------------------------------------------------- plot-cat-axis
	{
		name: 'bar/col clustered',
		data: THREE,
		opts: { type: ChartType.bar, barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 65, barOverlapPct: -20 },
	},
	{
		name: 'bar/col stacked',
		data: THREE,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			barGrouping: 'stacked',
			showValue: true,
			dataLabelPosition: 'ctr',
			dataLabelFormatCode: '$0.0',
		},
	},
	{
		name: 'bar/col percentStacked',
		data: THREE,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			barGrouping: 'percentStacked',
			chartColors: PALETTE,
			chartColorsOpacity: 70,
		},
	},
	{
		name: 'bar/bar horizontal',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'bar',
			dataBorder: { width: 1, color: '333333' },
			shadow: { type: 'outer', blur: 3, offset: 2, angle: 45, color: '000000', opacity: 0.4 },
		},
	},
	{
		name: 'bar3D',
		data: THREE,
		opts: {
			type: ChartType.bar3d,
			barDir: 'col',
			bar3DShape: 'cylinder',
			barGapDepthPct: 120,
			v3DRotX: 25,
			v3DRotY: 30,
			v3DRAngAx: true,
			v3DPerspective: 40,
		},
	},
	{
		name: 'bar inverted colors',
		data: TWO,
		opts: { type: ChartType.bar, barDir: 'col', invertedColors: ['C00000', '00B050'] },
	},
	{
		name: 'line markers + dash',
		data: THREE,
		opts: {
			type: ChartType.line,
			lineSize: 3,
			lineDash: 'dash',
			lineCap: 'round',
			lineDataSymbol: 'diamond',
			lineDataSymbolSize: 8,
			lineDataSymbolLineColor: '404040',
			lineDataSymbolLineSize: 1,
			lineSmooth: true,
			showSerName: true,
		},
	},
	{
		name: 'line per-series dash',
		data: THREE,
		opts: { type: ChartType.line, lineDashValues: ['solid', 'dash', 'sysDot'], showLeaderLines: true },
	},
	{ name: 'line no markers', data: TWO, opts: { type: ChartType.line, lineDataSymbol: 'none', lineSize: 0 } },
	{ name: 'area stacked', data: TWO, opts: { type: ChartType.area, barGrouping: 'stacked' } },
	{ name: 'radar marker', data: THREE, opts: { type: ChartType.radar, radarStyle: 'markers' } },
	{ name: 'radar filled', data: TWO, opts: { type: ChartType.radar, radarStyle: 'filled' } },
	{
		name: 'per-series colors',
		data: THREE,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			seriesOptions: [{ color: 'C00000' }, { color: '00B050' }, { color: '0070C0' }],
		},
	},
	{
		name: 'label backgrounds + fonts',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			showValue: true,
			dataLabelBkgrdColors: true,
			dataLabelColor: 'FFFFFF',
			dataLabelFontFace: 'Georgia',
			dataLabelFontSize: 11,
			dataLabelFontBold: true,
			dataLabelFontItalic: true,
		},
	},

	// ---------------------------------------------------------------- plot-pie
	{
		name: 'pie percent',
		data: [TWO[0]],
		opts: { type: ChartType.pie, showPercent: true, showLegend: true, legendPos: 'r' },
	},
	{
		name: 'pie value + leader lines',
		data: [TWO[0]],
		opts: {
			type: ChartType.pie,
			showValue: true,
			showLabel: true,
			showLeaderLines: true,
			firstSliceAng: 45,
			dataLabelPosition: 'outEnd',
		},
	},
	{
		name: 'doughnut',
		data: [TWO[1]],
		opts: { type: ChartType.doughnut, holeSize: 62, showPercent: true, dataLabelColor: 'FFFFFF' },
	},
	{ name: 'doughnut no labels', data: [TWO[1]], opts: { type: ChartType.doughnut, holeSize: 30, showLegend: false } },

	// ---------------------------------------------------------------- plot-scatter
	{ name: 'scatter plain', data: XY_PLAIN, opts: { type: ChartType.scatter, lineSize: 0, lineDataSymbol: 'circle' } },
	{
		name: 'scatter labels custom',
		data: XY_LABELLED,
		opts: { type: ChartType.scatter, showLabel: true, dataLabelFormatScatter: 'custom' },
	},
	{
		name: 'scatter labels customXY',
		data: XY_LABELLED,
		opts: { type: ChartType.scatter, showLabel: true, dataLabelFormatScatter: 'customXY' },
	},
	{
		name: 'scatter labels XY',
		data: XY_LABELLED,
		opts: { type: ChartType.scatter, showLabel: true, dataLabelFormatScatter: 'XY' },
	},
	{
		name: 'scatter smooth + format codes',
		data: XY_PLAIN,
		opts: {
			type: ChartType.scatter,
			lineSmooth: true,
			lineSize: 2,
			catAxisLabelFormatCode: '0.0',
			valAxisLabelFormatCode: '#,##0',
		},
	},

	// ---------------------------------------------------------------- plot-bubble
	{ name: 'bubble', data: BUBBLE, opts: { type: ChartType.bubble } },
	{ name: 'bubble size label', data: BUBBLE, opts: { type: ChartType.bubble, showBubbleSize: true, showValue: true } },
	{ name: 'bubble3D', data: BUBBLE, opts: { type: ChartType.bubble3d, showBubbleSize: true } },

	// ---------------------------------------------------------------- plot-stock
	{ name: 'stock hlc', data: STOCK.hlc, opts: { type: ChartType.stock, stockStyle: 'hlc' } },
	{ name: 'stock ohlc', data: STOCK.ohlc, opts: { type: ChartType.stock, stockStyle: 'ohlc' } },
	{ name: 'stock vhlc (4 axes)', data: STOCK.vhlc, opts: { type: ChartType.stock, stockStyle: 'vhlc' } },
	{
		name: 'stock vohlc (4 axes)',
		data: STOCK.vohlc,
		opts: { type: ChartType.stock, stockStyle: 'vohlc', showLegend: true, legendPos: 'b' },
	},

	// ---------------------------------------------------------------- plot-surface
	{ name: 'surface 3D', data: THREE, opts: { type: ChartType.surface } },
	{
		name: 'surface 3D wireframe',
		data: THREE,
		opts: { type: ChartType.surface, surfaceWireframe: true, v3DRotX: 20, v3DRotY: 40 },
	},
	{ name: 'surface 2D contour', data: THREE, opts: { type: ChartType.surface, surface3D: false } },
	{
		name: 'surface 2D contour wireframe',
		data: THREE,
		opts: { type: ChartType.surface, surface3D: false, surfaceWireframe: true },
	},

	// ---------------------------------------------------------------- chart-axes
	{
		name: 'axis titles all three',
		data: THREE,
		opts: {
			type: ChartType.bar3d,
			barDir: 'col',
			showCatAxisTitle: true,
			catAxisTitle: 'Category',
			catAxisTitleColor: '404040',
			catAxisTitleFontFace: 'Georgia',
			catAxisTitleFontSize: 12,
			catAxisTitleRotate: 0,
			showValAxisTitle: true,
			valAxisTitle: 'Value',
			valAxisTitleRotate: 270,
			showSerAxisTitle: true,
			serAxisTitle: 'Series',
		},
	},
	{
		name: 'axis gridlines + lines',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			catGridLine: { color: 'D9D9D9', size: 1, style: 'dash' },
			valGridLine: { color: 'BFBFBF', size: 2, style: 'solid' },
			catAxisLineShow: true,
			catAxisLineColor: '808080',
			catAxisLineSize: 2,
			catAxisLineStyle: 'sysDash',
			valAxisLineShow: false,
		},
	},
	{
		name: 'axis min/max/units/log',
		data: TWO,
		opts: {
			type: ChartType.line,
			valAxisMinVal: 0,
			valAxisMaxVal: 100,
			valAxisMajorUnit: 25,
			valAxisLogScaleBase: 10,
			valAxisDisplayUnit: 'thousands',
			valAxisDisplayUnitLabel: true,
			valAxisCrossBetween: 'midCat',
			valAxisCrossesAt: 0,
		},
	},
	{
		name: 'axis inverted + hidden',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			catAxisOrientation: 'maxMin',
			valAxisOrientation: 'maxMin',
			serAxisOrientation: 'maxMin',
			valAxisHidden: true,
		},
	},
	{
		name: 'axis cat hidden + tick marks',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			catAxisHidden: true,
			catAxisMajorTickMark: 'cross',
			catAxisMinorTickMark: 'in',
			valAxisMajorTickMark: 'out',
			valAxisMinorTickMark: 'cross',
			catAxisLabelPos: 'low',
			valAxisLabelPos: 'high',
		},
	},
	{
		name: 'axis label rotation + fonts',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			catAxisLabelRotate: 45,
			valAxisLabelRotate: 315,
			catAxisLabelFontFace: 'Georgia',
			catAxisLabelFontSize: 9,
			catAxisLabelFontBold: true,
			catAxisLabelFontItalic: true,
			catAxisLabelColor: '404040',
			valAxisLabelFontFace: 'Consolas',
			valAxisLabelColor: '808080',
			catAxisLabelFrequency: 2,
		},
	},
	{
		name: 'axis date (dateAx branch)',
		data: [
			{
				name: 'Daily',
				labels: ['2026-01-01', '2026-01-02', '2026-01-03', '2026-01-04', '2026-01-05'],
				values: [3, 6, 4, 9, 7],
			},
		],
		opts: {
			type: ChartType.line,
			catLabelFormatCode: 'yyyy-mm-dd',
			catAxisBaseTimeUnit: 'days',
			catAxisMajorTimeUnit: 'days',
			catAxisMinorTimeUnit: 'days',
			catAxisMajorUnit: 1,
			catAxisMinorUnit: 1,
		},
	},
	{
		name: 'axis multi-level labels',
		data: MULTILEVEL,
		opts: { type: ChartType.bar, barDir: 'col', catAxisMultiLevelLabels: true },
	},
	{ name: 'axis crossesAt', data: TWO, opts: { type: ChartType.line, catAxisCrossesAt: 2, valAxisCrossesAt: 5 } },
	{
		name: 'axis serAxis knobs',
		data: THREE,
		opts: {
			type: ChartType.bar3d,
			barDir: 'col',
			serAxisLabelColor: '404040',
			serAxisLabelFontFace: 'Georgia',
			serAxisLabelFontSize: 9,
			serAxisLabelFontBold: true,
			serAxisLabelFontItalic: true,
			serAxisLabelFrequency: 1,
			serAxisLabelPos: 'nextTo',
			serAxisLineShow: true,
			serAxisLineColor: '999999',
			serGridLine: { color: 'E0E0E0', size: 1, style: 'solid' },
			serLabelFormatCode: 'General',
		},
	},
	{ name: 'axis serAxis hidden', data: THREE, opts: { type: ChartType.bar3d, barDir: 'col', serAxisHidden: true } },

	// ---------------------------------------------------------------- chart-xml
	{
		name: 'title + legend right',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			showTitle: true,
			title: 'Revenue & "Growth" <FY26>',
			titleAlign: 'left',
			titleColor: '203864',
			titleFontFace: 'Georgia',
			titleFontSize: 16,
			titleBold: true,
			titleItalic: true,
			titleUnderline: true,
			titleRotate: 0,
			titlePos: { x: 0.1, y: 0.05 },
			showLegend: true,
			legendPos: 'r',
			legendColor: '595959',
			legendFontFace: 'Consolas',
			legendFontSize: 10,
		},
	},
	{
		name: 'legend top-right + layout',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			showLegend: true,
			legendPos: 'tr',
			legendLayout: { x: 0.7, y: 0.1, w: 0.25, h: 0.3 },
		},
	},
	{
		name: 'data table',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			showDataTable: true,
			showDataTableKeys: true,
			showDataTableHorzBorder: true,
			showDataTableVertBorder: true,
			showDataTableOutline: true,
			dataTableFontSize: 9,
			dataTableFormatCode: '#,##0',
		},
	},
	{
		name: 'plot + chart area fills',
		data: TWO,
		opts: {
			type: ChartType.bar,
			barDir: 'col',
			chartArea: { fill: { color: 'F2F2F2' }, border: { width: 1, color: '404040' }, roundedCorners: false },
			plotArea: { fill: { color: 'FFFFFF' }, border: { width: 1, color: 'BFBFBF' } },
			layout: { x: 0.1, y: 0.1, w: 0.8, h: 0.75 },
		},
	},
	{
		name: 'displayBlanksAs + format codes',
		data: TWO,
		opts: { type: ChartType.line, displayBlanksAs: 'gap', valLabelFormatCode: '#,##0.00', dataLabelFormatCode: '0.0%' },
	},
	{
		name: 'combo bar + line secondary',
		data: [
			{ type: ChartType.bar, data: TWO, options: { barDir: 'col' } },
			{
				type: ChartType.line,
				data: [S('Index', [3, 5, 4, 7, 6])],
				options: { secondaryValAxis: true, secondaryCatAxis: true, showValAxisTitle: true, valAxisTitle: 'Index' },
			},
		],
		opts: { showLegend: true, legendPos: 'b' },
	},

	// ---------------------------------------------------------------- escaping + embed-xlsx
	{
		name: 'escaping in names and labels',
		data: SPICY,
		opts: { type: ChartType.bar, barDir: 'col', showLegend: true, showValue: true },
	},
	{
		name: 'escaping through pie labels',
		data: [SPICY[1]],
		opts: { type: ChartType.pie, showLabel: true, showPercent: true },
	},
	{
		name: 'blank + long categories',
		data: [
			{
				name: 'Sparse',
				labels: ['', 'Two', '', 'Four', 'A rather long category label that wraps'],
				values: [1, 0, 3, 0, 5],
			},
		],
		opts: { type: ChartType.bar, barDir: 'col' },
	},
	{
		// The category builder's third tag: scatter/bubble put numbers on X, so the axis is a
		// `<c:valAx>`, which is the one arm besides `<c:dateAx>` with a slot for the numeric units.
		name: 'axis scatter X units (valAx branch)',
		data: XY_PLAIN,
		opts: { type: ChartType.scatter, lineSize: 0, catAxisMajorUnit: 2, catAxisMinorUnit: 1 },
	},
]

async function compose() {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_WIDE'
	pptx.author = 'ts-pptx byte-identity gate'
	pptx.title = 'Chart emitter matrix'

	for (const kase of CASES) {
		const slide = pptx.addSlide()
		slide.addText(kase.name, { x: 0.3, y: 0.15, w: 12.7, h: 0.4, fontSize: 14, bold: true })
		slide.addChart(kase.data, { x: 0.3, y: 0.7, w: 12.7, h: 6.2, ...kase.opts })
	}
	return pptx
}

/** @param {string} outFile @returns {Promise<string>} */
export async function build(outFile) {
	const pptx = await compose()
	return await pptx.writeFile({ fileName: outFile })
}

export const gateDeck = {
	slug: 'chart-matrix',
	title: 'Chart emitter matrix',
	description: 'Byte-identity corpus for src/gen/chart/ — every plot type and axis branch, one per slide.',
	fileName: 'gate_chart_matrix.pptx',
	build,
}
