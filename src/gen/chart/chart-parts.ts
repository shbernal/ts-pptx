/**
 * ts-pptx: Shared Chart Fragment Builders
 *
 * The leaf builders every chart region reuses -- titles, gridlines, series data points,
 * error bars, number caches, leader lines, borders. Each is a pure string builder with
 * no dependency on any other module in this directory, which is what lets the plot
 * builders ({@link ./plot-cat-axis}, {@link ./plot-scatter}, {@link ./plot-bubble},
 * {@link ./plot-pie}), the axes ({@link ./chart-axes}), and the chart envelope
 * ({@link ./chart-xml}) all draw on them without a cycle.
 */

import { ChartType } from '../../enums.js'
import {
	BARCHART_COLORS,
	DEF_CHART_GRIDLINE,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SHAPE_SHADOW,
	PIECHART_COLORS,
} from '../../constants-internal.js'
import type {
	BorderProps,
	Coord,
	ChartErrorBarOptions,
	ChartPropsTitle,
	ChartSeriesOpts,
	OptsChartGridLine,
} from '../../types/index.js'
import type { ChartOptsInternal, MaybeUndefined, OptsChartDataInternal } from '../../types/internal.js'
import { warn } from '../../diagnostics.js'
import { alphaEl, createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection, genXmlPatternFill } from '../drawingml/fill.js'
import { clampFontSizeSz } from '../drawingml/clamp.js'
import { borderLine, createLineCap, noStrokeLine, resolveDash, strokeDash, strokePaint } from '../drawingml/line.js'
import { gridLineStroke, gridLineSuppressed } from './chart-stroke.js'
import { convertAngleUnits, percentToFixedPercent, ptsToEmuLenient } from '../../units-internal.js'
import { coordToEmu, EMU_PER_INCH } from '../../units.js'
import { dataValues, type SheetLayout } from './data-refs.js'
import { el, raw, voidEl, type XmlChild } from '../oxml/el.js'
import { type XsdBool, xsdBool } from '../../ooxml/xsd-boolean.js'

/**
 * The six `c:show*` flags every `<c:dLbls>` carries, in schema order, as `XmlChild`s for the
 * parent's child list. Eight `<c:dLbls>` builders across five modules wrote all six out by
 * hand; each cared about one or two of them and hard-coded the rest to `0`.
 *
 * The order is `CT_DLbls`'s and is not negotiable, which is the other reason to state it once:
 * a flag inserted in the wrong place is a repair prompt rather than a wrong-looking chart.
 * Every flag defaults to `0`, so a caller names only what it means to turn on.
 *
 * @param flags - the values to emit, defaulting to `0`
 */
export function dLblShowFlags(flags: {
	legendKey?: XsdBool
	val?: XsdBool
	catName?: XsdBool
	serName?: XsdBool
	percent?: XsdBool
	bubbleSize?: XsdBool
}): XmlChild[] {
	return (
		[
			['c:showLegendKey', flags.legendKey],
			['c:showVal', flags.val],
			['c:showCatName', flags.catName],
			['c:showSerName', flags.serName],
			['c:showPercent', flags.percent],
			['c:showBubbleSize', flags.bubbleSize],
		] as const
	).map(([name, val]) => raw(voidEl(name, { val: val ?? 0 })))
}

/**
 * The data-label run properties, in the attribute order the scatter and custom-label paths
 * emit: `sz, b, i, u, strike`.
 *
 * There is a second ordering in this file — `b, i, strike, sz, u`, used by
 * {@link chartDataLabels} and by the bubble plot — and the two are **deliberately not
 * merged**. Attribute order carries no meaning in XML, so either is correct and a reader is
 * right to want one; but merging them moves bytes in a refactor whose whole claim is that it
 * does not, and the byte-identity gate is what makes that claim checkable. Unifying the
 * spelling is a decision to take on its own, with its own re-baseline.
 *
 * `over` layers a series' own `seriesOptions` entry on top with `??`, leaving the chart-level
 * reading below it exactly as it was — a series that states nothing produces the same bytes.
 * @param opts - the chart's normalized options
 * @param over - this series' `seriesOptions` entry, whose stated fields win over the chart's
 */
export function labelFontAttrs(opts: ChartOptsInternal, over?: ChartSeriesOpts): Record<string, string | number> {
	return {
		sz: clampFontSizeSz(over?.dataLabelFontSize ?? opts.dataLabelFontSize ?? DEF_FONT_SIZE, 'dataLabelFontSize'),
		b: xsdBool(over?.dataLabelFontBold ?? opts.dataLabelFontBold),
		i: xsdBool(over?.dataLabelFontItalic ?? opts.dataLabelFontItalic),
		u: 'none',
		strike: 'noStrike',
	}
}

/**
 * The colour and typeface children of those run properties.
 *
 * The fill goes through {@link genXmlColorSelection} rather than a hand-built `<a:solidFill>`
 * around {@link createColorElement} so that both data-label builders answer an empty
 * `dataLabelColor` the same way — that helper's contract is that a value naming nothing to
 * paint emits nothing. For any colour that names something the two spellings are the same
 * bytes.
 * @param opts - the chart's normalized options
 * @param over - this series' `seriesOptions` entry, whose stated fields win over the chart's
 */
export function labelFontChildren(opts: ChartOptsInternal, over?: ChartSeriesOpts): XmlChild[] {
	return [
		raw(genXmlColorSelection(over?.dataLabelColor ?? opts.dataLabelColor ?? DEF_FONT_COLOR)),
		raw(createChartTextFonts(over?.dataLabelFontFace ?? opts.dataLabelFontFace ?? 'Arial')),
	]
}

/**
 * The `<c:txPr>` a data-label block wraps its run properties in: an empty `<a:bodyPr>`, an empty
 * `<a:lstStyle>`, then one `<a:p>` carrying the `<a:pPr>` the caller built.
 *
 * Five sites spelled this out byte-identically — the shared `<c:dLbls>` builder, the bubble,
 * category-axis and pie plots, and the pie's plot-level block. The axis one in `chart-axes.ts`
 * is deliberately NOT this: it carries a rotation on the `bodyPr` and an `<a:endParaRPr>` after
 * the `<a:pPr>`.
 * @param defRPr - the already-built `<a:defRPr>`
 */
export function labelTextProps(defRPr: string): string {
	return el('c:txPr', null, [
		raw(voidEl('a:bodyPr', null)),
		raw(voidEl('a:lstStyle', null)),
		raw(el('a:p', null, raw(el('a:pPr', null, raw(defRPr))))),
	])
}

/**
 * The `<a:defRPr>` a `<c:dLbls>` text style carries, in the `b, i, strike, sz, u` ordering the
 * chart-level and bubble label blocks share. See {@link labelFontAttrs} for why there are two
 * orderings and why they stay apart.
 *
 * This and {@link labelFontAttrs} read the same five options and read them the same way:
 * `over?.X ?? opts.X ?? DEFAULT`. Attribute order is the only difference left between them.
 *
 * They disagreed on the operator once — `??` here against `||` there — and which one a chart
 * reached was decided by its type and its label format, so `dataLabelFontSize: 0` had three
 * answers: a bar warned and clamped, a pie dropped the value with no diagnostic at all, and an
 * `XY` scatter, which reaches both builders, emitted `sz="100"` on one label block and
 * `sz="1200"` on another of the same chart. `??` is the reading kept everywhere, because the
 * caller stating a value and the caller saying nothing are different. An explicit `0` therefore
 * goes through {@link clampFontSizeSz} and is corrected to the `ST_TextFontSize` minimum with a
 * warning, and **omitting the option is the only spelling of "use the default"**.
 * @param opts - the chart's normalized options
 * @param over - this series' `seriesOptions` entry, whose stated fields win over the chart's
 */
export function dataLabelDefRPr(opts: ChartOptsInternal, over?: ChartSeriesOpts): string {
	return el(
		'a:defRPr',
		{
			b: xsdBool(over?.dataLabelFontBold ?? opts.dataLabelFontBold),
			i: xsdBool(over?.dataLabelFontItalic ?? opts.dataLabelFontItalic),
			strike: 'noStrike',
			sz: clampFontSizeSz(over?.dataLabelFontSize ?? opts.dataLabelFontSize ?? DEF_FONT_SIZE, 'dataLabelFontSize'),
			u: 'none',
		},
		[
			raw(genXmlColorSelection(over?.dataLabelColor ?? opts.dataLabelColor ?? DEF_FONT_COLOR)),
			raw(createChartTextFonts(over?.dataLabelFontFace ?? opts.dataLabelFontFace ?? 'Arial')),
		]
	)
}

/**
 * The signature every axis-plot builder shares: `<c:areaChart>`, `<c:barChart>`,
 * `<c:lineChart>`, `<c:radarChart>`, `<c:scatterChart>`, `<c:bubbleChart>`,
 * `<c:stockChart>` and `<c:surfaceChart>` are all built from the same inputs and
 * dispatched from one `switch` in {@link ./chart-xml}.
 *
 * Naming it makes adding an input one edit instead of six, and makes a builder that
 * disagrees with the dispatch a type error rather than an argument quietly dropped on the
 * floor. `makePiePlot` is deliberately not of this type: a pie has no axes, so it takes
 * neither axis id, and widening the contract to fit it would mean handing it two arguments
 * it must ignore.
 *
 * `chartType` is passed even to the builders that serve a single kind (`stock`, `surface`),
 * which spell it `_chartType`, so the dispatch stays uniform.
 */
export type PlotBuilder = (
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string,
	valFmtCode: string,
	sheet: SheetLayout
) => string

const VALID_CHART_TIME_UNITS = ['days', 'months', 'years']

/**
 * A validated, lowercased axis time unit, or `undefined` when the value is not one.
 *
 * A garbage time unit does not degrade the chart -- it corrupts it, and PowerPoint renders
 * nothing at all -- so an unrecognized value warns and is dropped rather than emitted.
 *
 * This replaces two copies of a loop that wrote `opts[opt] = undefined` back onto the options
 * bag from inside an XML builder. Nothing downstream was corrupted (`ChartOptsInternal` is a
 * copy `gen/define/chart.ts` already made, guarded by
 * `test/regression/chart/chart-input-immutability.test.js`), but it meant calling the same axis
 * builder twice gave a different answer the second time, and it put validation somewhere nobody
 * would look for it.
 *
 * @param value - the caller's time unit, or `undefined`
 * @param optionName - the option as the caller spells it, for the warning
 */
export function validTimeUnit(value: string | undefined, optionName: string): string | undefined {
	if (!value) return undefined
	if (typeof value !== 'string' || !VALID_CHART_TIME_UNITS.includes(value.toLowerCase())) {
		warn('chart/invalid-axis-time-unit', `"${optionName}" must be one of: 'days','months','years' !`)
		return undefined
	}
	return value.toLowerCase()
}

/**
 * A count attribute that `ST_Skip` (an `xsd:unsignedInt` of at least 1) will accept, or
 * `undefined` when the caller's value is not one.
 *
 * `catAxisLabelFrequency` and `serAxisLabelFrequency` were typed as free-form strings and
 * emitted verbatim, so `'every other'` reached `<c:tickLblSkip val="every other"/>` — and the
 * type also rejected the natural `2`. Same shape as {@link validTimeUnit} beside it: warn and
 * emit nothing rather than write a value outside the attribute's own type.
 * @param value - the caller's value, a number or a numeric string
 * @param optionName - option name as the caller spells it, for the warning
 */
export function positiveIntAttr(value: number | string | undefined, optionName: string): number | undefined {
	if (value === undefined || value === null || value === '') return undefined
	const n = Number(value)
	if (!Number.isInteger(n) || n < 1) {
		warn(
			'chart/option-out-of-range',
			`"${optionName}" must be a whole number of at least 1; ignoring ${String(value)}.`
		)
		return undefined
	}
	return n
}

/**
 * Where one axis meets the other: `<c:crossesAt val>` for an explicit position, `<c:crosses val>`
 * for a rule, and `defaultRule` when the caller stated neither.
 *
 * The two axis builders each spelled this decision out, and both opened it with
 * `typeof value === 'number'` — the one numeric guard in the file that `NaN` passes, since every
 * other numeric axis option is guarded by truthiness. `{ valAxisCrossesAt: NaN }` therefore
 * emitted `<c:crossesAt val="NaN"/>`, and `ST_Double` has no such value.
 * @param value - the caller's `*AxisCrossesAt`: a position, a `ST_Crosses` rule, or nothing
 * @param defaultRule - the `ST_Crosses` rule to fall back on
 * @param optionName - option name as the caller spells it, for the warning
 */
export function axisCrossing(value: number | string | undefined, defaultRule: string, optionName: string): string {
	if (typeof value === 'number') {
		if (Number.isFinite(value)) return voidEl('c:crossesAt', { val: value })
		warn('chart/option-out-of-range', `${optionName} must be a finite number; using "${defaultRule}" instead.`)
		return voidEl('c:crosses', { val: defaultRule })
	}
	return voidEl('c:crosses', { val: value || defaultRule })
}

/**
 * The built-in series palette for a chart type.
 *
 * Pie and doughnut colour their *data points* rather than their series, so they take the wider,
 * flatter `PIECHART_COLORS`; everything else takes `BARCHART_COLORS`. A combo chart's `_type` is
 * an array, which is neither, and lands on the bar palette its subcharts already use.
 * @param {ChartOptsInternal['_type']} type - the normalized chart type
 * @returns {string[]} the palette that applies when the caller named none
 */
export function defaultChartPalette(type: ChartOptsInternal['_type']): string[] {
	return type === ChartType.pie || type === ChartType.doughnut ? PIECHART_COLORS : BARCHART_COLORS
}

/**
 * The palette a chart's series and data points draw from: the caller's `chartColors` when it has
 * entries, otherwise {@link defaultChartPalette} for this chart's type.
 *
 * `addChartDefinition` resolves `chartColors` the same way, so for a plain chart this agrees with
 * what normalization already put there. It is not redundant: a combo subchart's options are not
 * put through that pass (`SUBCHART_VALIDATED_KEYS` does not list `chartColors`), so this is the
 * only place a subchart's palette is decided. Six plot builders wrote the lookup out; going
 * through one function is what makes "no colours named" a single decision rather than six.
 */
export function resolveChartPalette(opts: ChartOptsInternal): string[] {
	return opts.chartColors?.length ? opts.chartColors : defaultChartPalette(opts._type)
}

/**
 * The palette colour for series/point `idx`, cycling back to the start once the palette runs out.
 *
 * Every palette lookup in this directory goes through here so that a deck with more series or
 * data points than the palette has entries still emits the same bytes on every build.
 */
export function paletteColor(palette: readonly string[], idx: number, fallback = '000000'): string {
	if (palette.length === 0) return fallback
	return palette[idx % palette.length] ?? fallback
}

// DEF_CHART_GRIDLINE.color is optional on the type but always present on the constant.
export const DEF_GRIDLINE_COLOR: string = DEF_CHART_GRIDLINE.color ?? '888888'

/**
 * The `a:prstDash/@val` one plotted series' stroke takes: this series' own entry in
 * `lineDashValues`, else the chart-wide `lineDash`, else solid.
 *
 * The bubble, category and scatter plot builders each spelled that fallback chain out
 * byte-identically, and none of the three checked the result against `ST_PresetLineDashVal`.
 * @param opts - the chart's normalized options
 * @param serIndex - the series' 0-based index
 * @returns a value legal for `a:prstDash/@val`
 */
export function seriesDash(opts: ChartOptsInternal, serIndex: number): string {
	return resolveDash(opts.lineDashValues?.[serIndex] ?? opts.lineDash, 'solid', 'lineDash')
}

/**
 * Fill fragment for a `chartColors`-derived series/line/marker colour.
 *
 * A `'transparent'` entry means "no fill" — an invisible series, connecting line, or marker
 * stroke — and maps to `<a:noFill/>`. Any real colour goes through the normal solid-fill path.
 * The series and marker *fill* paths already special-case `'transparent'`; without this the
 * *stroke* paths (`<a:ln>` on the series line and marker border) would instead pass the literal
 * `'transparent'` through colour validation, warn "not a valid scheme color or hex RGB", and
 * render as black — leaving a stray black line/border where a transparent series was requested.
 */
export function chartColorLineFill(color: string): string {
	return color === 'transparent' ? voidEl('a:noFill') : genXmlColorSelection(color)
}

/**
 * The series *fill* for one plotted series: `<a:noFill>` for a transparent series, a solid fill
 * carrying `<a:alpha>` when the chart sets `chartColorsOpacity`, and a plain solid fill
 * otherwise.
 *
 * Read against {@link chartColorLineFill}, which is the same decision for a *stroke* and stops
 * one step earlier — a stroke has no opacity option to fold in. Three plot builders (bubble,
 * scatter, and the whole category-axis family) each carried this expression verbatim, down to
 * the diagnostic code and the option name in its message, which is what made the `transparent`
 * arm easy to add to one and forget in another.
 *
 * @param opts - the chart's normalized options, read for `chartColorsOpacity`
 * @param serColor - the series' resolved palette colour, or `'transparent'`
 */
function seriesFill(opts: ChartOptsInternal, serColor: string): string {
	if (serColor === 'transparent') return voidEl('a:noFill')
	if (!opts.chartColorsOpacity) return genXmlColorSelection(serColor)
	const alpha = alphaEl(
		percentToFixedPercent(opts.chartColorsOpacity, 'chart/option-out-of-range', 'chartColorsOpacity')
	)
	return el('a:solidFill', null, raw(createColorElement(serColor, alpha)))
}

/**
 * One plotted series' stroke: a width, this series' dash, the palette colour and a round join —
 * or `<a:noFill/>` when the caller asked for no line at all.
 *
 * Three plot builders had this, differing only in whether a per-series `lineSize` could override
 * the chart's. `lineSize: 0` is the caller's "no outline" and is checked for explicitly, not by
 * truthiness, because it is a stated value rather than an absent one.
 * @param opts - the chart's normalized options
 * @param color - the series' resolved palette colour
 * @param serIndex - the series' index, for `lineDashValues`
 * @param sizePt - the width in points; `undefined` falls back to the chart's, then to 2
 */
export function seriesStroke(opts: ChartOptsInternal, color: string, serIndex: number, sizePt?: number): string {
	const widthPt = sizePt ?? opts.lineSize ?? 2
	if (widthPt === 0) return el('a:ln', null, raw(voidEl('a:noFill')))
	return el('a:ln', { w: ptsToEmuLenient(widthPt), cap: createLineCap(opts.lineCap) }, [
		raw(chartColorLineFill(color)),
		raw(voidEl('a:prstDash', { val: seriesDash(opts, serIndex) })),
		raw(voidEl('a:round')),
	])
}

/**
 * The `<c:spPr>` every plotted series carries: its fill, the outline the caller built, and the
 * chart's shadow. Only the outline differs between plot families -- a line-like series strokes
 * itself, a scatter series always does, a bubble series lets `dataBorder` win -- so the caller
 * builds `<a:ln>` and hands it over rather than describing it through flags.
 *
 * @param opts - the chart's normalized options, read for `chartColorsOpacity` and `shadow`
 * @param serColor - the series' resolved palette colour, or `'transparent'`
 * @param line - the already-built `<a:ln>`, or `''` for a series with no outline
 */
export function seriesShapeProps(opts: ChartOptsInternal, serColor: string, line: string): string {
	const fill = seriesFill(opts, serColor)
	return el('c:spPr', null, [raw(fill), raw(line), raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW))])
}

/**
 * The point marker for a series drawn as a line: the symbol, an optional size, and the marker's
 * own fill and outline. `<c:marker>` must precede `<c:dLbls>` in CT_LineSer and CT_ScatterSer
 * (schema order: spPr -> marker -> dPt -> dLbls).
 *
 * @param opts - the chart's normalized options, read for the `lineDataSymbol*` family
 * @param markerColor - the marker's fill colour, or `'transparent'`
 * @param seriesColor - what the marker's outline falls back to when `lineDataSymbolLineColor` is
 *   unset; a scatter series passes its own colour for both roles
 */
export function serMarker(opts: ChartOptsInternal, markerColor: string, seriesColor: string): string {
	const spPr = el('c:spPr', null, [
		raw(markerColor === 'transparent' ? voidEl('a:noFill') : genXmlColorSelection(markerColor)),
		raw(
			el('a:ln', { w: opts.lineDataSymbolLineSize, cap: 'flat' }, [
				raw(chartColorLineFill(opts.lineDataSymbolLineColor || seriesColor)),
				raw(voidEl('a:prstDash', { val: 'solid' })),
				raw(voidEl('a:round')),
			])
		),
		raw(voidEl('a:effectLst')),
	])
	return el('c:marker', null, [
		raw(voidEl('c:symbol', { val: opts.lineDataSymbol })),
		// Defaults to "auto" otherwise (but this is usually too small, so there is a default).
		opts.lineDataSymbolSize ? raw(voidEl('c:size', { val: opts.lineDataSymbolSize })) : null,
		raw(spPr),
	])
}
/**
 * Emit the `<a:latin>/<a:ea>/<a:cs>` font trio for a chart text run.
 *
 * In DrawingML run properties a typeface applies only to the script class of
 * its element: `<a:latin>` covers Latin/ASCII, `<a:ea>` covers East Asian, and
 * `<a:cs>` covers complex scripts. Emitting `<a:latin>` alone leaves East Asian
 * (e.g. Chinese) and complex-script glyphs falling back to the theme font, so a
 * user-specified font never takes effect for that text — most visibly on
 * PowerPoint for Mac. Stamping the same typeface onto all three classes is what
 * choosing a font in PowerPoint's UI does.
 * @param {string} typeface - font face name
 * @return {string} `<a:latin/><a:ea/><a:cs/>` XML
 */

export function createChartTextFonts(typeface: string): string {
	// Every caller passes a caller-supplied font option (dataLabelFontFace, catAxisLabelFontFace,
	// legendFontFace, ...), so escaping happens here — one site covers all of them. voidEl()'s
	// attrs escape by construction: an unescaped `"` or `&` would otherwise close the attribute
	// early and emit a non-parseable chart part.
	return voidEl('a:latin', { typeface }) + voidEl('a:ea', { typeface }) + voidEl('a:cs', { typeface })
}

/**
 * A chart's own `x`/`y` as inches, for folding into a manual title layout.
 *
 * `ChartOpts.x`/`y` are `Coord`, so `'10%'`, `'2in'` and `'50px'` are all legal and documented.
 * The title builder used to take them through an `as number` cast, so a string reached the
 * arithmetic below, `+` concatenated instead of adding, and `<c:x val="NaN"/>` came out. The
 * cast was both the cause and what hid it.
 *
 * A percentage is the one form that cannot be resolved here: it needs the slide axis, and the
 * chart part is built without a `PresLayout`. It returns `undefined` like an absent coordinate,
 * which leaves that axis on automatic layout rather than on a wrong number.
 * @param value - the chart's `x` or `y` as the caller spelled it
 */
function chartCoordInches(value: Coord | undefined): number | undefined {
	if (value === undefined || value === null) return undefined
	// A bare number is inches by definition. Taking it directly also keeps `coordToEmu` from
	// re-raising `coord/bare-number-is-inches`, which the graphic frame already reports for it.
	if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
	if (value.trim().endsWith('%')) return undefined
	const emu = coordToEmu(value, 0)
	return Number.isFinite(emu) ? emu / EMU_PER_INCH : undefined
}

export function genXmlTitle(opts: MaybeUndefined<ChartPropsTitle>, chartX?: Coord, chartY?: Coord): string {
	// `sizeAttr` is empty when the caller set no font size — PowerPoint then picks the default —
	// and interpolating it empty leaves TWO spaces between the tag name and `b=`. Those are
	// emitted bytes, and `el()` writes exactly one space before an attribute by design, so the
	// two run-property tags below stay hand-written. Dropping the padding would be the real fix —
	// no XML consumer can see it — but that is an output change, and a whitespace-only diff is a
	// stop rather than a cleanup. The chart flatten did NOT take it: that space is inside a tag
	// rather than between elements, which is a different claim needing different evidence, and
	// `prove-whitespace` freezes intra-tag whitespace so it stays visible. See
	// `docs/chart-whitespace-flatten.md`. Two other sites in this directory are in the same
	// position; the ratchet header lists them.
	// An empty `sizeAttr` spells one state and only one: the caller stated no size, which the
	// axis-title path still reaches (`catAxisTitleFontSize` and its two siblings are optional).
	// A stated `0` is a stated value, so it clamps and warns like every other size rather than
	// falling through to the same silence.
	const sizeAttr = opts.fontSize == null ? '' : `sz="${clampFontSizeSz(opts.fontSize, 'title fontSize')}"`
	const runAttrs = ` ${sizeAttr} b="${xsdBool(opts.titleBold)}" i="${xsdBool(opts.titleItalic)}" u="${opts.titleUnderline ? 'sng' : 'none'}" strike="noStrike">`
	const runChildren =
		genXmlColorSelection(opts.color ?? DEF_FONT_COLOR) + createChartTextFonts(opts.fontFace ?? 'Arial')

	// NOTE: manualLayout x/y vals are *relative to the entire slide*. Each axis is independent in
	// CT_ManualLayout: omitting xMode/x (or yMode/y) leaves that axis on automatic layout, so a
	// caller can center horizontally while still applying a manual vertical offset (and vice-versa).
	// Schema order is xMode, yMode, x, y.
	/**
	 * Fold a slide-relative offset into the fraction-of-chart value `c:x`/`c:y` want, or
	 * `undefined` when it does not read as a finite number — `ST_Double` has no `NaN`, and
	 * omitting the axis is what CT_ManualLayout already means by "leave this one automatic".
	 */
	const edgeFraction = (offset: number): number | undefined => {
		if (!Number.isFinite(offset)) return undefined
		let val = offset === 0 ? 0 : (offset * (offset / 5)) / 10
		if (val >= 1) val = val / 10
		if (val >= 0.1) val = val / 10
		return val
	}
	/**
	 * The chart's own offset along one axis, in inches, to fold into the caller's `titlePos`.
	 *
	 * Resolved here rather than by the caller so the warning fires only where the value is
	 * actually used — a chart with `showTitle` and no `titlePos` folds nothing and says nothing.
	 */
	const chartOffset = (value: Coord | undefined, axis: 'x' | 'y'): number => {
		const inches = chartCoordInches(value)
		if (inches !== undefined) return inches
		if (value !== undefined && value !== null)
			warn(
				'chart/option-out-of-range',
				`chart ${axis} "${String(value)}" needs a slide layout to resolve and the chart part is built without one; ` +
					`titlePos.${axis} is placed without it.`
			)
		return 0
	}
	// `Number.isFinite`, not `typeof === 'number'`: the latter is the one numeric guard `NaN`
	// passes, and a `titlePos: { x: NaN }` used to reach the attribute as written.
	const xVal = Number.isFinite(opts.titlePos?.x)
		? edgeFraction((opts.titlePos?.x ?? 0) + chartOffset(chartX, 'x'))
		: undefined
	const yVal = Number.isFinite(opts.titlePos?.y)
		? edgeFraction((opts.titlePos?.y ?? 0) + chartOffset(chartY, 'y'))
		: undefined
	let layout = voidEl('c:layout')
	if (xVal !== undefined || yVal !== undefined) {
		const modes =
			(xVal !== undefined ? voidEl('c:xMode', { val: 'edge' }) : '') +
			(yVal !== undefined ? voidEl('c:yMode', { val: 'edge' }) : '')
		const vals =
			(xVal !== undefined ? voidEl('c:x', { val: xVal }) : '') +
			(yVal !== undefined ? voidEl('c:y', { val: yVal }) : '')
		layout = el('c:layout', null, raw(el('c:manualLayout', null, raw(modes + vals))))
	}

	const paragraph = el('a:p', null, [
		raw(
			el(
				'a:pPr',
				opts.titleAlign === 'left' || opts.titleAlign === 'right' ? { algn: opts.titleAlign.slice(0, 1) } : null,
				raw('<a:defRPr' + runAttrs + runChildren + '</a:defRPr>')
			)
		),
		raw(el('a:r', null, [raw('<a:rPr' + runAttrs + runChildren + '</a:rPr>'), raw(el('a:t', null, opts.title ?? ''))])),
	])
	const rich = el('c:rich', null, [
		// Don't specify a rotation when none was asked for, so the default applies (which is
		// vertical on a category axis).
		raw(voidEl('a:bodyPr', { rot: opts.titleRotate ? convertAngleUnits(opts.titleRotate, 'titleRotate') : undefined })),
		raw(voidEl('a:lstStyle', null)),
		raw(paragraph),
	])
	return el('c:title', null, [raw(el('c:tx', null, raw(rich))), raw(layout), raw(voidEl('c:overlay', { val: 0 }))])
}

/**
 * A chart furniture line wrapped in its own element and a `<c:spPr>`: the gridlines a plot draws
 * behind its series, or the series lines a stacked bar draws between them.
 *
 * One `<a:ln>` body served two elements that differed only in their wrapping tag and in how the
 * caller decides whether to emit at all, so the tag is the parameter and the two guards stay
 * with their own entry points.
 * @param tag - `c:majorGridlines` or `c:serLines`
 * @param opts - the caller's line options
 * @param option - option name as the caller spells it, for the dash diagnostic
 */
function chartFurnitureLine(
	tag: 'c:majorGridlines' | 'c:serLines',
	opts: OptsChartGridLine | undefined,
	option: string
): string {
	if (gridLineSuppressed(opts)) return ''
	const stroke = gridLineStroke(opts as OptsChartGridLine)
	const line = el(
		'a:ln',
		{
			// `||` and not `??`: a width of `0` is not a hairline anyone asked for. On the three
			// gridline entry points `scrubGridLine` has already rejected `width <= 0` with a
			// warning and deleted it, so the two operators cannot differ there; `c:serLines` is
			// not scrubbed, and `||` is what keeps a zero from reaching the part undiagnosed.
			w: ptsToEmuLenient(stroke.width || DEF_CHART_GRIDLINE.width || 1),
			cap: createLineCap(stroke.cap || DEF_CHART_GRIDLINE.cap),
		},
		[
			raw(strokePaint(stroke, DEF_GRIDLINE_COLOR)),
			raw(voidEl('a:prstDash', { val: strokeDash(stroke, option) }) + voidEl('a:round')),
		]
	)
	return el(tag, null, raw(el('c:spPr', null, raw(line))))
}

/**
 * A plot's `<c:majorGridlines>`, or `''` when the caller asked for none.
 *
 * The "asked for none" test used to sit at each of the three axis builders as
 * `opts.catGridLine && opts.catGridLine.style !== 'none'`, which stopped being the whole test
 * once `type: 'none'` could say the same thing.
 * @param glOpts - the caller's gridline stroke, if any
 * @return {string} XML, or `''`
 */
export function createGridLineElement(glOpts: OptsChartGridLine | undefined): string {
	return chartFurnitureLine('c:majorGridlines', glOpts, 'gridLine dashType')
}

/**
 * The children of a `<c:dLbls>` (or a `<c:dLbl>`), in the ONE order `CT_DLbls` allows.
 *
 * `EG_DLblShared` sequences `numFmt, spPr, txPr, dLblPos`, then the six flags, then
 * `showLeaderLines` and `leaderLines`, with any per-point `<c:dLbl>` ahead of all of it and
 * `<c:extLst>` last. Five builders across four modules restated that sequence by hand, which is
 * the same shape as the axis-units bug: a child in the wrong place is a repair prompt, not a
 * wrong-looking chart. `dLblShowFlags` already owned the six-flag run for this reason and the
 * sequence around it was left copied.
 *
 * Every part is optional and stated by the caller, so the drift between the five shows up as a
 * missing argument rather than as a silently different chart: only one of them honoured
 * `dataLabelBkgrdColors`, only one a per-series `dataLabelFormatCode`, and two hard-coded
 * `showLeaderLines` to on while four read the option.
 *
 * @param parts - the children this site emits, each already built
 * @param tag - `c:dLbls` for a group of labels, `c:dLbl` for one point's
 */
export function dLblsBlock(
	parts: {
		/** Per-point `<c:dLbl>` blocks, or a `<c:idx>` on a `c:dLbl`. Emitted first, in order. */
		lead?: string | undefined
		numFmt?: string | undefined
		spPr?: string | undefined
		txPr?: string | undefined
		dLblPos?: string | undefined
		flags: XmlChild[]
		showLeaderLines?: string | undefined
		leaderLines?: string | undefined
		extLst?: string | undefined
	},
	tag: 'c:dLbls' | 'c:dLbl' = 'c:dLbls'
): string {
	return el(tag, null, [
		parts.lead ? raw(parts.lead) : null,
		parts.numFmt ? raw(parts.numFmt) : null,
		parts.spPr ? raw(parts.spPr) : null,
		parts.txPr ? raw(parts.txPr) : null,
		parts.dLblPos ? raw(parts.dLblPos) : null,
		...parts.flags,
		parts.showLeaderLines ? raw(parts.showLeaderLines) : null,
		parts.leaderLines ? raw(parts.leaderLines) : null,
		parts.extLst ? raw(parts.extLst) : null,
	])
}

/**
 * The `<c:numFmt>` a data-label block carries: the caller's format code, or `General`.
 *
 * `(x ?? '') || 'General'` stood at five sites verbatim. The two coalescers are not
 * interchangeable and both are load-bearing: `??` keeps an explicit empty string from becoming
 * `undefined`, and `||` turns that empty string into the schema's own default.
 * @param formatCode - the caller's format code, if any
 */
export function dLblNumFmt(formatCode: string | undefined): string {
	return voidEl('c:numFmt', { formatCode: (formatCode ?? '') || 'General', sourceLinked: 0 })
}

/**
 * The plot-level `<c:dLbls>` block: number format, run properties, and the six `show*` flags.
 *
 * Shared by the category-axis and scatter plots, which carried a 20-line copy each. The two
 * differed in exactly one thing that reaches the bytes -- whether `<c:showLeaderLines>` is
 * emitted -- so that is the parameter. (They also differed in `? 1 : 0` versus `? '1' : '0'`
 * for the bold/italic flags, which is the same byte either way.)
 *
 * `<c:showLeaderLines>` is last in CT_DLbls' group and optional, which is why omitting it is a
 * legal shape and not a schema gap: leader lines connect a moved label back to its point, and a
 * scatter plot has no `<c:dLblPos>` layout that moves one.
 *
 * @param opts - the chart's normalized options
 * @param leaderLines - emit the trailing `<c:showLeaderLines>` (category-axis plots only)
 */
export function chartDataLabels(opts: ChartOptsInternal, leaderLines: boolean): string {
	return dLblsBlock({
		numFmt: dLblNumFmt(opts.dataLabelFormatCode),
		txPr: labelTextProps(dataLabelDefRPr(opts)),
		dLblPos: opts.dataLabelPosition ? voidEl('c:dLblPos', { val: opts.dataLabelPosition }) : undefined,
		flags: dLblShowFlags({ val: xsdBool(opts.showValue), serName: xsdBool(opts.showSerName) }),
		showLeaderLines: leaderLines ? voidEl('c:showLeaderLines', { val: xsdBool(opts.showLeaderLines) }) : undefined,
	})
}

/**
 * An `<a:solidFill>` of the theme text colour dimmed towards the background — the grey the chart
 * furniture (stock hi-low lines and up/down bars, the data-table grid) is drawn in.
 * `lumMod`/`lumOff` are the two halves of one dimming: how much of the luminance survives, and
 * how much white is mixed back in.
 * @param lumMod - surviving luminance, in thousandths of a percent
 * @param lumOff - luminance added back, in thousandths of a percent
 */
export function dimmedTextFill(lumMod: number, lumOff: number): string {
	return el(
		'a:solidFill',
		null,
		raw(
			el('a:schemeClr', { val: 'tx1' }, [
				raw(voidEl('a:lumMod', { val: lumMod })),
				raw(voidEl('a:lumOff', { val: lumOff })),
			])
		)
	)
}

/** The hairline outline drawn in {@link dimmedTextFill}, at the same three dimming levels. */
export function dimmedTextLine(lumMod: number, lumOff: number): string {
	return el('a:ln', { w: 9525, cap: 'flat', cmpd: 'sng', algn: 'ctr' }, [
		raw(dimmedTextFill(lumMod, lumOff)),
		raw(voidEl('a:round')),
	])
}

/**
 * Build a `<c:pt>` numeric-cache data point, or '' to leave a gap.
 *
 * A `<c:v>` holding `NaN`, `INF` or `Infinity` makes PowerPoint refuse the package with
 * 0x80070570, the corrupt-file error, so a non-finite number is skipped with a warning per the
 * library's "warn rather than emit a degenerate result" policy. Null/undefined are intentional
 * gaps and are skipped silently (a sparse, idx-keyed cache is valid).
 *
 * **PowerPoint is the only oracle for this.** `<c:v>` is `s:ST_Xstring` in the schema, not
 * `xsd:double` — nothing above is a schema violation, and the OpenXmlValidator reports every one
 * of those three packages clean. An *empty* `<c:v>` is the case that shows why the distinction is
 * worth stating rather than reasoning about: it looks equally invalid and it is not, opening
 * cleanly and reading back as an empty point. See `pieValues` in
 * `src/gen/chart/plot-pie.ts`, which emitted both spellings until the measurement was made.
 * @param idx - zero-based data-point index (emitted as `idx`)
 * @param value - numeric value (or null/undefined gap)
 */
export function numCachePt(idx: number, value: number | null | undefined): string {
	if (value == null) return ''
	if (!Number.isFinite(value)) {
		warn('chart/non-finite-value', `chart value "${value}" at index ${idx} is not a finite number; data point omitted.`)
		return ''
	}
	return el('c:pt', { idx }, raw(el('c:v', null, value)))
}

/**
 * Build the error-bar elements (`<c:errBars>`) for a single series.
 *
 * Schema position (CT_*Ser): after `dLbls`/`trendline`, before `cat`/`val` (bar/line/area)
 * or `xVal`/`yVal` (scatter). CT_ErrBars child order is errDir → errBarType → errValType →
 * noEndCap → plus → minus → val → spPr.
 *
 * @param chartType - chart this series belongs to (used to bound how many bars are legal)
 * @param errorBars - one config, or an array (X+Y) for scatter/area; bar/line keep only the first
 * @param obj - the series data object (only `name`, for warnings)
 */
export function makeChartErrorBarsXml(
	chartType: ChartType,
	errorBars: ChartErrorBarOptions | ChartErrorBarOptions[] | undefined,
	obj: OptsChartDataInternal
): string {
	if (!errorBars) return ''
	const bars = Array.isArray(errorBars) ? errorBars : [errorBars]
	// CT_BarSer/CT_LineSer allow a single <c:errBars>; only scatter/area permit two (x + y).
	const maxBars = chartType === ChartType.scatter || chartType === ChartType.area ? 2 : 1
	let strXml = ''

	bars.slice(0, maxBars).forEach((eb) => {
		if (!eb) return
		const valueType = eb.valueType || 'fixedVal'
		const barType = eb.barType || 'both'
		const direction = eb.direction || 'y'

		let children =
			voidEl('c:errDir', { val: direction }) +
			voidEl('c:errBarType', { val: barType }) +
			voidEl('c:errValType', { val: valueType }) +
			voidEl('c:noEndCap', { val: xsdBool(eb.noEndCap) })

		if (valueType === 'cust') {
			// Custom amounts: <c:plus>/<c:minus> each hold a number source (we emit <c:numLit>).
			// `barType` decides which sides are present; warn (don't silently drop) on a missing side.
			if (barType !== 'minus') {
				if (!eb.plusValues?.length)
					warn(
						'chart/error-bars-missing-values',
						`chart series "${obj.name}" errorBars valueType 'cust' needs \`plusValues\` for barType '${barType}'.`
					)
				children += makeErrBarNumLit('plus', eb.plusValues || [])
			}
			if (barType !== 'plus') {
				if (!eb.minusValues?.length)
					warn(
						'chart/error-bars-missing-values',
						`chart series "${obj.name}" errorBars valueType 'cust' needs \`minusValues\` for barType '${barType}'.`
					)
				children += makeErrBarNumLit('minus', eb.minusValues || [])
			}
		} else if (valueType !== 'stdErr') {
			// fixedVal / percentage / stdDev use a single magnitude (stdErr derives it from the data).
			children += voidEl('c:val', { val: eb.value ?? 1 })
		}

		// The error bar is a stroke like any other, so it takes the same keys: `width` (with
		// `size` still accepted), plus the `dashType`, `cap` and `transparency` the old
		// two-key shape could not say. A bar that states none of them keeps emitting no
		// `<c:spPr>` at all, which is what leaves PowerPoint's own error-bar style in charge.
		// oxlint-disable-next-line typescript/no-deprecated -- `size` is the pre-4.0 spelling of `width` and is still honoured.
		const width = eb.width ?? eb.size
		// A bar that names a colour or a transparency paints one; `type: 'none'` paints an
		// explicit `<a:noFill/>`; a bar that names only a width or a dash leaves the paint to
		// PowerPoint's own error-bar style, which is what "no `<a:solidFill>`" has always meant
		// here. `type: 'dash'` must NOT conjure a colour out of the default.
		const paint =
			eb.type === 'none'
				? voidEl('a:noFill')
				: eb.color !== undefined || eb.transparency !== undefined
					? strokePaint(eb, DEF_GRIDLINE_COLOR)
					: ''
		const dash =
			eb.dashType !== undefined || eb.type === 'dash'
				? voidEl('a:prstDash', { val: strokeDash(eb, 'errorBars dashType') })
				: ''
		if (paint || dash || width != null || eb.cap !== undefined) {
			children += el(
				'c:spPr',
				null,
				raw(
					el(
						'a:ln',
						{
							w: width != null ? ptsToEmuLenient(width) : undefined,
							cap: eb.cap !== undefined ? createLineCap(eb.cap) : undefined,
						},
						raw(paint + dash)
					)
				)
			)
		}

		strXml += el('c:errBars', null, raw(children))
	})

	return strXml
}

/**
 * Build a `<c:plus>`/`<c:minus>` number-literal source for custom error-bar amounts.
 * @param tag - `'plus'` or `'minus'`
 * @param values - per-point magnitudes (index-aligned with the series values)
 */
function makeErrBarNumLit(tag: 'plus' | 'minus', values: number[]): string {
	const numLit = el('c:numLit', null, [
		raw(el('c:formatCode', null, 'General')),
		raw(voidEl('c:ptCount', { val: values.length })),
		raw(values.map((value, idx) => numCachePt(idx, value)).join('')),
	])
	return el(`c:${tag}`, null, raw(numLit))
}

/**
 * Build a `<c:serLines>` ("Series Lines") element for a bar chart.
 * @param opt - `true` for PowerPoint automatic styling, an {@link OptsChartGridLine}
 *   to customize the line, or falsy / `{ style: 'none' }` to omit the element.
 */
export function createSerLinesElement(opt?: boolean | OptsChartGridLine): string {
	if (!opt) return ''
	if (opt === true) return voidEl('c:serLines')
	return chartFurnitureLine('c:serLines', opt, 'serLine dashType')
}

/**
 * Build the `<c:leaderLines>` element for pie/doughnut data labels.
 *
 * Schema position: inside `<c:dLbls>`, immediately after `<c:showLeaderLines>`
 * (CT_DLbls / Group_DLbls order: showLeaderLines → leaderLines).
 *
 * Returns `''` unless the caller both enabled leader lines (`showLeaderLines`)
 * and configured their appearance (`leaderLineColor` / `leaderLineSize`). When
 * appearance is unset we leave the element off so PowerPoint applies its
 * automatic leader-line color, matching prior behavior.
 *
 * @param opts - chart options (reads `showLeaderLines`, `leaderLineColor`, `leaderLineSize`)
 */
export function createLeaderLinesElement(opts: ChartOptsInternal): string {
	if (!opts.showLeaderLines) return ''
	if (!opts.leaderLineColor && opts.leaderLineSize == null) return ''
	const w = ptsToEmuLenient(opts.leaderLineSize ?? 0.75)
	const color = opts.leaderLineColor ?? '808080'
	const line = el('a:ln', { w, cap: 'flat' }, [
		raw(genXmlColorSelection(color)),
		raw(voidEl('a:prstDash', { val: 'solid' })),
		raw(voidEl('a:round')),
	])
	return el('c:leaderLines', null, raw(el('c:spPr', null, [raw(line), raw(voidEl('a:effectLst'))])))
}

/**
 * Build a single custom `<c:dLbl>` (rich-text data label) overriding one data point's label.
 *
 * Used when a series supplies explicit per-point label text: the emitted `<c:idx>` pins the
 * override to that point, and the `<c:rich>` run carries the label's own font styling (size,
 * bold, italic, color, face) resolved from the chart-level dataLabel* options. All the
 * `show*` flags are forced off so only the literal `text` renders (no value/category/percent).
 * @param idx - zero-based data-point index this label overrides
 * @param text - the literal label text (XML-escaped here)
 * @param opts - chart options supplying dataLabel font/color defaults and `lang`
 * @return {string} a `<c:dLbl>` element
 */
export function makeCustomDLblXml(idx: number, text: string, opts: ChartOptsInternal): string {
	const lang = opts.lang || 'en-US'
	const runChildren = labelFontChildren(opts)
	const fontAttrs = labelFontAttrs(opts)
	const paragraph = el('a:p', null, [
		raw(el('a:pPr', null, raw(el('a:defRPr', fontAttrs, runChildren)))),
		raw(el('a:r', null, [raw(el('a:rPr', { lang, ...fontAttrs, dirty: 0 }, runChildren)), raw(el('a:t', null, text))])),
	])
	const rich = el('c:rich', null, [raw(voidEl('a:bodyPr')), raw(voidEl('a:lstStyle')), raw(paragraph)])
	return el('c:dLbl', null, [
		raw(voidEl('c:idx', { val: idx })),
		raw(el('c:tx', null, raw(rich))),
		...dLblShowFlags({}),
	])
}

/**
 * The `<a:ln>` a chart-level `dataBorder` paints around every data point — the bubble, the
 * pie slice, the bar.
 *
 * A near-sibling of {@link createChartBorderLine} rather than the same thing, and the three
 * remaining differences are all deliberate: a data-point outline defaults to 0.75pt against a
 * chart border's 1pt, to `363636` against `666666`, and is always solid because `dataBorder`
 * has no dash spelling of its own. It also has no `type: 'none'` arm, because the caller
 * decides whether to emit it at all. Three plot builders carried a byte-identical copy of it.
 *
 * `cap` is the one thing that genuinely varies: the bubble and pie paths pin `flat`, the
 * category-axis path resolves the chart's own `lineCap`.
 * @param dataBorder - the chart's `dataBorder`
 * @param cap - the `cap` attribute value, already resolved
 */
export function createDataBorderLine(dataBorder: BorderProps, cap: string): string {
	return borderLine('a:ln', dataBorder, {
		defaultWidth: 0.75,
		defaultColor: '363636',
		cap,
		dash: 'solid',
		tail: [voidEl('a:round')],
	})
}

/**
 * The `<a:ln>` a chart-area or plot-area `border` paints, or an explicit `<a:noFill>` when the
 * caller asked for none. See {@link createDataBorderLine} for why the two are separate.
 *
 * `type: 'dash'` here means the `dash` preset, not the `sysDash` a table border's `type`
 * resolves to — a chart border has no `dashType` of its own to name a preset with, so the
 * three-way switch is the whole vocabulary and it picks the plain one.
 * @param border - the chart's `border` (`type`, `color`, `width`, `transparency`)
 */
export function createChartBorderLine(border: BorderProps): string {
	if (border.type === 'none') return noStrokeLine()
	return borderLine('a:ln', border, {
		defaultWidth: 1,
		defaultColor: '666666',
		cap: 'flat',
		dash: border.type === 'dash' ? 'dash' : 'solid',
		tail: [voidEl('a:round')],
	})
}

/**
 * Build `<c:dPt>` entries for a series in the bar/line/area/scatter loops.
 *
 * Merges two sources into a single `c:dPt` per index so we never emit a
 * duplicate `<c:idx>` (which corrupts the chart):
 * - legacy single-series color-vary fills (bar/scatter), supplied via `varyColors`
 * - per-point `pointStyles` border/fill overrides
 *
 * Must be emitted in schema position *before* `c:dLbls` (CT_*Ser order).
 * RADAR is skipped: extra per-point markup historically corrupts the chart.
 *
 * @param chartType  - series chart type
 * @param obj        - series data (reads `values`, `pointStyles`)
 * @param opts       - chart options (fill/shadow/lineSize context)
 * @param varyColors - color array when single-series color-vary applies, else `null`
 */
export function makeSeriesDataPointsXml(
	chartType: ChartType,
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	varyColors: string[] | null
): string {
	if (chartType === ChartType.radar) return ''
	const pointStyles = obj.pointStyles
	if (!varyColors && !pointStyles?.length) return ''

	const isBar = chartType === ChartType.bar || chartType === ChartType.bar3d
	const isScatter = chartType === ChartType.scatter
	let xml = ''
	dataValues(obj).forEach((value, index) => {
		const ptStyle = pointStyles?.[index]
		const arrColors = varyColors
			? value < 0
				? opts.invertedColors || opts.chartColors || BARCHART_COLORS
				: varyColors
			: null
		const fillColor = ptStyle?.fill || (arrColors ? paletteColor(arrColors, index) : null)
		const pattern = ptStyle?.pattern
		const border = ptStyle?.border
		// Nothing to style for this point -> omit the c:dPt entirely
		if (!fillColor && !pattern && !border) return

		let shape = ''
		if ((isBar || isScatter) && opts.lineSize === 0 && !border && !ptStyle?.fill && !pattern) {
			// Preserve legacy color-vary behavior: hide outline when lineSize===0
			shape = noStrokeLine()
		} else {
			// Pattern fill takes precedence over a solid fill (OOXML allows only one fill per c:dPt).
			// Default the pattern foreground to this point's resolved color so it reads as a hatched bar.
			if (pattern) {
				shape = genXmlPatternFill(fillColor && !pattern.fgColor ? { ...pattern, fgColor: fillColor } : pattern)
			} else if (fillColor) {
				// BAR3D color-vary historically tints the edge line, not the face fill
				shape =
					chartType === ChartType.bar3d
						? el('a:ln', null, raw(genXmlColorSelection(fillColor)))
						: genXmlColorSelection(fillColor)
			}
			if (border) shape += createChartBorderLine(border)
		}
		xml += el('c:dPt', null, [
			raw(voidEl('c:idx', { val: index })),
			isBar ? raw(voidEl('c:invertIfNegative', { val: 0 })) : null,
			raw(voidEl('c:bubble3D', { val: 0 })),
			raw(el('c:spPr', null, [raw(shape), raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW))])),
		])
	})
	return xml
}

/**
 * A `<c:tx>` series-name block: the `<c:f>` formula plus the one-point `<c:strCache>` mirroring
 * the header cell it points at. The string counterpart of {@link numRefBlock}.
 *
 * The five copies this replaces were all hand-built strings, so routing them here takes
 * fifty-odd hand-written delimiters out of `src/gen/chart/` rather than moving them into one
 * file.
 * @param ref - the `<c:f>` formula, from {@link sheetCellRef} or written inline
 * @param name - the series name to cache; escaped on the way into `<c:v>`
 */
export function strRefBlock(ref: string, name: string): string {
	const pt = el('c:pt', { idx: 0 }, raw(el('c:v', null, name)))
	const strCache = el('c:strCache', null, [raw(voidEl('c:ptCount', { val: 1 })), raw(pt)])
	const strRef = el('c:strRef', null, [raw(el('c:f', null, ref)), raw(strCache)])
	return el('c:tx', null, raw(strRef))
}

/**
 * A `<c:cat>` category reference: the `<c:f>` formula plus the `<c:numCache>`/`<c:strCache>` that
 * mirrors the label cells it points at.
 *
 * `kind` picks the cache: `'num'` for numeric categories — dates, which need a `<c:formatCode>` so
 * PowerPoint renders them as dates rather than serial numbers — and `'str'` for text ones. Stock
 * and cat-axis each carried both arms of that choice, and surface carried the string one, five
 * copies of the same `<c:f>` + `<c:ptCount>` + `<c:pt idx>` loop.
 *
 * The `<c:multiLvlStrRef>` arm of a cat-axis is deliberately not folded in: multi-level
 * categories nest a `<c:lvl>` per label group, which is a different shape rather than a
 * different spelling.
 * @param kind - `num` for a numeric (date) category cache, `str` for a text one
 * @param ref - the `<c:f>` formula
 * @param labels - the category labels, in order; escaped on the way into `<c:v>`
 * @param formatCode - the cached `<c:formatCode>`; numeric caches only
 */
export function catRefBlock(kind: 'num' | 'str', ref: string, labels: string[], formatCode?: string): string {
	const points = labels.map((label, idx) => el('c:pt', { idx }, raw(el('c:v', null, label)))).join('')
	const cache = el(`c:${kind}Cache`, null, [
		formatCode === undefined ? null : raw(el('c:formatCode', null, formatCode)),
		raw(voidEl('c:ptCount', { val: labels.length }) + points),
	])
	return el(`c:${kind}Ref`, null, [raw(el('c:f', null, ref)), raw(cache)])
}

/**
 * A numeric-reference block: the `<c:f>` formula plus the `<c:numCache>` that mirrors the cells
 * it points at.
 *
 * Scatter and bubble each carried their own copy of this twice over, four near-identical
 * twelve-line blocks, and bubble's per-point sizes were a fifth — the same block under a third
 * wrapping tag with a constant format code. `values` is the exact point list to cache, so the
 * caller decides what a gap is: the Y series is emitted against the X series' length (a caller
 * may supply fewer Y values than X — a timeline with only the first few months filled in), and
 * the shorter array's tail arrives here as `undefined`, which {@link numCachePt} skips.
 *
 * `ptCount` is the length of the cached *range* — the number of cells `ref` spans — which is not
 * always `values.length`. A category chart caches one column against the chart's category count
 * and may hold fewer values than that, so it states the count explicitly; scatter and bubble pad
 * their arrays to the range instead and let the default stand. Both spellings mean the same
 * thing, and having only one of them derived by accident is how they drifted apart.
 *
 * @param tag - the wrapping element: `c:val`, `c:xVal`, `c:yVal` or `c:bubbleSize`
 * @param ref - the `<c:f>` formula, from {@link sheetRangeRef} or written inline
 * @param formatCode - the cached `<c:formatCode>`
 * @param values - the points to cache, in order; `null`/`undefined` entries are gaps
 * @param ptCount - cells spanned by `ref`; defaults to `values.length`
 */
export function numRefBlock(
	tag: 'c:val' | 'c:xVal' | 'c:yVal' | 'c:bubbleSize',
	ref: string,
	formatCode: string,
	values: Array<number | null | undefined>,
	ptCount: number = values.length
): string {
	const numCache = el('c:numCache', null, [
		raw(el('c:formatCode', null, formatCode)),
		raw(voidEl('c:ptCount', { val: ptCount })),
		raw(values.map((value, idx) => numCachePt(idx, value)).join('')),
	])
	const numRef = el('c:numRef', null, [raw(el('c:f', null, ref)), raw(numCache)])
	return el(tag, null, raw(numRef))
}
