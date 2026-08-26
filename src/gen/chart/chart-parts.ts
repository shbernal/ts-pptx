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
import type { BorderProps, ChartErrorBarOptions, ChartPropsTitle, OptsChartGridLine } from '../../types/index.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { warn } from '../../diagnostics.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection, genXmlPatternFill } from '../drawingml/fill.js'
import { createLineCap, resolveBorderWidth } from '../drawingml/line.js'
import { convertAngleUnits, ptsToEmuLenient } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { dataValues } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'

export const VALID_CHART_TIME_UNITS = ['days', 'months', 'years']

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
 * The palette colour for series/point `idx`, cycling back to the start once the palette runs out.
 *
 * Every palette lookup in this directory goes through here so that a deck with more series or
 * data points than the palette has entries still emits the same bytes on every build.
 */
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

export function paletteColor(palette: readonly string[], idx: number, fallback = '000000'): string {
	if (palette.length === 0) return fallback
	return palette[idx % palette.length] ?? fallback
}

// DEF_CHART_GRIDLINE.color is optional on the type but always present on the constant.
export const DEF_GRIDLINE_COLOR: string = DEF_CHART_GRIDLINE.color ?? '888888'

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

export function genXmlTitle(opts: ChartPropsTitle, chartX?: number, chartY?: number): string {
	const align =
		opts.titleAlign === 'left' || opts.titleAlign === 'right'
			? `<a:pPr algn="${opts.titleAlign.slice(0, 1)}">`
			: '<a:pPr>'
	const rotate = opts.titleRotate
		? `<a:bodyPr rot="${convertAngleUnits(opts.titleRotate, 'titleRotate')}"/>`
		: '<a:bodyPr/>' // don't specify rotation to get default (ex. vertical for cat axis)
	const sizeAttr = opts.fontSize ? `sz="${ptToHundredths(opts.fontSize)}"` : '' // only set the font size if specified.  Powerpoint will handle the default size
	const titleBold = opts.titleBold ? 1 : 0
	const titleItalic = opts.titleItalic ? 1 : 0
	const titleUnderline = opts.titleUnderline ? 'sng' : 'none'

	let layout = '<c:layout/>'
	const hasX = opts.titlePos && typeof opts.titlePos.x === 'number'
	const hasY = opts.titlePos && typeof opts.titlePos.y === 'number'
	if (hasX || hasY) {
		// NOTE: manualLayout x/y vals are *relative to entire slide*. Each axis is
		// independent in CT_ManualLayout: omitting xMode/x (or yMode/y) leaves that
		// axis on automatic layout, so a caller can center horizontally while still
		// applying a manual vertical offset (and vice-versa).
		// Schema order is xMode, yMode, x, y.
		let modes = ''
		let vals = ''
		if (hasX) {
			const totalX = (opts.titlePos?.x ?? 0) + (chartX ?? 0)
			let valX = totalX === 0 ? 0 : (totalX * (totalX / 5)) / 10
			if (valX >= 1) valX = valX / 10
			if (valX >= 0.1) valX = valX / 10
			modes += '<c:xMode val="edge"/>'
			vals += `<c:x val="${valX}"/>`
		}
		if (hasY) {
			const totalY = (opts.titlePos?.y ?? 0) + (chartY ?? 0)
			let valY = totalY === 0 ? 0 : (totalY * (totalY / 5)) / 10
			if (valY >= 1) valY = valY / 10
			if (valY >= 0.1) valY = valY / 10
			modes += '<c:yMode val="edge"/>'
			vals += `<c:y val="${valY}"/>`
		}
		layout = `<c:layout><c:manualLayout>${modes}${vals}</c:manualLayout></c:layout>`
	}

	return `<c:title>
      <c:tx>
        <c:rich>
          ${rotate}
          <a:lstStyle/>
          <a:p>
            ${align}
            <a:defRPr ${sizeAttr} b="${titleBold}" i="${titleItalic}" u="${titleUnderline}" strike="noStrike">
              ${genXmlColorSelection(opts.color || DEF_FONT_COLOR)}
              ${createChartTextFonts(opts.fontFace || 'Arial')}
            </a:defRPr>
          </a:pPr>
          <a:r>
            <a:rPr ${sizeAttr} b="${titleBold}" i="${titleItalic}" u="${titleUnderline}" strike="noStrike">
              ${genXmlColorSelection(opts.color || DEF_FONT_COLOR)}
              ${createChartTextFonts(opts.fontFace || 'Arial')}
            </a:rPr>
            ${el('a:t', null, opts.title ?? '')}
          </a:r>
        </a:p>
        </c:rich>
      </c:tx>
      ${layout}
      <c:overlay val="0"/>
    </c:title>`
}

/**
 * Create Grid Line Element
 * @param {OptsChartGridLine} glOpts {size, color, style}
 * @return {string} XML
 */
export function createGridLineElement(glOpts: OptsChartGridLine): string {
	const line = el(
		'a:ln',
		{
			w: ptsToEmuLenient(glOpts.size || DEF_CHART_GRIDLINE.size || 1),
			cap: createLineCap(glOpts.cap || DEF_CHART_GRIDLINE.cap),
		},
		[
			raw('  ' + el('a:solidFill', null, raw(createColorElement(glOpts.color || DEF_GRIDLINE_COLOR)))),
			raw('   ' + voidEl('a:prstDash', { val: glOpts.style || DEF_CHART_GRIDLINE.style }) + voidEl('a:round')),
		],
		{ openPrefix: '  ', closePrefix: '  ' }
	)

	return el('c:majorGridlines', null, raw(el('c:spPr', null, raw(line), { openPrefix: ' ', closePrefix: ' ' })))
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
	let xml = '  <c:dLbls>'
	xml += '    ' + voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
	xml += '    <c:txPr>'
	xml += '      <a:bodyPr/>'
	xml += '      <a:lstStyle/>'
	xml += '      <a:p><a:pPr>'
	xml += `        <a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" u="none">`
	xml += genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)
	xml += '          ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
	xml += '        </a:defRPr>'
	xml += '      </a:pPr></a:p>'
	xml += '    </c:txPr>'
	if (opts.dataLabelPosition) xml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>'
	xml += '    <c:showLegendKey val="0"/>'
	xml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
	xml += '    <c:showCatName val="0"/>'
	xml += '    <c:showSerName val="' + (opts.showSerName ? '1' : '0') + '"/>'
	xml += '    <c:showPercent val="0"/>'
	xml += '    <c:showBubbleSize val="0"/>'
	if (leaderLines) xml += `    <c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
	xml += '  </c:dLbls>'
	return xml
}

/**
 * Build a `<c:pt>` numeric-cache data point, or '' to leave a gap.
 *
 * `<c:v>` inside a `<c:numCache>` is an `xsd:double`; emitting `NaN`, `Infinity`
 * or an empty string yields an invalid value that makes PowerPoint report the
 * package as needing repair. Null/undefined are intentional gaps and are skipped
 * silently (a sparse, idx-keyed cache is valid); other non-finite numbers are
 * skipped with a warning, per the library's "warn rather than emit a degenerate
 * result" policy.
 * @param idx - zero-based data-point index (emitted as `idx`)
 * @param value - numeric value (or null/undefined gap)
 */
export function numCachePt(idx: number, value: number | null | undefined): string {
	if (value == null) return ''
	if (!Number.isFinite(value)) {
		warn('chart/non-finite-value', `chart value "${value}" at index ${idx} is not a finite number; data point omitted.`)
		return ''
	}
	return `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`
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

		strXml += '<c:errBars>'
		strXml += `<c:errDir val="${direction}"/>`
		strXml += `<c:errBarType val="${barType}"/>`
		strXml += `<c:errValType val="${valueType}"/>`
		strXml += `<c:noEndCap val="${eb.noEndCap ? '1' : '0'}"/>`

		if (valueType === 'cust') {
			// Custom amounts: <c:plus>/<c:minus> each hold a number source (we emit <c:numLit>).
			// `barType` decides which sides are present; warn (don't silently drop) on a missing side.
			if (barType !== 'minus') {
				if (!eb.plusValues?.length)
					warn(
						'chart/error-bars-missing-values',
						`chart series "${obj.name}" errorBars valueType 'cust' needs \`plusValues\` for barType '${barType}'.`
					)
				strXml += makeErrBarNumLit('plus', eb.plusValues || [])
			}
			if (barType !== 'plus') {
				if (!eb.minusValues?.length)
					warn(
						'chart/error-bars-missing-values',
						`chart series "${obj.name}" errorBars valueType 'cust' needs \`minusValues\` for barType '${barType}'.`
					)
				strXml += makeErrBarNumLit('minus', eb.minusValues || [])
			}
		} else if (valueType !== 'stdErr') {
			// fixedVal / percentage / stdDev use a single magnitude (stdErr derives it from the data).
			strXml += `<c:val val="${eb.value ?? 1}"/>`
		}

		if (eb.color || eb.size != null) {
			strXml += '<c:spPr><a:ln'
			strXml += eb.size != null ? ` w="${ptsToEmuLenient(eb.size)}"` : ''
			strXml += '>'
			strXml += eb.color ? genXmlColorSelection(eb.color) : ''
			strXml += '</a:ln></c:spPr>'
		}

		strXml += '</c:errBars>'
	})

	return strXml
}

/**
 * Build a `<c:plus>`/`<c:minus>` number-literal source for custom error-bar amounts.
 * @param tag - `'plus'` or `'minus'`
 * @param values - per-point magnitudes (index-aligned with the series values)
 */
function makeErrBarNumLit(tag: 'plus' | 'minus', values: number[]): string {
	let strXml = `<c:${tag}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>`
	values.forEach((value, idx) => {
		strXml += numCachePt(idx, value)
	})
	strXml += `</c:numLit></c:${tag}>`

	return strXml
}

/**
 * Build a `<c:serLines>` ("Series Lines") element for a bar chart.
 * @param opt - `true` for PowerPoint automatic styling, an {@link OptsChartGridLine}
 *   to customize the line, or falsy / `{ style: 'none' }` to omit the element.
 */
export function createSerLinesElement(opt?: boolean | OptsChartGridLine): string {
	if (!opt) return ''
	if (opt === true) return '<c:serLines/>'
	if (opt.style === 'none') return ''
	const line = el(
		'a:ln',
		{
			w: ptsToEmuLenient(opt.size || DEF_CHART_GRIDLINE.size || 1),
			cap: createLineCap(opt.cap || DEF_CHART_GRIDLINE.cap),
		},
		[
			raw(el('a:solidFill', null, raw(createColorElement(opt.color || DEF_GRIDLINE_COLOR)))),
			raw(voidEl('a:prstDash', { val: opt.style || DEF_CHART_GRIDLINE.style }) + voidEl('a:round')),
		]
	)

	return el('c:serLines', null, raw(el('c:spPr', null, raw(line))))
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
	const color = opts.leaderLineColor || '808080'
	return (
		'<c:leaderLines><c:spPr>' +
		`<a:ln w="${w}" cap="flat">${genXmlColorSelection(color)}<a:prstDash val="solid"/><a:round/></a:ln>` +
		'<a:effectLst/></c:spPr></c:leaderLines>'
	)
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
	const sz = ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)
	const bold = opts.dataLabelFontBold ? '1' : '0'
	const italic = opts.dataLabelFontItalic ? '1' : '0'
	const color = createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)
	const face = opts.dataLabelFontFace || 'Arial'
	const lang = opts.lang || 'en-US'
	return (
		`<c:dLbl><c:idx val="${idx}"/>` +
		'<c:tx><c:rich><a:bodyPr/><a:lstStyle/>' +
		`<a:p><a:pPr><a:defRPr sz="${sz}" b="${bold}" i="${italic}" u="none" strike="noStrike">` +
		`<a:solidFill>${color}</a:solidFill>${createChartTextFonts(face)}</a:defRPr></a:pPr>` +
		`<a:r><a:rPr lang="${lang}" sz="${sz}" b="${bold}" i="${italic}" u="none" strike="noStrike" dirty="0">` +
		`<a:solidFill>${color}</a:solidFill>${createChartTextFonts(face)}</a:rPr>` +
		el('a:t', null, text) +
		'</a:r></a:p>' +
		'</c:rich></c:tx>' +
		'<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>' +
		'<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>'
	)
}

/**
 * Build an `<a:ln>` border element from a per-data-point `BorderProps`.
 * @param border - point border style (`type`, `color`, `pt`)
 */
export function createChartBorderLine(border: BorderProps): string {
	if (border.type === 'none') return '<a:ln><a:noFill/></a:ln>'
	const dash = border.type === 'dash' ? 'dash' : 'solid'
	return `<a:ln w="${ptsToEmuLenient(resolveBorderWidth(border, 1))}" cap="flat">${genXmlColorSelection({ color: border.color || '666666', transparency: border.transparency })}<a:prstDash val="${dash}"/><a:round/></a:ln>`
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
		const fillColor = ptStyle?.fill || (arrColors ? arrColors[index % arrColors.length] : null)
		const pattern = ptStyle?.pattern
		const border = ptStyle?.border
		// Nothing to style for this point -> omit the c:dPt entirely
		if (!fillColor && !pattern && !border) return

		xml += '<c:dPt>'
		xml += `<c:idx val="${index}"/>`
		if (isBar) xml += '<c:invertIfNegative val="0"/>'
		xml += '<c:bubble3D val="0"/>'
		xml += '<c:spPr>'
		if ((isBar || isScatter) && opts.lineSize === 0 && !border && !ptStyle?.fill && !pattern) {
			// Preserve legacy color-vary behavior: hide outline when lineSize===0
			xml += '<a:ln><a:noFill/></a:ln>'
		} else {
			// Pattern fill takes precedence over a solid fill (OOXML allows only one fill per c:dPt).
			// Default the pattern foreground to this point's resolved color so it reads as a hatched bar.
			if (pattern) {
				xml += genXmlPatternFill(fillColor && !pattern.fgColor ? { ...pattern, fgColor: fillColor } : pattern)
			} else if (fillColor) {
				// BAR3D color-vary historically tints the edge line, not the face fill
				if (chartType === ChartType.bar3d) xml += `<a:ln>${genXmlColorSelection(fillColor)}</a:ln>`
				else xml += genXmlColorSelection(fillColor)
			}
			if (border) xml += createChartBorderLine(border)
		}
		xml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)
		xml += '</c:spPr>'
		xml += '</c:dPt>'
	})
	return xml
}

/**
 * A `<c:xVal>`/`<c:yVal>` numeric-reference block: the `<c:f>` formula plus the `<c:numCache>`
 * that mirrors the cells it points at.
 *
 * Scatter and bubble each carried their own copy of this twice over, four near-identical
 * twelve-line blocks. `values` is the exact point list to cache, so the caller decides what a
 * gap is: the Y series is emitted against the X series' length (a caller may supply fewer Y
 * values than X — a timeline with only the first few months filled in), and the shorter array's
 * tail arrives here as `undefined`, which {@link numCachePt} skips.
 *
 * @param tag - the wrapping element, `c:xVal` or `c:yVal`
 * @param ref - the `<c:f>` formula, from {@link sheetRangeRef} or written inline
 * @param formatCode - the cached `<c:formatCode>`
 * @param values - the points to cache, in order; `null`/`undefined` entries are gaps
 * @param refIndent - leading whitespace before `<c:f>`. Bubble's `c:yVal` has none where every
 *   other block has four spaces. That is inert inter-element whitespace, but it is *emitted*
 *   whitespace, and this extraction is behaviour-preserving — normalizing it would be a byte
 *   change wearing a cleanup's clothes, which the byte-identity gate exists to refuse.
 */
export function numRefBlock(
	tag: 'c:xVal' | 'c:yVal',
	ref: string,
	formatCode: string,
	values: Array<number | null | undefined>,
	refIndent = '    '
): string {
	let xml = `<${tag}>`
	xml += '  <c:numRef>'
	xml += `${refIndent}<c:f>${ref}</c:f>`
	xml += '    <c:numCache>'
	xml += '      <c:formatCode>' + formatCode + '</c:formatCode>'
	xml += `      <c:ptCount val="${values.length}"/>`
	values.forEach((value, idx) => {
		xml += numCachePt(idx, value)
	})
	xml += '    </c:numCache>'
	xml += '  </c:numRef>'
	xml += `</${tag}>`
	return xml
}
