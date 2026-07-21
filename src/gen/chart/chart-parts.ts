/**
 * PptxGenJS: Shared Chart Fragment Builders
 *
 * The leaf builders every chart region reuses -- titles, gridlines, series data points,
 * error bars, number caches, leader lines, borders. Each is a pure string builder with
 * no dependency on any other module in this directory, which is what lets the plot
 * builders ({@link ./plot-cat-axis}, {@link ./plot-scatter}, {@link ./plot-bubble},
 * {@link ./plot-pie}), the axes ({@link ./chart-axes}), and the chart envelope
 * ({@link ./chart-xml}) all draw on them without a cycle.
 */

import { ChartType } from '../../core-enums.js'
import {
	BARCHART_COLORS,
	DEF_CHART_GRIDLINE,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SHAPE_SHADOW,
} from '../../core-enums-internal.js'
import type { BorderProps, ChartErrorBarOptions, ChartPropsTitle, OptsChartGridLine } from '../../core-interfaces.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { warn } from '../../log.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection, genXmlPatternFill } from '../drawingml/fill.js'
import { createLineCap, resolveBorderWidth } from '../drawingml/line.js'
import { convertRotationDegrees, valToPts } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { dataValues } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'

export const VALID_CHART_TIME_UNITS = ['days', 'months', 'years']

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
	const rotate = opts.titleRotate ? `<a:bodyPr rot="${convertRotationDegrees(opts.titleRotate)}"/>` : '<a:bodyPr/>' // don't specify rotation to get default (ex. vertical for cat axis)
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
	// NOTE: emits a bare `<a:srgbClr>` rather than going through `createColorElement`,
	// so a scheme color name is not resolved here -- tracked as
	// `fork-chart-gridline-scheme-color` in docs/backlog.yml.
	const line = el(
		'a:ln',
		{
			w: valToPts(glOpts.size || DEF_CHART_GRIDLINE.size || 1),
			cap: createLineCap(glOpts.cap || DEF_CHART_GRIDLINE.cap),
		},
		[
			raw('  ' + el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: glOpts.color || DEF_CHART_GRIDLINE.color })))),
			raw('   ' + voidEl('a:prstDash', { val: glOpts.style || DEF_CHART_GRIDLINE.style }) + voidEl('a:round')),
		],
		{ openPrefix: '  ', closePrefix: '  ' }
	)

	return el('c:majorGridlines', null, raw(el('c:spPr', null, raw(line), { openPrefix: ' ', closePrefix: ' ' })))
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
		warn(`chart value "${value}" at index ${idx} is not a finite number; data point omitted.`)
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
					warn(`chart series "${obj.name}" errorBars valueType 'cust' needs \`plusValues\` for barType '${barType}'.`)
				strXml += makeErrBarNumLit('plus', eb.plusValues || [])
			}
			if (barType !== 'plus') {
				if (!eb.minusValues?.length)
					warn(`chart series "${obj.name}" errorBars valueType 'cust' needs \`minusValues\` for barType '${barType}'.`)
				strXml += makeErrBarNumLit('minus', eb.minusValues || [])
			}
		} else if (valueType !== 'stdErr') {
			// fixedVal / percentage / stdDev use a single magnitude (stdErr derives it from the data).
			strXml += `<c:val val="${eb.value ?? 1}"/>`
		}

		if (eb.color || eb.size != null) {
			strXml += '<c:spPr><a:ln'
			strXml += eb.size != null ? ` w="${valToPts(eb.size)}"` : ''
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
		{ w: valToPts(opt.size || DEF_CHART_GRIDLINE.size || 1), cap: createLineCap(opt.cap || DEF_CHART_GRIDLINE.cap) },
		[
			raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: opt.color || DEF_CHART_GRIDLINE.color })))),
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
	const w = valToPts(opts.leaderLineSize ?? 0.75)
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
	return `<a:ln w="${valToPts(resolveBorderWidth(border, 1))}" cap="flat">${genXmlColorSelection({ color: border.color || '666666', transparency: border.transparency })}<a:prstDash val="${dash}"/><a:round/></a:ln>`
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
