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
	// `sizeAttr` is empty when the caller set no font size — PowerPoint then picks the default —
	// and interpolating it empty leaves TWO spaces between the tag name and `b=`. Those are
	// emitted bytes, and `el()` writes exactly one space before an attribute by design, so the
	// two run-property tags below stay hand-written. Dropping the padding would be the real fix —
	// no XML consumer can see it — but that is an output change, and a whitespace-only diff is a
	// stop rather than a cleanup. Two other sites in this directory are in the same position; the
	// ratchet header lists them.
	const sizeAttr = opts.fontSize ? `sz="${ptToHundredths(opts.fontSize)}"` : ''
	const runAttrs = ` ${sizeAttr} b="${opts.titleBold ? 1 : 0}" i="${opts.titleItalic ? 1 : 0}" u="${opts.titleUnderline ? 'sng' : 'none'}" strike="noStrike">`
	const runChildren =
		'\n              ' +
		genXmlColorSelection(opts.color || DEF_FONT_COLOR) +
		'\n              ' +
		createChartTextFonts(opts.fontFace || 'Arial') +
		'\n            '

	// NOTE: manualLayout x/y vals are *relative to the entire slide*. Each axis is independent in
	// CT_ManualLayout: omitting xMode/x (or yMode/y) leaves that axis on automatic layout, so a
	// caller can center horizontally while still applying a manual vertical offset (and vice-versa).
	// Schema order is xMode, yMode, x, y.
	const hasX = opts.titlePos && typeof opts.titlePos.x === 'number'
	const hasY = opts.titlePos && typeof opts.titlePos.y === 'number'
	/** Fold a slide-relative offset into the fraction-of-chart value `c:x`/`c:y` want. */
	const edgeFraction = (offset: number): number => {
		let val = offset === 0 ? 0 : (offset * (offset / 5)) / 10
		if (val >= 1) val = val / 10
		if (val >= 0.1) val = val / 10
		return val
	}
	let layout = voidEl('c:layout')
	if (hasX || hasY) {
		const modes = (hasX ? voidEl('c:xMode', { val: 'edge' }) : '') + (hasY ? voidEl('c:yMode', { val: 'edge' }) : '')
		const vals =
			(hasX ? voidEl('c:x', { val: edgeFraction((opts.titlePos?.x ?? 0) + (chartX ?? 0)) }) : '') +
			(hasY ? voidEl('c:y', { val: edgeFraction((opts.titlePos?.y ?? 0) + (chartY ?? 0)) }) : '')
		layout = el('c:layout', null, raw(el('c:manualLayout', null, raw(modes + vals))))
	}

	const paragraph = el(
		'a:p',
		null,
		[
			raw(
				el(
					'a:pPr',
					opts.titleAlign === 'left' || opts.titleAlign === 'right' ? { algn: opts.titleAlign.slice(0, 1) } : null,
					raw('\n            <a:defRPr' + runAttrs + runChildren + '</a:defRPr>'),
					{ openPrefix: '\n            ', closePrefix: '\n          ' }
				)
			),
			raw(
				el(
					'a:r',
					null,
					[
						raw('\n            <a:rPr' + runAttrs + runChildren + '</a:rPr>'),
						raw('\n            ' + el('a:t', null, opts.title ?? '')),
					],
					{ openPrefix: '\n          ', closePrefix: '\n          ' }
				)
			),
		],
		{ openPrefix: '\n          ', closePrefix: '\n        ' }
	)
	const rich = el(
		'c:rich',
		null,
		[
			// Don't specify a rotation when none was asked for, so the default applies (which is
			// vertical on a category axis).
			raw(
				voidEl(
					'a:bodyPr',
					{ rot: opts.titleRotate ? convertAngleUnits(opts.titleRotate, 'titleRotate') : undefined },
					{ openPrefix: '\n          ' }
				)
			),
			raw(voidEl('a:lstStyle', null, { openPrefix: '\n          ' })),
			raw(paragraph),
		],
		{ openPrefix: '\n        ', closePrefix: '\n        ' }
	)
	return el(
		'c:title',
		null,
		[
			raw(el('c:tx', null, raw(rich), { openPrefix: '\n      ', closePrefix: '\n      ' })),
			raw('\n      ' + layout),
			raw(voidEl('c:overlay', { val: 0 }, { openPrefix: '\n      ' })),
		],
		{ closePrefix: '\n    ' }
	)
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
	const defRPr = el(
		'a:defRPr',
		{
			b: opts.dataLabelFontBold ? 1 : 0,
			i: opts.dataLabelFontItalic ? 1 : 0,
			strike: 'noStrike',
			sz: ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE),
			u: 'none',
		},
		[
			raw(genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)),
			raw('          ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')),
		],
		{ openPrefix: '        ', closePrefix: '        ' }
	)
	const txPr = el(
		'c:txPr',
		null,
		[
			raw(voidEl('a:bodyPr', null, { openPrefix: '      ' })),
			raw(voidEl('a:lstStyle', null, { openPrefix: '      ' })),
			raw(el('a:p', null, raw(el('a:pPr', null, raw(defRPr), { closePrefix: '      ' })), { openPrefix: '      ' })),
		],
		{ openPrefix: '    ', closePrefix: '    ' }
	)
	return el(
		'c:dLbls',
		null,
		[
			raw(
				voidEl(
					'c:numFmt',
					{ formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 },
					{ openPrefix: '    ' }
				)
			),
			raw(txPr),
			opts.dataLabelPosition ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition }, { openPrefix: ' ' })) : null,
			raw(voidEl('c:showLegendKey', { val: 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showVal', { val: opts.showValue ? 1 : 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showCatName', { val: 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showSerName', { val: opts.showSerName ? 1 : 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showPercent', { val: 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showBubbleSize', { val: 0 }, { openPrefix: '    ' })),
			leaderLines
				? raw(voidEl('c:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 }, { openPrefix: '    ' }))
				: null,
		],
		{ openPrefix: '  ', closePrefix: '  ' }
	)
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
export function dimmedTextLine(lumMod: number, lumOff: number, fmt?: { openPrefix?: string }): string {
	return el(
		'a:ln',
		{ w: 9525, cap: 'flat', cmpd: 'sng', algn: 'ctr' },
		[raw(dimmedTextFill(lumMod, lumOff)), raw(voidEl('a:round'))],
		fmt
	)
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
			voidEl('c:noEndCap', { val: eb.noEndCap ? 1 : 0 })

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

		if (eb.color || eb.size != null) {
			children += el(
				'c:spPr',
				null,
				raw(
					el(
						'a:ln',
						{ w: eb.size != null ? ptsToEmuLenient(eb.size) : undefined },
						raw(eb.color ? genXmlColorSelection(eb.color) : '')
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
	const sz = ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)
	const bold = opts.dataLabelFontBold ? 1 : 0
	const italic = opts.dataLabelFontItalic ? 1 : 0
	const color = createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)
	const face = opts.dataLabelFontFace || 'Arial'
	const lang = opts.lang || 'en-US'
	const runChildren = [raw(el('a:solidFill', null, raw(color))), raw(createChartTextFonts(face))]
	const fontAttrs = { sz, b: bold, i: italic, u: 'none', strike: 'noStrike' }
	const paragraph = el('a:p', null, [
		raw(el('a:pPr', null, raw(el('a:defRPr', fontAttrs, runChildren)))),
		raw(el('a:r', null, [raw(el('a:rPr', { lang, ...fontAttrs, dirty: 0 }, runChildren)), raw(el('a:t', null, text))])),
	])
	const rich = el('c:rich', null, [raw(voidEl('a:bodyPr')), raw(voidEl('a:lstStyle')), raw(paragraph)])
	return el('c:dLbl', null, [
		raw(voidEl('c:idx', { val: idx })),
		raw(el('c:tx', null, raw(rich))),
		raw(voidEl('c:showLegendKey', { val: 0 })),
		raw(voidEl('c:showVal', { val: 0 })),
		raw(voidEl('c:showCatName', { val: 0 })),
		raw(voidEl('c:showSerName', { val: 0 })),
		raw(voidEl('c:showPercent', { val: 0 })),
		raw(voidEl('c:showBubbleSize', { val: 0 })),
	])
}

/**
 * Build an `<a:ln>` border element from a per-data-point `BorderProps`.
 * @param border - point border style (`type`, `color`, `pt`)
 */
export function createChartBorderLine(border: BorderProps): string {
	if (border.type === 'none') return el('a:ln', null, raw(voidEl('a:noFill')))
	return el('a:ln', { w: ptsToEmuLenient(resolveBorderWidth(border, 1)), cap: 'flat' }, [
		raw(genXmlColorSelection({ color: border.color || '666666', transparency: border.transparency })),
		raw(voidEl('a:prstDash', { val: border.type === 'dash' ? 'dash' : 'solid' })),
		raw(voidEl('a:round')),
	])
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

		let shape = ''
		if ((isBar || isScatter) && opts.lineSize === 0 && !border && !ptStyle?.fill && !pattern) {
			// Preserve legacy color-vary behavior: hide outline when lineSize===0
			shape = el('a:ln', null, raw(voidEl('a:noFill')))
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
 * The whitespace spelling a {@link strRefBlock} is emitted with.
 *
 * The five hand-written copies this replaces produced the same elements three different ways —
 * `plot-scatter`, `plot-bubble` and `plot-cat-axis` agreed on one indentation, `plot-surface` ran
 * everything together with none, and `plot-pie` broke the cache across four lines. All three are
 * inert to PowerPoint, and all three are *emitted* bytes, so this extraction has to preserve which
 * one each call site produced. That is what this parameter is for, and it is exactly what a
 * whitespace normalization would erase — which is why that is its own deliberate, re-baselined
 * piece of work and this is not.
 */
export type StrRefLayout = 'indented' | 'compact' | 'expanded'

/**
 * A `<c:tx>` series-name block: the `<c:f>` formula plus the one-point `<c:strCache>` mirroring
 * the header cell it points at. The string counterpart of {@link numRefBlock}.
 *
 * Built through `el()`/`voidEl()` rather than by concatenation, with the indentation carried on
 * the `openPrefix`/`childPrefix`/`closePrefix` byte-layout hooks those helpers already have. The
 * five copies this replaces were all hand-built strings, so routing them here takes fifty-odd
 * hand-written delimiters out of `src/gen/chart/` rather than moving them into one file.
 * @param ref - the `<c:f>` formula, from {@link sheetCellRef} or written inline
 * @param name - the series name to cache; escaped on the way into `<c:v>`
 * @param layout - which of the three historical whitespace spellings to emit; see {@link StrRefLayout}
 */
export function strRefBlock(ref: string, name: string, layout: StrRefLayout = 'indented'): string {
	const compact = layout === 'compact'
	const pt = el('c:pt', { idx: 0 }, raw(el('c:v', null, name)))
	const cacheChildren = [raw(voidEl('c:ptCount', { val: 1 })), raw(pt)]
	const strCache =
		layout === 'expanded'
			? el('c:strCache', null, cacheChildren, { childPrefix: '        ', closePrefix: '      ' })
			: el('c:strCache', null, cacheChildren)
	const indent = (spaces: string): string => (compact ? '' : spaces)
	const strRef = el('c:strRef', null, [raw(el('c:f', null, ref)), raw(strCache)], {
		openPrefix: indent('    '),
		childPrefix: indent('      '),
		closePrefix: indent('    '),
	})
	return el('c:tx', null, raw(strRef), { openPrefix: indent('  '), closePrefix: indent('  ') })
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
 * `layout` preserves which whitespace spelling each site emitted, for the reason
 * {@link StrRefLayout} gives. The `<c:multiLvlStrRef>` arm of a cat-axis is deliberately not
 * folded in: multi-level categories nest a `<c:lvl>` per label group, which is a different shape
 * rather than a different spelling.
 * @param kind - `num` for a numeric (date) category cache, `str` for a text one
 * @param ref - the `<c:f>` formula
 * @param labels - the category labels, in order; escaped on the way into `<c:v>`
 * @param formatCode - the cached `<c:formatCode>`; numeric caches only
 * @param layout - which whitespace spelling to emit
 */
export function catRefBlock(
	kind: 'num' | 'str',
	ref: string,
	labels: string[],
	formatCode?: string,
	layout: 'compact' | 'indented' = 'compact'
): string {
	const indent = (spaces: string): string => (layout === 'indented' ? spaces : '')
	// `<c:ptCount>` and the points it counts are ONE child: the indented sites put whitespace
	// before the count and none before each point, so splitting them would indent every point.
	const points = labels.map((label, idx) => el('c:pt', { idx }, raw(el('c:v', null, label)))).join('')
	const cache = el(
		`c:${kind}Cache`,
		null,
		[
			formatCode === undefined ? null : raw(el('c:formatCode', null, formatCode)),
			raw(voidEl('c:ptCount', { val: labels.length }) + points),
		],
		{ childPrefix: indent('      '), closePrefix: indent('    ') }
	)
	return el(`c:${kind}Ref`, null, [raw(el('c:f', null, ref)), raw(cache)], {
		openPrefix: indent('  '),
		childPrefix: indent('    '),
		closePrefix: indent('  '),
	})
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
	const numCache = el(
		'c:numCache',
		null,
		[
			raw(el('c:formatCode', null, formatCode, { openPrefix: '      ' })),
			raw(voidEl('c:ptCount', { val: values.length }, { openPrefix: '      ' })),
			raw(values.map((value, idx) => numCachePt(idx, value)).join('')),
		],
		{ openPrefix: '    ', closePrefix: '    ' }
	)
	const numRef = el('c:numRef', null, [raw(el('c:f', null, ref, { openPrefix: refIndent })), raw(numCache)], {
		openPrefix: '  ',
		closePrefix: '  ',
	})
	return el(tag, null, raw(numRef))
}
