/**
 * Read-only proxies for a chart (`c:chart`) hosted in a `p:graphicFrame`.
 *
 * A chart lives in its own part (`/ppt/charts/chartN.xml`), referenced from the
 * graphic frame by relationship id. `Chart` reads the chart type, title, series,
 * and the **cached** category/value data (`c:numCache` / `c:strCache`) that
 * PowerPoint stores alongside the embedded workbook. Editing the underlying
 * workbook is out of scope; this surface is read-only.
 */
import type { Part } from '../opc/part.js'
import {
	attr,
	boolValue,
	concatDrawingMLText,
	type Element,
	ELEMENT_NODE,
	firstChild,
	firstChildElement,
	getElements,
	numberValue,
	OOXML_NS,
} from '../oxml/dom.js'
import { readIndexedPoints } from '../oxml/point-cache.js'
import { readLineBasics } from './line.js'
import { readColorRef, type ColorRef } from './theme-context.js'

/** A chart axis number format (`c:numFmt`). */
export interface AxisNumberFormat {
	/** The format mask (`@formatCode`), e.g. `"General"`, `"0.0%"`. */
	formatCode: string | null
	/** Whether the mask is linked to the source cell format (`@sourceLinked`). */
	sourceLinked: boolean | null
}

/** A chart legend (`c:legend`). */
export interface ChartLegend {
	/** Legend position (`c:legendPos/@val`): `r`/`l`/`t`/`b`/`tr`. */
	position: string | null
	/** Whether the legend overlays the plot area (`c:overlay/@val`). */
	overlay: boolean | null
}

/** Data-label settings (`c:dLbls`), read from a plot group or from one series. */
export interface ChartDataLabels {
	/** Show the point value (`c:showVal`). */
	showValue: boolean | null
	/** Show the series name (`c:showSerName`). */
	showSeriesName: boolean | null
	/** Show the category name (`c:showCatName`). */
	showCategoryName: boolean | null
	/** Show the percentage (`c:showPercent`). */
	showPercent: boolean | null
	/** Show the legend key swatch (`c:showLegendKey`). */
	showLegendKey: boolean | null
	/** Show the bubble size (`c:showBubbleSize`). */
	showBubbleSize: boolean | null
	/** Show leader lines (`c:showLeaderLines`). */
	showLeaderLines: boolean | null
	/** Label position (`c:dLblPos/@val`), or `null` when auto. */
	position: string | null
	/** Label number format (`c:numFmt`), or `null` when absent. */
	numberFormat: AxisNumberFormat | null
}

/** A series' solid fill (`c:ser/c:spPr`). */
export interface ChartFill {
	/**
	 * The fill's solid colour; every field is `null` for a non-solid fill. A chart part is read
	 * without a theme, so `resolved` is always `null` and a scheme colour stays a token.
	 */
	colorRef: ColorRef
	/** Whether the fill is explicitly suppressed (`a:noFill`). */
	noFill: boolean
}

/** A series' line/stroke (`c:ser/c:spPr/a:ln`). */
export interface ChartLine {
	/** Stroke width in points (`@w`/12700), or `null` when unset. */
	widthPt: number | null
	/** Dash style (`a:prstDash/@val`), or `null` when solid/unset. */
	dash: string | null
	/** The stroke's solid colour; unresolved, as {@link ChartFill.colorRef} is. */
	colorRef: ColorRef
	/** Whether the line is explicitly suppressed (`a:noFill`). */
	noFill: boolean
}

/** A chart graphic-frame's chart, backed by its chart part. */
export class Chart {
	constructor(
		/** The chart's OPC part (`/ppt/charts/chartN.xml`). */
		readonly part: Part
	) {}

	/** Partname of the chart part. */
	get partName(): string {
		return this.part.partName
	}

	/**
	 * The chart-group type tokens present in the plot area, in document order
	 * (e.g. `['line']`, `['bar']`, `['bar', 'line']` for a combo chart). Derived
	 * from the plot-area element local names (`lineChart` → `line`).
	 */
	get chartTypes(): string[] {
		return this.#chartGroups().map((group) => (group.localName ?? '').replace(/Chart$/, ''))
	}

	/** The first chart-group type, or `null` for an empty plot area. */
	get chartType(): string | null {
		return this.chartTypes[0] ?? null
	}

	/** The chart title (`c:chart/c:title` rich text), or `null` when absent/auto. */
	get title(): string | null {
		const chart = this.#chart()
		return readTitleText(chart && firstChild(chart, 'c:title'))
	}

	/**
	 * The category/value/series axes (`c:catAx`/`c:valAx`/`c:dateAx`/`c:serAx`)
	 * in plot-area document order. Empty for pie/doughnut charts (no axes).
	 */
	get axes(): ChartAxis[] {
		const plotArea = this.#plotArea()
		if (!plotArea) return []
		const out: ChartAxis[] = []
		for (let node = plotArea.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const element = node as Element
			if (element.namespaceURI === OOXML_NS.c && (element.localName ?? '').endsWith('Ax'))
				out.push(new ChartAxis(element, this.part))
		}
		return out
	}

	/** The category axis (`c:catAx`/`c:dateAx`), or `null` when the chart has none. */
	get categoryAxis(): ChartAxis | null {
		return this.axes.find((axis) => axis.kind === 'cat' || axis.kind === 'date') ?? null
	}

	/** The (primary) value axis (`c:valAx`), or `null` when the chart has none. */
	get valueAxis(): ChartAxis | null {
		return this.axes.find((axis) => axis.kind === 'val') ?? null
	}

	/** The legend (`c:chart/c:legend`) position + overlay, or `null` when hidden. */
	get legend(): ChartLegend | null {
		const chart = this.#chart()
		const legend = chart && firstChild(chart, 'c:legend')
		if (!legend) return null
		const pos = firstChild(legend, 'c:legendPos')
		const overlay = firstChild(legend, 'c:overlay')
		return {
			position: pos ? attr(pos, 'val') : null,
			overlay: overlay ? boolValue(attr(overlay, 'val')) : null,
		}
	}

	/**
	 * The chart-level data-label settings, read from the first plot group's
	 * aggregate `c:dLbls` (the group-wide block after the series). `null` when the
	 * plot group carries no data-label block.
	 *
	 * A pie keeps its labels on the series instead: PowerPoint writes the flags the
	 * user set into `c:ser/c:dLbls` and leaves this block all off, and this library's
	 * writer emits no group block for a pie at all. Read those through
	 * {@link ChartSeries.dataLabels}.
	 */
	get dataLabels(): ChartDataLabels | null {
		const group = this.#chartGroups()[0]
		return readDataLabels(group ? firstChild(group, 'c:dLbls') : null)
	}

	/** The data series (`c:ser`) across all chart groups, in document order. */
	get series(): ChartSeries[] {
		return this.#chartGroups()
			.flatMap((group) => getElements(group, 'c:ser'))
			.map((ser) => new ChartSeries(ser, this.part))
	}

	/**
	 * Category labels, read from the first series' cached categories
	 * (`c:cat`), as written. Empty when the chart has no category axis.
	 */
	get categories(): (string | null)[] {
		const firstSer = this.series[0]
		return firstSer ? firstSer.categories : []
	}

	/** Every category level of the first series, leaf first; see {@link ChartSeries.categoryLevels}. */
	get categoryLevels(): (string | null)[][] {
		const firstSer = this.series[0]
		return firstSer ? firstSer.categoryLevels : []
	}

	/** Escape hatch: the underlying `c:chartSpace` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element | null {
		return this.part.dom.documentElement
	}

	/** Mark the chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	#chart(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'c:chart') : null
	}

	#plotArea(): Element | null {
		const chart = this.#chart()
		return chart ? firstChild(chart, 'c:plotArea') : null
	}

	/** Plot-area children that are chart groups (local name ends with `Chart`). */
	#chartGroups(): Element[] {
		const plotArea = this.#plotArea()
		if (!plotArea) return []
		const groups: Element[] = []
		for (let node = plotArea.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== 1) continue
			const element = node as Element
			if (element.namespaceURI === OOXML_NS.c && element.localName?.endsWith('Chart')) groups.push(element)
		}
		return groups
	}
}

/**
 * One axis (`c:catAx`/`c:valAx`/`c:dateAx`/`c:serAx`) of a chart's plot area.
 * All four share the scaling/delete/position/tickmark structure; `kind`
 * distinguishes them.
 */
export class ChartAxis {
	constructor(
		private readonly ax: Element,
		/** The owning chart's OPC part, so {@link markDirty} can reach it from {@link element_}. */
		private readonly part: Part
	) {}

	/** Axis kind derived from the element name: `cat`/`val`/`date`/`ser`. */
	get kind(): 'cat' | 'val' | 'date' | 'ser' | null {
		switch (this.ax.localName) {
			case 'catAx':
				return 'cat'
			case 'valAx':
				return 'val'
			case 'dateAx':
				return 'date'
			case 'serAx':
				return 'ser'
			default:
				return null
		}
	}

	/** Axis id (`c:axId/@val`). */
	get id(): number | null {
		const el = firstChild(this.ax, 'c:axId')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Scaling orientation (`c:scaling/c:orientation/@val`): `minMax`/`maxMin`. */
	get orientation(): string | null {
		const orient = this.#scaling('c:orientation')
		return orient ? attr(orient, 'val') : null
	}

	/** Scale minimum (`c:scaling/c:min/@val`), or `null` when auto. */
	get min(): number | null {
		const el = this.#scaling('c:min')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Scale maximum (`c:scaling/c:max/@val`), or `null` when auto. */
	get max(): number | null {
		const el = this.#scaling('c:max')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Logarithmic scale base (`c:scaling/c:logBase/@val`), or `null` when linear. */
	get logBase(): number | null {
		const el = this.#scaling('c:logBase')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Whether the axis is hidden (`c:delete/@val` = 1). */
	get hidden(): boolean {
		const el = firstChild(this.ax, 'c:delete')
		return el ? boolValue(attr(el, 'val')) === true : false
	}

	/** Axis position (`c:axPos/@val`): `b`/`l`/`r`/`t`. */
	get position(): string | null {
		const el = firstChild(this.ax, 'c:axPos')
		return el ? attr(el, 'val') : null
	}

	/** Whether the axis draws major gridlines (`c:majorGridlines`). */
	get majorGridlines(): boolean {
		return !!firstChild(this.ax, 'c:majorGridlines')
	}

	/** Whether the axis draws minor gridlines (`c:minorGridlines`). */
	get minorGridlines(): boolean {
		return !!firstChild(this.ax, 'c:minorGridlines')
	}

	/** Axis title text (`c:title` rich text), or `null` when absent. */
	get title(): string | null {
		return readTitleText(firstChild(this.ax, 'c:title'))
	}

	/** Axis number format (`c:numFmt`), or `null` when absent. */
	get numberFormat(): AxisNumberFormat | null {
		return readNumberFormat(this.ax)
	}

	/** Major tick-mark style (`c:majorTickMark/@val`): `out`/`in`/`cross`/`none`. */
	get majorTickMark(): string | null {
		const el = firstChild(this.ax, 'c:majorTickMark')
		return el ? attr(el, 'val') : null
	}

	/** Minor tick-mark style (`c:minorTickMark/@val`). */
	get minorTickMark(): string | null {
		const el = firstChild(this.ax, 'c:minorTickMark')
		return el ? attr(el, 'val') : null
	}

	/** Tick-label position (`c:tickLblPos/@val`): `nextTo`/`high`/`low`/`none`. */
	get tickLabelPosition(): string | null {
		const el = firstChild(this.ax, 'c:tickLblPos')
		return el ? attr(el, 'val') : null
	}

	/** Major unit (`c:majorUnit/@val`), or `null` when auto. */
	get majorUnit(): number | null {
		const el = firstChild(this.ax, 'c:majorUnit')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Minor unit (`c:minorUnit/@val`), or `null` when auto. */
	get minorUnit(): number | null {
		const el = firstChild(this.ax, 'c:minorUnit')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Escape hatch: the underlying axis element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ax
	}

	/** Mark the owning chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/** A named child of `c:scaling`. */
	#scaling(qname: string): Element | null {
		const scaling = firstChild(this.ax, 'c:scaling')
		return scaling ? firstChild(scaling, qname) : null
	}
}

/** One data series (`c:ser`) of a chart. */
export class ChartSeries {
	constructor(
		private readonly ser: Element,
		/** The owning chart's OPC part, so {@link markDirty} can reach it from {@link element_}. */
		private readonly part: Part
	) {}

	/**
	 * Series fill (`c:spPr` solid fill / no-fill). `null` when the series carries
	 * no `c:spPr` or its `c:spPr` declares no fill choice (inherits the theme).
	 */
	get fill(): ChartFill | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		return spPr ? readSolidFill(spPr) : null
	}

	/**
	 * Series line/stroke (`c:spPr/a:ln`) — width, dash, colour. `null` when the
	 * series `c:spPr` has no `a:ln` (bar/area fills carry no line by default).
	 */
	get line(): ChartLine | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		const ln = spPr && firstChild(spPr, 'a:ln')
		if (!ln) return null
		// The chart part is read without a theme, so the colour stays unresolved.
		const { widthPt, dash, colorRef, noFill } = readLineBasics(ln, null)
		return { widthPt, dash, colorRef, noFill }
	}

	/** Series index (`c:idx/@val`), or `null` if absent. */
	get index(): number | null {
		const idx = firstChild(this.ser, 'c:idx')
		return idx ? numberValue(attr(idx, 'val')) : null
	}

	/** Series name from the cached `c:tx`, or `null` when unnamed. */
	get name(): string | null {
		const tx = firstChild(this.ser, 'c:tx')
		if (!tx) return null
		const direct = firstChild(tx, 'c:v')
		if (direct) return direct.textContent ?? null
		const points = readPoints(findCache(tx))
		return points[0] ?? null
	}

	/**
	 * Cached numeric values (`c:val`); non-numeric or missing points are `null`. Empty on a
	 * scatter or bubble series, which caches its values as {@link yValues} instead.
	 */
	get values(): (number | null)[] {
		return this.#numbers('c:val')
	}

	/**
	 * Cached X values of a scatter or bubble series (`c:xVal`); missing points are `null`.
	 * A series plotted against text X labels reads `null` at every point, a label that looks like a
	 * number included: {@link xLabels} has the text. A series with no `c:xVal` reads `[]`, and
	 * PowerPoint plots it at X = 1, 2, … n, as it does a series plotted against text.
	 */
	get xValues(): (number | null)[] {
		const cache = findCache(firstChild(this.ser, 'c:xVal'))
		if (cache && isTextCache(cache)) return textPoints(cache).map(() => null)
		return readPoints(cache).map(numberValue)
	}

	/**
	 * The X labels of a series plotted against text (`c:xVal` holding a string cache rather than
	 * numbers), as written, or `null` when its X values are numbers or absent.
	 *
	 * One text cell makes the whole X column a string cache, so the numbers in it arrive as text too.
	 * PowerPoint then plots the points at X = 1, 2, … n in point order and uses none of the labels
	 * as a coordinate, not even one that reads as a number.
	 */
	get xLabels(): (string | null)[] | null {
		const cache = findCache(firstChild(this.ser, 'c:xVal'))
		return cache && isTextCache(cache) ? textPoints(cache) : null
	}

	/** Cached Y values of a scatter or bubble series (`c:yVal`); non-numeric or missing points are `null`. */
	get yValues(): (number | null)[] {
		return this.#numbers('c:yVal')
	}

	/** Cached bubble sizes of a bubble series (`c:bubbleSize`); non-numeric or missing points are `null`. */
	get bubbleSizes(): (number | null)[] {
		return this.#numbers('c:bubbleSize')
	}

	/**
	 * Cached category labels for this series (`c:cat`), as written. On a multi-level axis this
	 * is the leaf level, the labels next to the plot; {@link categoryLevels} has the rest.
	 */
	get categories(): (string | null)[] {
		return this.categoryLevels[0] ?? []
	}

	/**
	 * Every level of cached category labels (`c:cat`), leaf first: the order PowerPoint writes the
	 * `c:multiLvlStrCache/c:lvl` children in, and the order `OptsChartData.labels` takes them.
	 * Each level is as long as the leaf level. An outer level names a group once, at the group's
	 * first category, so the other slots of that group are `null`. A single-level axis reads as
	 * one level, and a series with no categories as none.
	 */
	get categoryLevels(): (string | null)[][] {
		const cat = firstChild(this.ser, 'c:cat')
		const cache = cat && findCache(cat)
		if (!cache) return []
		if (cache.localName !== 'multiLvlStrCache') return [readPoints(cache)]
		return readCategoryLevels(cache)
	}

	/**
	 * The data-label settings this series carries itself (`c:ser/c:dLbls`), or `null` when it has
	 * none. This is where a pie's labels are: see {@link Chart.dataLabels}.
	 */
	get dataLabels(): ChartDataLabels | null {
		return readDataLabels(firstChild(this.ser, 'c:dLbls'))
	}

	/** Escape hatch: the underlying `c:ser` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ser
	}

	/** Mark the owning chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/** The cached points of a numeric child of the series (`c:val`, `c:xVal`, …), as numbers. */
	#numbers(qname: string): (number | null)[] {
		const container = firstChild(this.ser, qname)
		return readPoints(container && findCache(container)).map(numberValue)
	}
}

/** Read a `c:dLbls` block's show flags, position and number format, or `null` when there is no block. */
function readDataLabels(dLbls: Element | null): ChartDataLabels | null {
	if (!dLbls) return null
	const flag = (qname: string): boolean | null => {
		const el = firstChild(dLbls, qname)
		return el ? boolValue(attr(el, 'val')) : null
	}
	const pos = firstChild(dLbls, 'c:dLblPos')
	return {
		showValue: flag('c:showVal'),
		showSeriesName: flag('c:showSerName'),
		showCategoryName: flag('c:showCatName'),
		showPercent: flag('c:showPercent'),
		showLegendKey: flag('c:showLegendKey'),
		showBubbleSize: flag('c:showBubbleSize'),
		showLeaderLines: flag('c:showLeaderLines'),
		position: pos ? attr(pos, 'val') : null,
		numberFormat: readNumberFormat(dLbls),
	}
}

/**
 * Read a `c:multiLvlStrCache` into its levels, in document order.
 *
 * The cache's `c:ptCount` counts categories, which is the leaf level's length, so only the leaf
 * is checked against it. An outer level is sparse by design, one point per group, and checking
 * it would warn about every multi-level chart PowerPoint writes. Every level is then widened to
 * the longest, so `levels[n][i]` is category `i`'s label at level `n`.
 */
function readCategoryLevels(cache: Element): (string | null)[][] {
	const ptCount = firstChild(cache, 'c:ptCount')
	const declared = ptCount ? numberValue(attr(ptCount, 'val')) : null
	const levels = getElements(cache, 'c:lvl').map((lvl, depth) =>
		readIndexedPoints(
			getElements(lvl, 'c:pt'),
			depth === 0 ? declared : null,
			(pt) => firstChild(pt, 'c:v')?.textContent ?? null,
			'c:multiLvlStrCache/c:ptCount'
		)
	)
	const width = Math.max(0, ...levels.map((level) => level.length))
	for (const level of levels) while (level.length < width) level.push(null)
	return levels
}

/** Concatenate the rich-text runs of a `c:title` element, or `null` when absent/empty. */
function readTitleText(titleEl: Element | null): string | null {
	if (!titleEl) return null
	const tx = firstChild(titleEl, 'c:tx')
	return concatDrawingMLText(tx && firstChild(tx, 'c:rich'))
}

/** Read a `c:numFmt` child of `parent` into `{ formatCode, sourceLinked }`, or `null` when absent. */
function readNumberFormat(parent: Element): AxisNumberFormat | null {
	const nf = firstChild(parent, 'c:numFmt')
	if (!nf) return null
	return { formatCode: attr(nf, 'formatCode'), sourceLinked: boolValue(attr(nf, 'sourceLinked')) }
}

/**
 * The solid fill (or explicit `a:noFill`) a series' `c:spPr` declares, or `null` when it declares
 * neither and the series inherits the theme. The chart part carries no theme context, so the
 * colour is left unresolved and a scheme token surfaces raw rather than flattened to hex.
 *
 * Whether a fill is *there* is decided by the element, not by the colour inside it. The caller
 * used to ask whether the decoded `colorRef` carried an `srgb` or a `scheme` value -- two of the
 * three literal forms {@link readColorRef} reads, and none of the forms it does not, such as
 * `a:sysClr`. So a series filled with an `a:prstClr` reported no fill at all, and re-authoring it
 * silently dropped a fill the file states.
 */
function readSolidFill(container: Element): ChartFill | null {
	const solidFill = firstChild(container, 'a:solidFill')
	const noFill = !!firstChild(container, 'a:noFill')
	if (!solidFill && !noFill) return null
	return { colorRef: readColorRef(solidFill ? firstChildElement(solidFill) : null, null), noFill }
}

/** Resolve a `c:cat`/`c:val`/`c:tx` container to its cache element (`c:numCache`/`c:strCache`/literal). */
function findCache(container: Element | null): Element | null {
	if (!container) return null
	for (const refName of ['c:numRef', 'c:strRef', 'c:multiLvlStrRef']) {
		const ref = firstChild(container, refName)
		if (ref) {
			for (const cacheName of ['c:numCache', 'c:strCache', 'c:multiLvlStrCache']) {
				const cache = firstChild(ref, cacheName)
				if (cache) return cache
			}
		}
	}
	// Inline literals (no workbook reference).
	return firstChild(container, 'c:numLit') ?? firstChild(container, 'c:strLit')
}

/**
 * Whether a cache holds text rather than numbers. `c:xVal` takes any of the data-source forms a
 * `c:cat` does, so a scatter's X values can arrive as a string cache, a string literal or even a
 * multi-level cache.
 */
function isTextCache(cache: Element): boolean {
	return cache.localName !== 'numCache' && cache.localName !== 'numLit'
}

/** A text cache's labels, idx-ordered: the leaf level of a multi-level cache. */
function textPoints(cache: Element): (string | null)[] {
	return cache.localName === 'multiLvlStrCache' ? (readCategoryLevels(cache)[0] ?? []) : readPoints(cache)
}

/** Read a cache's points (`c:pt[@idx]/c:v`) into an idx-ordered array; `c:ptCount` is the declared count. */
function readPoints(cache: Element | null): (string | null)[] {
	if (!cache) return []
	const ptCount = firstChild(cache, 'c:ptCount')
	return readIndexedPoints(
		getElements(cache, 'c:pt'),
		ptCount ? numberValue(attr(ptCount, 'val')) : null,
		(pt) => firstChild(pt, 'c:v')?.textContent ?? null,
		'c:ptCount'
	)
}
