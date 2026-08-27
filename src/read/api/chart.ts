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
	ELEMENT_NODE,
	OOXML_NS,
	attr,
	boolValue,
	firstChild,
	getElements,
	intValue,
	type Element,
} from '../oxml/dom.js'
import { readIndexedPoints } from '../oxml/point-cache.js'

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

/** A chart's data-label settings (`c:dLbls`), read from the first plot group. */
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
	/** Literal fill colour (`a:srgbClr/@val`), or `null` for a scheme/no fill. */
	color: string | null
	/** Raw scheme-colour token (`a:schemeClr/@val`), unresolved. */
	schemeColor: string | null
	/** Whether the fill is explicitly suppressed (`a:noFill`). */
	noFill: boolean
}

/** A series' line/stroke (`c:ser/c:spPr/a:ln`). */
export interface ChartLine {
	/** Stroke width in points (`@w`/12700), or `null` when unset. */
	widthPt: number | null
	/** Dash style (`a:prstDash/@val`), or `null` when solid/unset. */
	dash: string | null
	/** Literal stroke colour (`a:srgbClr/@val`), or `null`. */
	color: string | null
	/** Raw scheme-colour token (`a:schemeClr/@val`), unresolved. */
	schemeColor: string | null
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
	 * plot group carries no data-label block (e.g. pie via a different builder).
	 */
	get dataLabels(): ChartDataLabels | null {
		const group = this.#chartGroups()[0]
		const dLbls = group && firstChild(group, 'c:dLbls')
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
		return el ? intValue(attr(el, 'val')) : null
	}

	/** Scaling orientation (`c:scaling/c:orientation/@val`): `minMax`/`maxMin`. */
	get orientation(): string | null {
		const orient = this.#scaling('c:orientation')
		return orient ? attr(orient, 'val') : null
	}

	/** Scale minimum (`c:scaling/c:min/@val`), or `null` when auto. */
	get min(): number | null {
		const el = this.#scaling('c:min')
		return el ? intValue(attr(el, 'val')) : null
	}

	/** Scale maximum (`c:scaling/c:max/@val`), or `null` when auto. */
	get max(): number | null {
		const el = this.#scaling('c:max')
		return el ? intValue(attr(el, 'val')) : null
	}

	/** Logarithmic scale base (`c:scaling/c:logBase/@val`), or `null` when linear. */
	get logBase(): number | null {
		const el = this.#scaling('c:logBase')
		return el ? intValue(attr(el, 'val')) : null
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
		return el ? intValue(attr(el, 'val')) : null
	}

	/** Minor unit (`c:minorUnit/@val`), or `null` when auto. */
	get minorUnit(): number | null {
		const el = firstChild(this.ax, 'c:minorUnit')
		return el ? intValue(attr(el, 'val')) : null
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
		if (!spPr) return null
		const solid = readSolid(spPr)
		if (solid.color === null && solid.schemeColor === null && !solid.noFill) return null
		return solid
	}

	/**
	 * Series line/stroke (`c:spPr/a:ln`) — width, dash, colour. `null` when the
	 * series `c:spPr` has no `a:ln` (bar/area fills carry no line by default).
	 */
	get line(): ChartLine | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		const ln = spPr && firstChild(spPr, 'a:ln')
		if (!ln) return null
		const w = intValue(attr(ln, 'w'))
		const dash = firstChild(ln, 'a:prstDash')
		const solid = readSolid(ln)
		return {
			widthPt: w === null ? null : w / 12700,
			dash: dash ? (attr(dash, 'val') ?? null) : null,
			color: solid.color,
			schemeColor: solid.schemeColor,
			noFill: solid.noFill,
		}
	}

	/** Series index (`c:idx/@val`), or `null` if absent. */
	get index(): number | null {
		const idx = firstChild(this.ser, 'c:idx')
		return idx ? intValue(attr(idx, 'val')) : null
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

	/** Cached numeric values (`c:val`); non-numeric or missing points are `null`. */
	get values(): (number | null)[] {
		const val = firstChild(this.ser, 'c:val')
		return readPoints(val && findCache(val)).map(intValue)
	}

	/** Cached category labels for this series (`c:cat`), as written. */
	get categories(): (string | null)[] {
		const cat = firstChild(this.ser, 'c:cat')
		return readPoints(cat && findCache(cat))
	}

	/** Escape hatch: the underlying `c:ser` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ser
	}

	/** Mark the owning chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}

/** Concatenate the rich-text runs of a `c:title` element, or `null` when absent/empty. */
function readTitleText(titleEl: Element | null): string | null {
	if (!titleEl) return null
	const tx = firstChild(titleEl, 'c:tx')
	const rich = tx && firstChild(tx, 'c:rich')
	if (!rich) return null
	let out = ''
	for (const t of rich.getElementsByTagNameNS(OOXML_NS.a, 't')) out += t.textContent ?? ''
	return out === '' ? null : out
}

/** Read a `c:numFmt` child of `parent` into `{ formatCode, sourceLinked }`, or `null` when absent. */
function readNumberFormat(parent: Element): AxisNumberFormat | null {
	const nf = firstChild(parent, 'c:numFmt')
	if (!nf) return null
	return { formatCode: attr(nf, 'formatCode'), sourceLinked: boolValue(attr(nf, 'sourceLinked')) }
}

/**
 * Decode the solid colour of a properties container (`c:spPr` or `a:ln`): the
 * literal `a:srgbClr` hex, the raw `a:schemeClr` token, and whether the fill is
 * explicitly suppressed via `a:noFill`. The chart part carries no theme context,
 * so scheme tokens are surfaced raw (unresolved) rather than flattened to hex.
 */
function readSolid(container: Element): { color: string | null; schemeColor: string | null; noFill: boolean } {
	const solidFill = firstChild(container, 'a:solidFill')
	const srgb = solidFill && firstChild(solidFill, 'a:srgbClr')
	const scheme = solidFill && firstChild(solidFill, 'a:schemeClr')
	return {
		color: srgb ? (attr(srgb, 'val') ?? null) : null,
		schemeColor: scheme ? (attr(scheme, 'val') ?? null) : null,
		noFill: !!firstChild(container, 'a:noFill'),
	}
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

/** Read a cache's points (`c:pt[@idx]/c:v`) into an idx-ordered array; `c:ptCount` is the declared count. */
function readPoints(cache: Element | null): (string | null)[] {
	if (!cache) return []
	const ptCount = firstChild(cache, 'c:ptCount')
	return readIndexedPoints(
		getElements(cache, 'c:pt'),
		ptCount ? intValue(attr(ptCount, 'val')) : null,
		(pt) => firstChild(pt, 'c:v')?.textContent ?? null,
		'c:ptCount'
	)
}
