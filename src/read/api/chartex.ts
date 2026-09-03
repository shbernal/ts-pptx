/**
 * Read-only proxies for a chartEx (`cx:chartSpace`) chart hosted in a
 * `p:graphicFrame`. chartEx is the Office-2016 chart family — waterfall, funnel,
 * treemap, sunburst, histogram, pareto, box-and-whisker, region map — a separate
 * subsystem from the classic `c:chart` (see {@link ./chart.js `Chart`}).
 *
 * A chartEx chart lives in its own part (`/ppt/charts/chartEx{N}.xml`) whose root
 * is `cx:chartSpace`, referenced from the graphic frame by relationship id (the MS
 * `chartEx` rel, not the ECMA `chart` rel). The frame itself is wrapped in an
 * `mc:AlternateContent` — the reader unwraps the `mc:Choice` to reach it.
 *
 * Unlike the classic side, the layout geometry is keyed on `cx:series/@layoutId`
 * (`waterfall`, `funnel`, `treemap`, `clusteredColumn` for histogram/pareto, …)
 * rather than the plot-group element name, and category/value data lives in a
 * top-level `cx:chartData/cx:data` block the series point at by `cx:dataId`. This
 * surface reads the layout, title, legend, series, and the cached data; the
 * embedded workbook, the style/colors sidecars, and (for a region map) the online
 * geo-cache are out of scope.
 */
import type { Part } from '../opc/part.js'
import {
	attr,
	boolAttr,
	concatDrawingMLText,
	type Element,
	firstChild,
	getElements,
	numberAttr,
	numberValue,
} from '../oxml/dom.js'
import { readIndexedPoints } from '../oxml/point-cache.js'

/** A chartEx legend (`cx:legend`) — a leaf element carrying position/alignment attributes. */
export interface ChartExLegend {
	/** Legend position (`@pos`): `t`/`b`/`l`/`r`. */
	position: string | null
	/** Legend alignment along its edge (`@align`), e.g. `ctr`. */
	align: string | null
	/** Whether the legend overlays the plot area (`@overlay`). */
	overlay: boolean | null
}

/** A chartEx series' data-label toggles (`cx:dataLabels` + its `cx:visibility` child). */
export interface ChartExDataLabels {
	/** Label position (`@pos`), e.g. `outEnd`, or `null` when unset. */
	position: string | null
	/** Show the point value (`cx:visibility/@value`). */
	value: boolean | null
	/** Show the series name (`cx:visibility/@seriesName`). */
	seriesName: boolean | null
	/** Show the category name (`cx:visibility/@categoryName`). */
	categoryName: boolean | null
}

/** A chartEx graphic-frame's chart, backed by its `cx:chartSpace` part. */
export class ChartEx {
	constructor(
		/** The chartEx chart's OPC part (`/ppt/charts/chartEx{N}.xml`). */
		readonly part: Part
	) {}

	/** Partname of the chartEx chart part. */
	get partName(): string {
		return this.part.partName
	}

	/**
	 * The `cx:series/@layoutId` tokens in plot-area document order — the raw chartEx
	 * layout names (`waterfall`, `funnel`, `treemap`, `sunburst`, `boxWhisker`,
	 * `regionMap`, or `clusteredColumn`/`paretoLine` for histogram & pareto). This is
	 * the OOXML truth: a histogram and a pareto both surface as `clusteredColumn`, so
	 * the token does not always round-trip to the write-side `ChartType`.
	 */
	get layoutIds(): string[] {
		return this.#seriesElements()
			.map((ser) => attr(ser, 'layoutId'))
			.filter((id): id is string => id !== null && id !== '')
	}

	/** The first series' `@layoutId`, or `null` for an empty plot area. */
	get layoutId(): string | null {
		return this.layoutIds[0] ?? null
	}

	/** The chart title (`cx:chart/cx:title` rich text), or `null` when absent. */
	get title(): string | null {
		const chart = this.#chart()
		const title = chart && firstChild(chart, 'cx:title')
		if (!title) return null
		const tx = firstChild(title, 'cx:tx')
		return concatDrawingMLText(tx && firstChild(tx, 'cx:rich'))
	}

	/** The legend (`cx:chart/cx:legend`) position/alignment, or `null` when hidden. */
	get legend(): ChartExLegend | null {
		const chart = this.#chart()
		const legend = chart && firstChild(chart, 'cx:legend')
		if (!legend) return null
		return {
			position: attr(legend, 'pos'),
			align: attr(legend, 'align'),
			overlay: boolAttr(legend, 'overlay'),
		}
	}

	/** The data series (`cx:plotAreaRegion/cx:series`) in document order. */
	get series(): ChartExSeries[] {
		return this.#seriesElements().map((ser) => new ChartExSeries(ser, this))
	}

	/** The plot-area axes (`cx:plotArea/cx:axis`) in document order. */
	get axes(): ChartExAxis[] {
		const plotArea = this.#plotArea()
		return plotArea ? getElements(plotArea, 'cx:axis').map((ax) => new ChartExAxis(ax, this)) : []
	}

	/**
	 * Category labels from the first series' cached leaf level (`cx:strDim`'s first
	 * `cx:lvl`). Empty for a category-less layout (a histogram bins raw observations).
	 */
	get categories(): (string | null)[] {
		return this.series[0]?.categories ?? []
	}

	/** Escape hatch: the underlying `cx:chartSpace` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element | null {
		return this.part.dom.documentElement
	}

	/** Mark the chartEx part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/**
	 * The `cx:data` block a series' `cx:dataId` points at (`cx:chartData/cx:data[@id]`),
	 * or `null` when the id is unset or unmatched.
	 */
	dataElementById(id: number | null): Element | null {
		if (id === null) return null
		const chartData = this.#chartData()
		if (!chartData) return null
		return getElements(chartData, 'cx:data').find((data) => numberValue(attr(data, 'id')) === id) ?? null
	}

	#root(): Element | null {
		return this.part.dom.documentElement
	}

	#chartData(): Element | null {
		const root = this.#root()
		return root ? firstChild(root, 'cx:chartData') : null
	}

	#chart(): Element | null {
		const root = this.#root()
		return root ? firstChild(root, 'cx:chart') : null
	}

	#plotArea(): Element | null {
		const chart = this.#chart()
		return chart ? firstChild(chart, 'cx:plotArea') : null
	}

	#seriesElements(): Element[] {
		const plotArea = this.#plotArea()
		const region = plotArea && firstChild(plotArea, 'cx:plotAreaRegion')
		return region ? getElements(region, 'cx:series') : []
	}
}

/** One data series (`cx:series`) of a chartEx chart. */
export class ChartExSeries {
	constructor(
		private readonly ser: Element,
		private readonly chart: ChartEx
	) {}

	/** The layout token (`@layoutId`) that keys this series' geometry, e.g. `waterfall`. */
	get layoutId(): string | null {
		return attr(this.ser, 'layoutId')
	}

	/** The deterministic series GUID (`@uniqueId`), or `null`. */
	get uniqueId(): string | null {
		return attr(this.ser, 'uniqueId')
	}

	/**
	 * The owner series index (`@ownerIdx`) for a derived series (a `paretoLine`
	 * derives its data from series 0), or `null` for a self-contained series.
	 */
	get ownerIndex(): number | null {
		return numberValue(attr(this.ser, 'ownerIdx'))
	}

	/** Series name from the cached `cx:tx/cx:txData/cx:v`, or `null` when unnamed. */
	get name(): string | null {
		const tx = firstChild(this.ser, 'cx:tx')
		const txData = tx && firstChild(tx, 'cx:txData')
		const v = txData && firstChild(txData, 'cx:v')
		return v ? (v.textContent ?? null) : null
	}

	/** The data-block id this series plots (`cx:dataId/@val`), or `null`. */
	get dataId(): number | null {
		const dataId = firstChild(this.ser, 'cx:dataId')
		return dataId ? numberValue(attr(dataId, 'val')) : null
	}

	/** Cached numeric values (`cx:numDim`); non-numeric or missing points are `null`. */
	get values(): (number | null)[] {
		const data = this.chart.dataElementById(this.dataId)
		const numDim = data && firstChild(data, 'cx:numDim')
		return readLevelPoints(numDim && firstChild(numDim, 'cx:lvl')).map(numberValue)
	}

	/**
	 * Cached category labels (`cx:strDim`'s leaf `cx:lvl` — the first level, which
	 * the writer emits leaf-first). Empty when the data block carries no `cx:strDim`.
	 */
	get categories(): (string | null)[] {
		const data = this.chart.dataElementById(this.dataId)
		const strDim = data && firstChild(data, 'cx:strDim')
		return readLevelPoints(strDim && firstChild(strDim, 'cx:lvl'))
	}

	/** The data-label toggles (`cx:dataLabels`), or `null` when the series shows none. */
	get dataLabels(): ChartExDataLabels | null {
		const dLbls = firstChild(this.ser, 'cx:dataLabels')
		if (!dLbls) return null
		const vis = firstChild(dLbls, 'cx:visibility')
		return {
			position: attr(dLbls, 'pos'),
			value: vis ? boolAttr(vis, 'value') : null,
			seriesName: vis ? boolAttr(vis, 'seriesName') : null,
			categoryName: vis ? boolAttr(vis, 'categoryName') : null,
		}
	}

	/** Escape hatch: the underlying `cx:series` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ser
	}

	/** Mark the owning chartEx part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.chart.part.markDirty()
	}
}

/**
 * One axis (`cx:axis`) of a chartEx plot area. chartEx axes carry no `catAx`/`valAx`
 * element-name distinction; the kind is read from the scaling child
 * (`cx:catScaling` vs `cx:valScaling`).
 */
export class ChartExAxis {
	constructor(
		private readonly ax: Element,
		/** The owning chartEx chart, so {@link markDirty} can reach its part from {@link element_}. */
		private readonly chart: ChartEx
	) {}

	/** Axis id (`@id`), the value a series' `cx:axisId` binds to. */
	get id(): number | null {
		return numberValue(attr(this.ax, 'id'))
	}

	/** Axis kind from its scaling child: `cat` (`cx:catScaling`) / `val` (`cx:valScaling`). */
	get kind(): 'cat' | 'val' | null {
		if (firstChild(this.ax, 'cx:catScaling')) return 'cat'
		if (firstChild(this.ax, 'cx:valScaling')) return 'val'
		return null
	}

	/**
	 * Category-axis gap width (`cx:catScaling/@gapWidth`) as a fraction (1.0 = 100%,
	 * the chartEx convention, unlike the classic integer percent), or `null`.
	 */
	get gapWidth(): number | null {
		const cat = firstChild(this.ax, 'cx:catScaling')
		return cat ? numberAttr(cat, 'gapWidth') : null
	}

	/** Value-axis scale minimum (`cx:valScaling/@min`), or `null` when auto. */
	get min(): number | null {
		const val = firstChild(this.ax, 'cx:valScaling')
		return val ? numberAttr(val, 'min') : null
	}

	/** Value-axis scale maximum (`cx:valScaling/@max`), or `null` when auto. */
	get max(): number | null {
		const val = firstChild(this.ax, 'cx:valScaling')
		return val ? numberAttr(val, 'max') : null
	}

	/** Whether the axis draws major gridlines (`cx:majorGridlines`). */
	get majorGridlines(): boolean {
		return !!firstChild(this.ax, 'cx:majorGridlines')
	}

	/** Whether the axis draws tick labels (`cx:tickLabels`). */
	get tickLabels(): boolean {
		return !!firstChild(this.ax, 'cx:tickLabels')
	}

	/** Escape hatch: the underlying `cx:axis` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ax
	}

	/** Mark the owning chartEx part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.chart.part.markDirty()
	}
}

/** Read a `cx:lvl`'s points (`cx:pt[@idx]` text) into an idx-ordered array; `@ptCount` is the declared count. */
function readLevelPoints(lvl: Element | null): (string | null)[] {
	if (!lvl) return []
	return readIndexedPoints(
		getElements(lvl, 'cx:pt'),
		numberValue(attr(lvl, 'ptCount')),
		(pt) => pt.textContent ?? null,
		'cx:lvl/@ptCount'
	)
}
