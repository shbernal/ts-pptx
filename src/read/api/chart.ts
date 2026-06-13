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
import { OOXML_NS, attr, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'

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
		return this.#chartGroups().map((group) => group.localName.replace(/Chart$/, ''))
	}

	/** The first chart-group type, or `null` for an empty plot area. */
	get chartType(): string | null {
		return this.chartTypes[0] ?? null
	}

	/** The chart title (`c:chart/c:title` rich text), or `null` when absent/auto. */
	get title(): string | null {
		const title = this.#chart() && firstChild(this.#chart()!, 'c:title')
		if (!title) return null
		const tx = firstChild(title, 'c:tx')
		const rich = tx && firstChild(tx, 'c:rich')
		if (!rich) return null
		const texts = rich.getElementsByTagNameNS(OOXML_NS.a, 't')
		let out = ''
		for (let i = 0; i < texts.length; i++) out += texts[i].textContent ?? ''
		return out === '' ? null : out
	}

	/** The data series (`c:ser`) across all chart groups, in document order. */
	get series(): ChartSeries[] {
		return this.#chartGroups()
			.flatMap((group) => getElements(group, 'c:ser'))
			.map((ser) => new ChartSeries(ser))
	}

	/**
	 * Category labels, read from the first series' cached categories
	 * (`c:cat`), as written. Empty when the chart has no category axis.
	 */
	get categories(): (string | null)[] {
		const firstSer = this.series[0]
		return firstSer ? firstSer.categories : []
	}

	/** The underlying `c:chartSpace` element, for advanced reads. */
	get element_(): Element | null {
		return this.part.dom.documentElement
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
			if (element.namespaceURI === OOXML_NS.c && element.localName.endsWith('Chart')) groups.push(element)
		}
		return groups
	}
}

/** One data series (`c:ser`) of a chart. */
export class ChartSeries {
	constructor(private readonly ser: Element) {}

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
		return readPoints(val && findCache(val)).map(toNumberOrNull)
	}

	/** Cached category labels for this series (`c:cat`), as written. */
	get categories(): (string | null)[] {
		const cat = firstChild(this.ser, 'c:cat')
		return readPoints(cat && findCache(cat))
	}

	/** The underlying `c:ser` element. */
	get element_(): Element {
		return this.ser
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

/** Read a cache's points (`c:pt[@idx]/c:v`) into an idx-ordered array sized by `c:ptCount`. */
function readPoints(cache: Element | null): (string | null)[] {
	if (!cache) return []
	const ptCount = firstChild(cache, 'c:ptCount')
	let count = ptCount ? (intValue(attr(ptCount, 'val')) ?? 0) : 0
	const pts = getElements(cache, 'c:pt')
	for (const pt of pts) {
		const idx = intValue(attr(pt, 'idx')) ?? 0
		if (idx + 1 > count) count = idx + 1
	}
	const points: (string | null)[] = new Array(count).fill(null)
	for (const pt of pts) {
		const idx = intValue(attr(pt, 'idx')) ?? 0
		points[idx] = firstChild(pt, 'c:v')?.textContent ?? null
	}
	return points
}

function toNumberOrNull(value: string | null): number | null {
	if (value === null || value === '') return null
	const number = Number(value)
	return Number.isFinite(number) ? number : null
}
