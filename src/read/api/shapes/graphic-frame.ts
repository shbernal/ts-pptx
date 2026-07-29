/**
 * A graphic frame (`p:graphicFrame`) — the host element for a table, a chart, or a chartEx chart.
 *
 * Which one it holds is told by the `a:graphicData/@uri`, so the accessors below each check that
 * URI before wrapping the payload; `graphicDataUri` is there for the frames this library does not
 * model (SmartArt, OLE, ink), which read as a frame with no table and no chart.
 */

import { attr, firstChild, getOrAddChild, type Element } from '../../oxml/dom.js'
import { Chart } from '../chart.js'
import { ChartEx } from '../chartex.js'
import { Table } from '../table.js'
import { Shape } from './base.js'
import type { ShapeProperties } from './oxml.js'
import { UnsupportedFeatureError } from '../../../errors.js'

const A_TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'
const A_CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
// chartEx (Office-2016 chart family) graphicData URI + `cx:chart` reference child namespace.
const A_CHARTEX_URI = 'http://schemas.microsoft.com/office/drawing/2014/chartex'

/** A graphic frame (`p:graphicFrame`) — host for tables and charts. */
export class GraphicFrame extends Shape {
	readonly shapeType = 'graphicFrame' as const

	protected xfrm(): Element | null {
		// graphicFrame carries its own `p:xfrm` directly, not inside `p:spPr`.
		return firstChild(this.element, 'p:xfrm')
	}

	protected getOrAddXfrm(): Element {
		// p:xfrm sits between p:nvGraphicFramePr and a:graphic.
		return getOrAddChild(this.element, 'p:xfrm', ['a:graphic', 'p:extLst'])
	}

	// A graphicFrame has no p:spPr; its hosted table/chart carries its own fill
	// model. There is nothing to get-or-add, so fill and line setters reject it.
	protected override getOrAddProperties(): ShapeProperties {
		throw new UnsupportedFeatureError(
			'shape/shape-properties-unsupported',
			'graphicFrame shapes have no shape properties; fill and line colours are not supported'
		)
	}

	/** Whether this frame hosts a table (`a:graphicData/@uri` is the table URI). */
	get hasTable(): boolean {
		return this.#graphicDataUri() === A_TABLE_URI
	}

	/** Whether this frame hosts a classic chart (`a:graphicData/@uri` is the chart URI). */
	get hasChart(): boolean {
		return this.#graphicDataUri() === A_CHART_URI
	}

	/** Whether this frame hosts a chartEx chart (`a:graphicData/@uri` is the chartEx URI). */
	get hasChartEx(): boolean {
		return this.#graphicDataUri() === A_CHARTEX_URI
	}

	/** The hosted table, or `null` when this frame is not a table. */
	get table(): Table | null {
		if (!this.hasTable) return null
		const graphicData = this.#graphicData()
		const tbl = graphicData && firstChild(graphicData, 'a:tbl')
		if (!tbl) return null
		return new Table(
			tbl,
			this.slide.part,
			this.slide.themeContext(),
			this.slide.presentation.opc,
			this.slide.relationships
		)
	}

	/** The hosted chart, or `null` when this frame is not a chart or its part is missing. */
	get chart(): Chart | null {
		if (!this.hasChart) return null
		const graphicData = this.#graphicData()
		const chartRef = graphicData && firstChild(graphicData, 'c:chart')
		const relId = chartRef && attr(chartRef, 'r:id')
		if (!relId) return null
		const partName = this.slide.relationships.resolveTarget(relId)
		const part = this.slide.presentation.opc.part(partName)
		return part ? new Chart(part) : null
	}

	/**
	 * The hosted chartEx chart (waterfall/funnel/treemap/…), or `null` when this
	 * frame is not a chartEx chart or its part is missing. The reference child is
	 * `cx:chart` (not the classic `c:chart`), carrying the MS `chartEx` rel id.
	 */
	get chartEx(): ChartEx | null {
		if (!this.hasChartEx) return null
		const graphicData = this.#graphicData()
		const chartRef = graphicData && firstChild(graphicData, 'cx:chart')
		const relId = chartRef && attr(chartRef, 'r:id')
		if (!relId) return null
		const partName = this.slide.relationships.resolveTarget(relId)
		const part = this.slide.presentation.opc.part(partName)
		return part ? new ChartEx(part) : null
	}

	#graphicData(): Element | null {
		const graphic = firstChild(this.element, 'a:graphic')
		return graphic ? firstChild(graphic, 'a:graphicData') : null
	}

	#graphicDataUri(): string | null {
		const graphicData = this.#graphicData()
		return graphicData ? attr(graphicData, 'uri') : null
	}
}
