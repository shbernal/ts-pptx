/**
 * A graphic frame (`p:graphicFrame`) — the host element for a table, a chart, a chartEx
 * chart, or a SmartArt diagram.
 *
 * Which one it holds is told by the `a:graphicData/@uri`, so the accessors below each check
 * that URI before wrapping the payload. {@link GraphicFrame.graphicDataUri} reports it raw,
 * which is the only signal for the frames this library still does not model (OLE, ink):
 * without it a consumer cannot tell "this frame holds content I cannot reach" from "this
 * frame holds a chart with nothing in it", since every predicate answers `false` for both.
 */

import { attr, firstChild, getOrAddChild, type Element } from '../../oxml/dom.js'
import { Chart } from '../chart.js'
import { ChartEx } from '../chartex.js'
import { Diagram } from '../diagram.js'
import { Table } from '../table.js'
import { Shape } from './base.js'
import type { ShapeProperties } from './oxml.js'
import { UnsupportedFeatureError } from '../../../errors.js'

const A_TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'
const A_CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
// chartEx (Office-2016 chart family) graphicData URI + `cx:chart` reference child namespace.
const A_CHARTEX_URI = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
// SmartArt: the DrawingML diagram namespace, whose payload is a `dgm:relIds` reference.
const A_DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'

/** A graphic frame (`p:graphicFrame`) — host for tables, charts and SmartArt diagrams. */
export class GraphicFrame extends Shape {
	readonly shapeType = 'graphicFrame' as const

	protected override xfrm(): Element | null {
		// graphicFrame carries its own `p:xfrm` directly, not inside `p:spPr`.
		return firstChild(this.element, 'p:xfrm')
	}

	protected override getOrAddXfrm(): Element {
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

	/** Whether this frame hosts a SmartArt diagram (`a:graphicData/@uri` is the diagram URI). */
	get hasDiagram(): boolean {
		return this.#graphicDataUri() === A_DIAGRAM_URI
	}

	/**
	 * The `a:graphicData/@uri` verbatim — the namespace naming what this frame hosts, or
	 * `null` when the frame carries no `a:graphicData`.
	 *
	 * Every `has*` predicate is a comparison against this, so a frame whose URI matches none
	 * of them is one the read model does not decode (an OLE object, an ink `p:contentPart`).
	 * Reading the URI is what lets a consumer say *which* construct it could not reach
	 * instead of inferring loss from four `false`s.
	 */
	get graphicDataUri(): string | null {
		return this.#graphicDataUri()
	}

	/** The hosted table, or `null` when this frame is not a table. */
	get table(): Table | null {
		if (!this.hasTable) return null
		const graphicData = this.#graphicData()
		const tbl = graphicData && firstChild(graphicData, 'a:tbl')
		if (!tbl) return null
		return new Table(tbl, this.host.part, this.host.themeContext(), this.host.opc, this.host.relationships)
	}

	/** The hosted chart, or `null` when this frame is not a chart or its part is missing. */
	get chart(): Chart | null {
		if (!this.hasChart) return null
		const graphicData = this.#graphicData()
		const chartRef = graphicData && firstChild(graphicData, 'c:chart')
		const relId = chartRef && attr(chartRef, 'r:id')
		if (!relId) return null
		const partName = this.host.relationships.resolveTarget(relId)
		const part = this.host.opc.part(partName)
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
		const partName = this.host.relationships.resolveTarget(relId)
		const part = this.host.opc.part(partName)
		return part ? new ChartEx(part) : null
	}

	/**
	 * The hosted SmartArt diagram, or `null` when this frame is not a diagram or its data
	 * part is missing. The payload child is `dgm:relIds`, which names four parts by
	 * relationship id; the `@r:dm` one is the data model this resolves.
	 */
	get diagram(): Diagram | null {
		if (!this.hasDiagram) return null
		const graphicData = this.#graphicData()
		const relIds = graphicData && firstChild(graphicData, 'dgm:relIds')
		if (!relIds) return null
		const relId = attr(relIds, 'r:dm')
		if (!relId) return null
		const partName = this.host.relationships.resolveTarget(relId)
		const part = this.host.opc.part(partName)
		return part ? new Diagram(part, relIds, this.host.opc, this.host.relationships, this.host.themeContext()) : null
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
