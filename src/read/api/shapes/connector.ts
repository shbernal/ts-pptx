/**
 * A connector / line (`p:cxnSp`).
 */

import { attr, firstChild, intValue, type Element } from '../../oxml/dom.js'
import { Shape } from './base.js'
import { getOrAddSpPrXfrm } from './oxml.js'
import type { ConnectionSite } from './types.js'

/** A connector / line (`p:cxnSp`). */
export class Connector extends Shape {
	readonly shapeType = 'connector' as const

	protected xfrm(): Element | null {
		const spPr = firstChild(this.element, 'p:spPr')
		return spPr ? firstChild(spPr, 'a:xfrm') : null
	}

	protected getOrAddXfrm(): Element {
		return getOrAddSpPrXfrm(this.element)
	}

	/**
	 * The connector's **start**-point shape binding (`p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn`),
	 * or `null` when the start point is unbound (a bare `p:cNvCxnSpPr`, i.e. a
	 * connector placed by static endpoint geometry). Mirrors the write API's
	 * `startShape`/`startShapeIdx` split; see {@link endConnection} for the other end.
	 */
	get startConnection(): ConnectionSite | null {
		return this.#connection('a:stCxn')
	}

	/**
	 * The connector's **end**-point shape binding (`p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn`),
	 * or `null` when the end point is unbound. See {@link startConnection}.
	 */
	get endConnection(): ConnectionSite | null {
		return this.#connection('a:endCxn')
	}

	/** Decode one `a:stCxn` / `a:endCxn` binding, resolving its `@id` to a slide shape. */
	#connection(qname: string): ConnectionSite | null {
		const nvCxnSpPr = firstChild(this.element, 'p:nvCxnSpPr')
		const cNvCxnSpPr = nvCxnSpPr && firstChild(nvCxnSpPr, 'p:cNvCxnSpPr')
		const cxn = cNvCxnSpPr && firstChild(cNvCxnSpPr, qname)
		if (!cxn) return null
		const shapeId = intValue(attr(cxn, 'id'))
		const siteIndex = intValue(attr(cxn, 'idx'))
		// CT_Connection requires both @id and @idx; an unparseable pair degrades to null
		// rather than a half-populated site.
		if (shapeId === null || siteIndex === null) return null
		return { shapeId, siteIndex, boundShape: this.slide.shapeByIdDeep(shapeId) ?? null }
	}
}
