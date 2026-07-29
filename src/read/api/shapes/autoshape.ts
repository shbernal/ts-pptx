/**
 * An auto shape, text box, or placeholder (`p:sp`) — the only shape kind that holds text.
 */

import { attr, firstChild, getElements, type Element } from '../../oxml/dom.js'
import { resolveStyleFontRef, type PlaceholderRef } from '../theme-context.js'
import { TextFrame } from '../text.js'
import { Shape } from './base.js'
import { readGeometryPath } from './geometry.js'
import { getOrAddSpPrXfrm } from './oxml.js'
import type { CustomGeometry } from './types.js'

/** An auto shape, text box, or placeholder (`p:sp`). The only kind that holds text. */
export class AutoShape extends Shape {
	readonly shapeType = 'autoShape' as const

	protected xfrm(): Element | null {
		const spPr = firstChild(this.element, 'p:spPr')
		return spPr ? firstChild(spPr, 'a:xfrm') : null
	}

	protected getOrAddXfrm(): Element {
		return getOrAddSpPrXfrm(this.element)
	}

	override get hasTextFrame(): boolean {
		return firstChild(this.element, 'p:txBody') !== null
	}

	override get textFrame(): TextFrame | null {
		const txBody = firstChild(this.element, 'p:txBody')
		if (!txBody) return null
		const flatten = this.slide.themeContext()
		// Every slide shape resolves its `p:style/a:fontRef` and the presentation's
		// `p:defaultTextStyle`, so a context is always supplied; `ph` is null for a
		// non-placeholder shape, which then skips only the layout/master placeholder tiers.
		const ph = this.placeholder
		const fontRef = resolveStyleFontRef(this.element, flatten)
		return new TextFrame(txBody, this.slide.part, flatten, { ph, flatten, fontRef }, this.slide.relationships)
	}

	/**
	 * This shape's placeholder identity (`p:ph` `type`/`idx`), or `null` when it is
	 * not a placeholder. `idx` defaults to `'0'` when the attribute is absent, as
	 * PowerPoint does. Use {@link Slide.placeholder} to find a placeholder by type.
	 */
	override get placeholder(): PlaceholderRef | null {
		const nvSpPr = firstChild(this.element, 'p:nvSpPr')
		const nvPr = nvSpPr && firstChild(nvSpPr, 'p:nvPr')
		const ph = nvPr && firstChild(nvPr, 'p:ph')
		return ph ? { type: attr(ph, 'type'), idx: attr(ph, 'idx') ?? '0' } : null
	}

	/**
	 * Custom freeform geometry (`spPr/a:custGeom/a:pathLst`), or `null` when the
	 * shape uses preset geometry / none. The faithful, multi-path counterpart of
	 * {@link presetGeometry}: each `a:path` keeps its own path-unit viewport
	 * (`w`/`h`) and ordered {@link GeometryCommand}s. Coordinates are raw path-unit
	 * integers, not EMU — pair the path `w`/`h` with the shape's box size to map
	 * them into slide space.
	 */
	get customGeometry(): CustomGeometry | null {
		const props = this.properties()
		const custGeom = props && firstChild(props, 'a:custGeom')
		if (!custGeom) return null
		const pathLst = firstChild(custGeom, 'a:pathLst')
		const paths = pathLst ? getElements(pathLst, 'a:path').map((p) => readGeometryPath(p)) : []
		return { paths }
	}
}
