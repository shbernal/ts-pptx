/**
 * EMU geometry rescaling for slides imported across differing canvas sizes.
 *
 * {@link computeRescale} turns a source/target {@link SlideSize} pair into an affine
 * transform; the `rescale*` functions apply it in place to a live `p:spTree` and to the
 * table grids inside it. Pure DOM mutation -- the caller decides which parts to visit
 * and tracks which have already been rewritten.
 */

import type { SlideSize } from '../presentation-types.js'
import {
	ELEMENT_NODE,
	OOXML_NS,
	attr,
	firstChild,
	getElements,
	intValue,
	setAttr,
	type Element,
} from '../../oxml/dom.js'

/** An EMU coordinate remap: `newX = x*sx + dx`, `newY = y*sy + dy`; sizes scale by `sx`/`sy` only. */
export interface RescaleTransform {
	sx: number
	sy: number
	dx: number
	dy: number
}

/** Build the EMU transform mapping source-canvas coordinates onto the target canvas (see {@link ImportSlideOptions.rescale}). */
export function computeRescale(source: SlideSize, target: SlideSize, mode: 'fit' | 'stretch'): RescaleTransform {
	if (mode === 'stretch') {
		return { sx: target.widthEmu / source.widthEmu, sy: target.heightEmu / source.heightEmu, dx: 0, dy: 0 }
	}
	// 'fit': uniform scale (no distortion), centering the slack on the longer axis.
	const scale = Math.min(target.widthEmu / source.widthEmu, target.heightEmu / source.heightEmu)
	return {
		sx: scale,
		sy: scale,
		dx: (target.widthEmu - source.widthEmu * scale) / 2,
		dy: (target.heightEmu - source.heightEmu * scale) / 2,
	}
}

/**
 * Rescale every top-level transform on a `p:spTree`: shapes/pictures/connectors
 * (`p:spPr/a:xfrm`), groups (`p:grpSpPr/a:xfrm` — only the group's own off/ext, so
 * its children remap via the unchanged `chOff`/`chExt` rather than being recursed
 * into), and graphic frames (`p:graphicFrame/p:xfrm`, plus any inner table grid).
 * Placeholders that inherit geometry carry no `a:xfrm` and are left to inherit from
 * the (separately rescaled) layout/master.
 */
export function rescaleSpTree(spTree: Element, t: RescaleTransform): void {
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (el.namespaceURI !== OOXML_NS.p) continue
		switch (el.localName) {
			case 'sp':
			case 'pic':
			case 'cxnSp': {
				const spPr = firstChild(el, 'p:spPr')
				rescaleXfrm(spPr && firstChild(spPr, 'a:xfrm'), t)
				break
			}
			case 'grpSp': {
				const grpSpPr = firstChild(el, 'p:grpSpPr')
				rescaleXfrm(grpSpPr && firstChild(grpSpPr, 'a:xfrm'), t)
				break
			}
			case 'graphicFrame': {
				rescaleXfrm(firstChild(el, 'p:xfrm'), t)
				rescaleTableGrid(el, t)
				break
			}
		}
	}
}

/** Rewrite an `a:xfrm`/`p:xfrm`: `a:off` is repositioned (scale + translate); `a:ext` is resized (scale only). */
function rescaleXfrm(xfrm: Element | null, t: RescaleTransform): void {
	if (!xfrm) return
	const off = firstChild(xfrm, 'a:off')
	if (off) {
		const x = intValue(attr(off, 'x'))
		const y = intValue(attr(off, 'y'))
		if (x !== null) setAttr(off, 'x', String(Math.round(x * t.sx + t.dx)))
		if (y !== null) setAttr(off, 'y', String(Math.round(y * t.sy + t.dy)))
	}
	const ext = firstChild(xfrm, 'a:ext')
	if (ext) {
		const cx = intValue(attr(ext, 'cx'))
		const cy = intValue(attr(ext, 'cy'))
		if (cx !== null) setAttr(ext, 'cx', String(Math.max(0, Math.round(cx * t.sx))))
		if (cy !== null) setAttr(ext, 'cy', String(Math.max(0, Math.round(cy * t.sy))))
	}
}

/**
 * Scale a graphic-frame table's intrinsic dimensions — column widths by `sx`, row
 * heights by `sy` — so the table resizes with the slide (PowerPoint derives table
 * geometry from the grid, not the frame `a:ext`). No-op for non-table frames.
 */
function rescaleTableGrid(graphicFrame: Element, t: RescaleTransform): void {
	const graphic = firstChild(graphicFrame, 'a:graphic')
	const graphicData = graphic && firstChild(graphic, 'a:graphicData')
	const tbl = graphicData && firstChild(graphicData, 'a:tbl')
	if (!tbl) return
	const grid = firstChild(tbl, 'a:tblGrid')
	if (grid) {
		for (const col of getElements(grid, 'a:gridCol')) {
			const w = intValue(attr(col, 'w'))
			if (w !== null) setAttr(col, 'w', String(Math.max(0, Math.round(w * t.sx))))
		}
	}
	for (const tr of getElements(tbl, 'a:tr')) {
		const h = intValue(attr(tr, 'h'))
		if (h !== null) setAttr(tr, 'h', String(Math.max(0, Math.round(h * t.sy))))
	}
}
