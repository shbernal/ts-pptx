/**
 * Reading geometry out of a shape: transform boxes, rotation and flips, and custom-geometry paths.
 *
 * Every value here comes straight off `a:xfrm` / `a:path` attributes, so these are pure element
 * readers with no theme or relationship context.
 */

import {
	ELEMENT_NODE,
	OOXML_NS,
	attr,
	boolValue,
	firstChild,
	getElements,
	intValue,
	type Element,
} from '../../oxml/dom.js'
import type { CustomGeometryPath, GeometryCommand } from './types.js'

/**
 * One `a:pt` coordinate as a raw path-unit integer. A guide-name reference
 * (the `ST_AdjCoordinate` string form) is not produced by authored freeforms;
 * a non-numeric value degrades to `0` rather than crashing (documented edge).
 */
export function ptAxis(pt: Element | undefined, axis: 'x' | 'y'): number {
	return (pt ? intValue(attr(pt, axis)) : null) ?? 0
}

/** Parse one `<a:path>` into its viewport attrs (with schema defaults) and ordered segments. */
export function readGeometryPath(path: Element): CustomGeometryPath {
	const commands: GeometryCommand[] = []
	for (let node = path.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const seg = node as Element
		if (seg.namespaceURI !== OOXML_NS.a) continue
		const pts = getElements(seg, 'a:pt')
		switch (seg.localName) {
			case 'moveTo':
				commands.push({ cmd: 'moveTo', x: ptAxis(pts[0], 'x'), y: ptAxis(pts[0], 'y') })
				break
			case 'lnTo':
				commands.push({ cmd: 'lnTo', x: ptAxis(pts[0], 'x'), y: ptAxis(pts[0], 'y') })
				break
			case 'cubicBezTo':
				commands.push({
					cmd: 'cubicBezTo',
					x1: ptAxis(pts[0], 'x'),
					y1: ptAxis(pts[0], 'y'),
					x2: ptAxis(pts[1], 'x'),
					y2: ptAxis(pts[1], 'y'),
					x: ptAxis(pts[2], 'x'),
					y: ptAxis(pts[2], 'y'),
				})
				break
			case 'quadBezTo':
				commands.push({
					cmd: 'quadBezTo',
					x1: ptAxis(pts[0], 'x'),
					y1: ptAxis(pts[0], 'y'),
					x: ptAxis(pts[1], 'x'),
					y: ptAxis(pts[1], 'y'),
				})
				break
			case 'arcTo':
				commands.push({
					cmd: 'arcTo',
					wR: intValue(attr(seg, 'wR')) ?? 0,
					hR: intValue(attr(seg, 'hR')) ?? 0,
					stAng: (intValue(attr(seg, 'stAng')) ?? 0) / 60000,
					swAng: (intValue(attr(seg, 'swAng')) ?? 0) / 60000,
				})
				break
			case 'close':
				commands.push({ cmd: 'close' })
				break
		}
	}
	return {
		w: intValue(attr(path, 'w')) ?? 0,
		h: intValue(attr(path, 'h')) ?? 0,
		fill: attr(path, 'fill') ?? 'norm',
		stroke: boolValue(attr(path, 'stroke')) ?? true,
		commands,
	}
}

/** Validate and round an EMU geometry value; extents (`cx`/`cy`) must be non-negative. */
export function toEmu(value: number, attribute: string, allowNegative: boolean): number {
	if (!Number.isFinite(value)) throw new Error(`${attribute} must be a finite number of EMU, got ${value}`)
	if (!allowNegative && value < 0) throw new Error(`${attribute} must be non-negative, got ${value}`)
	return Math.round(value)
}

/** A point + extent pair (`a:off`/`a:ext` or `a:chOff`/`a:chExt`) from a transform, or `null` if either is incomplete. */
export function readBox(
	xfrm: Element,
	offName: string,
	extName: string
): { x: number; y: number; cx: number; cy: number } | null {
	const off = firstChild(xfrm, offName)
	const ext = firstChild(xfrm, extName)
	const x = off && intValue(attr(off, 'x'))
	const y = off && intValue(attr(off, 'y'))
	const cx = ext && intValue(attr(ext, 'cx'))
	const cy = ext && intValue(attr(ext, 'cy'))
	if (x === null || y === null || cx === null || cy === null) return null
	return { x, y, cx, cy }
}

export function rotationDegrees(xfrm: Element): number {
	const rot = intValue(attr(xfrm, 'rot'))
	return rot === null ? 0 : rot / 60000
}

export function transformFlipH(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipH')) === true
}

export function transformFlipV(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipV')) === true
}
