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
	numberValue,
	type Element,
} from '../../oxml/dom.js'
import type { CustomGeometryPath, GeometryCommand } from './types.js'
import { InvalidOptionError } from '../../../errors.js'
import { ANGLE_UNITS_PER_DEGREE } from '../../../units.js'

/**
 * One `a:pt` coordinate as a raw path-unit integer. A guide-name reference
 * (the `ST_AdjCoordinate` string form) is not produced by authored freeforms;
 * a non-numeric value degrades to `0` rather than crashing (documented edge).
 */
function ptAxis(pt: Element | undefined, axis: 'x' | 'y'): number {
	return (pt ? numberValue(attr(pt, axis)) : null) ?? 0
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
					wR: numberValue(attr(seg, 'wR')) ?? 0,
					hR: numberValue(attr(seg, 'hR')) ?? 0,
					stAng: (numberValue(attr(seg, 'stAng')) ?? 0) / ANGLE_UNITS_PER_DEGREE,
					swAng: (numberValue(attr(seg, 'swAng')) ?? 0) / ANGLE_UNITS_PER_DEGREE,
				})
				break
			case 'close':
				commands.push({ cmd: 'close' })
				break
		}
	}
	return {
		w: numberValue(attr(path, 'w')) ?? 0,
		h: numberValue(attr(path, 'h')) ?? 0,
		fill: attr(path, 'fill') ?? 'norm',
		stroke: boolValue(attr(path, 'stroke')) ?? true,
		commands,
	}
}

/** Validate and round an EMU geometry value; extents (`cx`/`cy`) must be non-negative. */
export function toEmu(value: number, attribute: string, allowNegative: boolean): number {
	if (!Number.isFinite(value))
		throw new InvalidOptionError('coord/non-finite', `${attribute} must be a finite number of EMU, got ${value}`)
	if (!allowNegative && value < 0)
		throw new InvalidOptionError('coord/negative', `${attribute} must be non-negative, got ${value}`)
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
	const x = off && numberValue(attr(off, 'x'))
	const y = off && numberValue(attr(off, 'y'))
	const cx = ext && numberValue(attr(ext, 'cx'))
	const cy = ext && numberValue(attr(ext, 'cy'))
	if (x === null || y === null || cx === null || cy === null) return null
	return { x, y, cx, cy }
}

export function rotationDegrees(xfrm: Element): number {
	const rot = numberValue(attr(xfrm, 'rot'))
	return rot === null ? 0 : rot / ANGLE_UNITS_PER_DEGREE
}

export function transformFlipH(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipH')) === true
}

export function transformFlipV(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipV')) === true
}
