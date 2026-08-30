/**
 * ts-pptx: Surface Plot Assembly
 *
 * Emits the classic `<c:surface3DChart>` (a 3-D surface) or `<c:surfaceChart>` (a 2-D
 * contour / top view), selected by the `surface3D` option; `wireframe` toggles the
 * mesh-only look. A surface chart interpolates a colored sheet over the category (X) and
 * series (Z) axes with the value as height (Y), so — like `bar3D` — it needs all three
 * axes and a `<c:view3D>` + floor/side/back walls (emitted by the surface branch of
 * `makeChartHeaderXml`). Reached through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { AXIS_ID_SERIES_PRIMARY } from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	catRefBlock,
	numCachePt,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
	type PlotBuilder,
} from './chart-parts.js'

/** True when the (normalized) surface options select the 3-D surface rather than a 2-D contour. */
const isSurface3D = (opts: ChartOptsInternal): boolean => opts.surface3D !== false

/** Emit the shared `<c:cat>` (strRef) + `<c:val>` (numRef) refs for a surface series. */
function surfaceCatVal(obj: OptsChartDataInternal, valFmtCode: string): string {
	const cats = firstLabelGroup(obj)
	const valCol = obj._dataIndex + dataLabels(obj).length + 1
	const numCache = el('c:numCache', null, [
		raw(el('c:formatCode', null, valFmtCode)),
		raw(voidEl('c:ptCount', { val: cats.length })),
		raw(
			dataValues(obj)
				.map((value, idx) => numCachePt(idx, value))
				.join('')
		),
	])
	const numRef = el('c:numRef', null, [
		raw(el('c:f', null, sheetRangeRef(valCol, 2, valCol, cats.length + 1))),
		raw(numCache),
	])
	return (
		el('c:cat', null, raw(catRefBlock('str', `Sheet1!$A$2:$A$${cats.length + 1}`, cats))) +
		el('c:val', null, raw(numRef))
	)
}

/** Emit a single surface series: name ref, 3-D shape props, and cat/val refs. */
function makeSurfaceSer(obj: OptsChartDataInternal, valFmtCode: string, seriesColor: string): string {
	const nameCol = obj._dataIndex + dataLabels(obj).length + 1
	// A surface series carries 3-D shape props; the surface itself is colored by value band, but the
	// per-series fill still styles the wireframe / legend key.
	const spPr = el('c:spPr', null, [
		raw(genXmlColorSelection(seriesColor)),
		raw(voidEl('a:ln')),
		raw(voidEl('a:effectLst')),
		raw(voidEl('a:sp3d')),
	])
	return el('c:ser', null, [
		raw(voidEl('c:idx', { val: obj._dataIndex })),
		raw(voidEl('c:order', { val: obj._dataIndex })),
		raw(strRefBlock(sheetCellRef(nameCol, 1), obj.name ?? '', 'compact')),
		raw(spPr),
		raw(surfaceCatVal(obj, valFmtCode)),
	])
}

/**
 * Plot a surface chart into a `<c:surface3DChart>` (3-D) or `<c:surfaceChart>` (2-D contour). Every
 * series shares the category axis; the series axis (Z) is the third axis. `valAxisId`/`catAxisId`
 * are the primary ids passed by the dispatch; the series axis uses `AXIS_ID_SERIES_PRIMARY`.
 */
export const makeSurfacePlot: PlotBuilder = (_chartType, data, opts, valAxisId, catAxisId, valFmtCode) => {
	const tag = isSurface3D(opts) ? 'surface3DChart' : 'surfaceChart'
	const chartColors = resolveChartPalette(opts)
	const sers = data
		.map((obj, idx) => makeSurfaceSer(obj, valFmtCode, paletteColor(chartColors, idx, '4472C4')))
		.join('')
	return el(`c:${tag}`, null, [
		raw(voidEl('c:wireframe', { val: opts.surfaceWireframe ? 1 : 0 })),
		raw(sers),
		// Surface, value and series axes (category X, value Y/height, series Z).
		raw(voidEl('c:axId', { val: catAxisId })),
		raw(voidEl('c:axId', { val: valAxisId })),
		raw(voidEl('c:axId', { val: AXIS_ID_SERIES_PRIMARY })),
	])
}

/**
 * Build the `<c:view3D>` + floor/side/back walls for a surface chart. A 3-D surface tilts the scene
 * (using the shared `v3DRotX`/`v3DRotY` options, which the define layer defaults to 30); a 2-D
 * contour looks straight down the value axis (rotX 90, flat perspective). These precede
 * `<c:plotArea>` in CT_Chart document order (view3D → floor → sideWall → backWall → plotArea).
 */
export function makeSurfaceScene(opts: ChartOptsInternal): string {
	const wall =
		voidEl('c:thickness', { val: 0 }) +
		el('c:spPr', null, [
			raw(voidEl('a:noFill')),
			raw(el('a:ln', null, raw(voidEl('a:noFill')))),
			raw(voidEl('a:effectLst')),
			raw(voidEl('a:sp3d')),
		])
	const scene = isSurface3D(opts)
		? [
				raw(voidEl('c:rotX', { val: opts.v3DRotX })),
				raw(voidEl('c:rotY', { val: opts.v3DRotY })),
				raw(voidEl('c:rAngAx', { val: 0 })),
			]
		: // Contour / top view: look straight down the value axis, flat perspective.
			[
				raw(voidEl('c:rotX', { val: 90 })),
				raw(voidEl('c:rotY', { val: 0 })),
				raw(voidEl('c:rAngAx', { val: 0 })),
				raw(voidEl('c:perspective', { val: 0 })),
			]
	return (
		el('c:view3D', null, scene) +
		el('c:floor', null, raw(wall)) +
		el('c:sideWall', null, raw(wall)) +
		el('c:backWall', null, raw(wall))
	)
}
