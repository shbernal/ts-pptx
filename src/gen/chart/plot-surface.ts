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
import type { WorksheetLayout } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { xsdBool } from '../../ooxml/xsd-boolean.js'
import { catValRefs, paletteColor, resolveChartPalette, seriesNameRef, type PlotBuilder } from './chart-parts.js'

/** True when the (normalized) surface options select the 3-D surface rather than a 2-D contour. */
const isSurface3D = (opts: ChartOptsInternal): boolean => opts.surface3D !== false

/** Emit a single surface series: name ref, 3-D shape props, and cat/val refs. */
function makeSurfaceSer(
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	valFmtCode: string,
	seriesColor: string,
	sheet: WorksheetLayout
): string {
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
		raw(seriesNameRef(obj, sheet)),
		raw(spPr),
		raw(catValRefs(obj, opts, valFmtCode, sheet, { multiLevel: false })),
	])
}

/**
 * Plot a surface chart into a `<c:surface3DChart>` (3-D) or `<c:surfaceChart>` (2-D contour). Every
 * series shares the category axis; the series axis (Z) is the third axis. `valAxisId`/`catAxisId`
 * are the primary ids passed by the dispatch; the series axis uses `AXIS_ID_SERIES_PRIMARY`.
 */
export const makeSurfacePlot: PlotBuilder = (_chartType, data, opts, valAxisId, catAxisId, valFmtCode, sheet) => {
	const tag = isSurface3D(opts) ? 'surface3DChart' : 'surfaceChart'
	const chartColors = resolveChartPalette(opts)
	const sers = data
		.map((obj, idx) => makeSurfaceSer(obj, opts, valFmtCode, paletteColor(chartColors, idx, '4472C4'), sheet))
		.join('')
	return el(`c:${tag}`, null, [
		raw(voidEl('c:wireframe', { val: xsdBool(opts.surfaceWireframe) })),
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
