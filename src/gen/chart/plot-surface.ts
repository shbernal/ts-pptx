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

import { ChartType } from '../../core-enums.js'
import { AXIS_ID_SERIES_PRIMARY, BARCHART_COLORS } from '../../core-enums-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el } from '../oxml/el.js'
import { numCachePt } from './chart-parts.js'

/** True when the (normalized) surface options select the 3-D surface rather than a 2-D contour. */
export const isSurface3D = (opts: ChartOptsInternal): boolean => opts.surface3D !== false

/** Emit the shared `<c:cat>` (strRef) + `<c:val>` (numRef) refs for a surface series. */
function surfaceCatVal(obj: OptsChartDataInternal, valFmtCode: string): string {
	const cats = firstLabelGroup(obj)
	const valCol = obj._dataIndex + dataLabels(obj).length + 1
	let strXml = '<c:cat><c:strRef>'
	strXml += `<c:f>Sheet1!$A$2:$A$${cats.length + 1}</c:f>`
	strXml += `<c:strCache><c:ptCount val="${cats.length}"/>`
	cats.forEach((label, idx) => (strXml += `<c:pt idx="${idx}">${el('c:v', null, label)}</c:pt>`))
	strXml += '</c:strCache></c:strRef></c:cat>'

	strXml += '<c:val><c:numRef>'
	strXml += `<c:f>${sheetRangeRef(valCol, 2, valCol, cats.length + 1)}</c:f>`
	strXml += `<c:numCache><c:formatCode>${valFmtCode}</c:formatCode><c:ptCount val="${cats.length}"/>`
	dataValues(obj).forEach((value, idx) => (strXml += numCachePt(idx, value)))
	strXml += '</c:numCache></c:numRef></c:val>'
	return strXml
}

/** Emit a single surface series: name ref, 3-D shape props, and cat/val refs. */
function makeSurfaceSer(obj: OptsChartDataInternal, valFmtCode: string, seriesColor: string): string {
	const nameCol = obj._dataIndex + dataLabels(obj).length + 1
	let strXml = '<c:ser>'
	strXml += `<c:idx val="${obj._dataIndex}"/><c:order val="${obj._dataIndex}"/>`
	strXml +=
		'<c:tx><c:strRef>' +
		`<c:f>${sheetCellRef(nameCol, 1)}</c:f>` +
		'<c:strCache><c:ptCount val="1"/><c:pt idx="0">' +
		el('c:v', null, obj.name ?? '') +
		'</c:pt></c:strCache></c:strRef></c:tx>'
	// A surface series carries 3-D shape props; the surface itself is colored by value band, but the
	// per-series fill still styles the wireframe / legend key.
	strXml += `<c:spPr>${genXmlColorSelection(seriesColor)}<a:ln/><a:effectLst/><a:sp3d/></c:spPr>`
	strXml += surfaceCatVal(obj, valFmtCode)
	strXml += '</c:ser>'
	return strXml
}

/**
 * Plot a surface chart into a `<c:surface3DChart>` (3-D) or `<c:surfaceChart>` (2-D contour). Every
 * series shares the category axis; the series axis (Z) is the third axis. `valAxisId`/`catAxisId`
 * are the primary ids passed by the dispatch; the series axis uses `AXIS_ID_SERIES_PRIMARY`.
 */
export function makeSurfacePlot(
	_chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string,
	valFmtCode: string
): string {
	const tag = isSurface3D(opts) ? 'surface3DChart' : 'surfaceChart'
	const chartColors = opts.chartColors?.length ? opts.chartColors : BARCHART_COLORS
	let strXml = `<c:${tag}>`
	strXml += `<c:wireframe val="${opts.surfaceWireframe ? 1 : 0}"/>`
	data.forEach((obj, idx) => {
		strXml += makeSurfaceSer(obj, valFmtCode, chartColors[idx % chartColors.length] ?? '4472C4')
	})
	// Surface, value and series axes (category X, value Y/height, series Z).
	strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/><c:axId val="${AXIS_ID_SERIES_PRIMARY}"/>`
	strXml += `</c:${tag}>`
	return strXml
}

/**
 * Build the `<c:view3D>` + floor/side/back walls for a surface chart. A 3-D surface tilts the scene
 * (using the shared `v3DRotX`/`v3DRotY` options, which the define layer defaults to 30); a 2-D
 * contour looks straight down the value axis (rotX 90, flat perspective). These precede
 * `<c:plotArea>` in CT_Chart document order (view3D → floor → sideWall → backWall → plotArea).
 */
export function makeSurfaceScene(opts: ChartOptsInternal): string {
	const wall = '<c:thickness val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr>'
	let strXml = '<c:view3D>'
	if (isSurface3D(opts)) {
		strXml += `<c:rotX val="${opts.v3DRotX}"/><c:rotY val="${opts.v3DRotY}"/><c:rAngAx val="0"/>`
	} else {
		// Contour / top view: look straight down the value axis, flat perspective.
		strXml += '<c:rotX val="90"/><c:rotY val="0"/><c:rAngAx val="0"/><c:perspective val="0"/>'
	}
	strXml += '</c:view3D>'
	strXml += `<c:floor>${wall}</c:floor><c:sideWall>${wall}</c:sideWall><c:backWall>${wall}</c:backWall>`
	return strXml
}
