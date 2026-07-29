/**
 * ts-pptx: connector slide-object serialization
 *
 * Emits a `connector` slide object as a `<p:cxnSp>`. The start/end shape bindings resolve a
 * caller-supplied `objectName` to the target's slide-wide `<p:cNvPr>` id; an unresolved name
 * falls back to the static endpoint geometry rather than emitting a dangling id.
 */

import type { SlideObject } from '../../../types/internal.js'
import { warn } from '../../../diagnostics.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import { resolveObjectNameToId } from '../shape-ids.js'
import { cNvPrOpen, genXmlShapeLine } from './shared.js'

/**
 * Render a `connector` slide object to its `<p:cxnSp>` XML (start/end shape bindings via shapeIds).
 */
export function renderConnectorObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	locationAttrs: XmlAttrs,
	shapeIds: Map<SlideObject, number>
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// A connector is emitted as <p:cxnSp> (a connector shape) rather than <p:sp>, so
	// PowerPoint treats it as a connector. Geometry/flip come from the shared resolution
	// above; the preset (straightConnector1 / bentConnector3 / curvedConnector3) is on `shape`.
	strSlideXml += '<p:cxnSp><p:nvCxnSpPr>'
	strSlideXml += cNvPrOpen(idx + 2, slideItemObj.options.objectName, slideItemObj.options.altText || '') + '/>'
	{
		// Shape binding: resolve each bound target's objectName to its cNvPr id and emit
		// <a:stCxn>/<a:endCxn> in schema order. Resolution goes through `shapeIds`, so a shape
		// inside a group binds like any other (it is cNvPr-named on this slide); the old
		// `_slideObjects`-only lookup missed those and warned that the shape did not exist.
		// An unresolved name falls back to the static endpoint geometry (warn, don't corrupt)
		// rather than a dangling id.
		const cxnTag = (binding: { name: string; idx: number } | undefined, tag: 'a:stCxn' | 'a:endCxn'): string => {
			if (!binding) return ''
			const id = resolveObjectNameToId(shapeIds, binding.name)
			if (id === null) {
				warn(
					'connector/unresolved-binding',
					`addConnector could not bind to shape "${binding.name}" (no object with that objectName on the slide); using endpoint coordinates instead.`
				)
				return ''
			}
			return `<${tag} id="${id}" idx="${binding.idx}"/>`
		}
		const cxnSpPr = cxnTag(slideItemObj.options._startCxn, 'a:stCxn') + cxnTag(slideItemObj.options._endCxn, 'a:endCxn')
		strSlideXml += cxnSpPr ? `<p:cNvCxnSpPr>${cxnSpPr}</p:cNvCxnSpPr>` : '<p:cNvCxnSpPr/>'
	}
	strSlideXml += '<p:nvPr/></p:nvCxnSpPr><p:spPr>'
	strSlideXml += el('a:xfrm', locationAttrs, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])
	{
		// Bent/curved connectors carry adjustable jogs as `<a:gd name="adjN" fmla="val …"/>`
		// (1000ths-of-a-percent). With none, the empty `<a:avLst/>` leaves the preset default (50%).
		const adj = slideItemObj.options._connectorAdj || []
		const avLst = adj.map((val, i) => voidEl('a:gd', { name: `adj${i + 1}`, fmla: `val ${val}` })).join('')
		strSlideXml += el('a:prstGeom', { prst: slideItemObj.shape }, raw(el('a:avLst', null, raw(avLst))))
	}
	strSlideXml += genXmlShapeLine(slideItemObj.options.line || {})
	strSlideXml += '</p:spPr></p:cxnSp>'
	return strSlideXml
}
