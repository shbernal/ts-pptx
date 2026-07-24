/**
 * ts-pptx: slide-layout part
 *
 * Emit a slide layout (`ppt/slideLayouts/slideLayoutN.xml`) from a layout object.
 */

import { XML_DECL } from '../../core-enums-internal.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { slideObjectToXml } from './object.js'

/**
 * Generates the XML layout resource from a layout object
 * @param {SlideLayoutInternal} layout - slide layout (master)
 * @return {string} XML
 */
export function makeXmlLayout(layout: SlideLayoutInternal): string {
	return `${XML_DECL}
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">
		${slideObjectToXml(layout)}
		<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
}
