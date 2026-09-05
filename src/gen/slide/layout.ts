/**
 * ts-pptx: slide-layout part
 *
 * Emit a slide layout (`ppt/slideLayouts/slideLayoutN.xml`) from a layout object.
 */

import { XML_DECL } from '../../constants-internal.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { PML_ROOT_NS } from '../../ooxml/namespaces.js'
import { slideObjectToXml } from './object.js'
import type { RendererTable } from './objects/shared.js'

/**
 * Generates the XML layout resource from a layout object
 * @param layout - slide layout (master)
 * @param renderers - which renderer emits each shape family, passed through to the shape walk
 * @return XML
 */
export function makeXmlLayout(layout: SlideLayoutInternal, renderers: RendererTable): string {
	return (
		XML_DECL +
		el(
			'p:sldLayout',
			{ ...PML_ROOT_NS, preserve: '1' },
			[raw(slideObjectToXml(layout, renderers)), raw(el('p:clrMapOvr', null, raw(voidEl('a:masterClrMapping'))))],
			// The root and both children sit on their own two-tab line; the closing tag does not.
			{ openPrefix: '\n\t\t', childPrefix: '\n\t\t' }
		)
	)
}
