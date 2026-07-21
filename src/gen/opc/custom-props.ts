/**
 * PptxGenJS: `docProps/custom.xml`
 *
 * Emit the custom-properties part from caller-supplied name/value pairs (bool,
 * date, number and string value types).
 */

import { CRLF, XML_DECL } from '../../core-enums-internal.js'
import type { CustomPropertyValue } from '../../core-interfaces.js'
import { el, raw } from '../oxml/el.js'

const CUSTOM_PROPS_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

const PROPS_NS = {
	xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
	'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
}

/** Serialize one custom-property value to its typed `vt:` element. */
function valueEl(value: CustomPropertyValue): string {
	if (typeof value === 'boolean') return el('vt:bool', null, String(value))
	if (value instanceof Date) return el('vt:filetime', null, value.toISOString().replace(/\.\d{3}Z$/, 'Z'))
	if (typeof value === 'number') return el(Number.isInteger(value) ? 'vt:i4' : 'vt:r8', null, value)
	return el('vt:lpwstr', null, String(value))
}

/**
 * Creates `docProps/custom.xml`
 * @param props - custom property name/value pairs
 * @returns XML
 */
export function makeXmlCustomProperties(props: Array<{ name: string; value: CustomPropertyValue }>): string {
	// pid is 1-based with 1 reserved by the spec, so caller properties start at 2.
	const properties = props.map(({ name, value }, idx) =>
		raw(el('property', { fmtid: CUSTOM_PROPS_FMTID, pid: idx + 2, name }, raw(valueEl(value))))
	)
	return XML_DECL + CRLF + el('Properties', PROPS_NS, properties)
}
