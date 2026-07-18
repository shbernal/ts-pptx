/**
 * PptxGenJS: `docProps/custom.xml`
 *
 * Emit the custom-properties part from caller-supplied name/value pairs (bool,
 * date, number and string value types).
 */

import { CRLF, XML_DECL } from '../../core-enums.js'
import type { CustomPropertyValue } from '../../core-interfaces.js'
import { encodeXmlEntities } from '../../gen-utils.js'

const CUSTOM_PROPS_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

/**
 * Creates `docProps/custom.xml`
 * @param props - custom property name/value pairs
 * @returns XML
 */
export function makeXmlCustomProperties(props: Array<{ name: string; value: CustomPropertyValue }>): string {
	const propertiesXml = props
		.map(({ name, value }, idx) => {
			let valueXml: string
			if (typeof value === 'boolean') {
				valueXml = `<vt:bool>${value}</vt:bool>`
			} else if (value instanceof Date) {
				valueXml = `<vt:filetime>${value.toISOString().replace(/\.\d{3}Z$/, 'Z')}</vt:filetime>`
			} else if (typeof value === 'number') {
				valueXml = Number.isInteger(value) ? `<vt:i4>${value}</vt:i4>` : `<vt:r8>${value}</vt:r8>`
			} else {
				valueXml = `<vt:lpwstr>${encodeXmlEntities(String(value))}</vt:lpwstr>`
			}
			return `<property fmtid="${CUSTOM_PROPS_FMTID}" pid="${idx + 2}" name="${encodeXmlEntities(name)}">${valueXml}</property>`
		})
		.join('')
	return `${XML_DECL}${CRLF}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${propertiesXml}</Properties>`
}
