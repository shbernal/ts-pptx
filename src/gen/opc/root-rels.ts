/**
 * PptxGenJS: `_rels/.rels`
 *
 * Emit the package root relationships (app/core/presentation, plus custom-props
 * when present).
 */

import { CRLF, XML_DECL } from '../../core-enums.js'

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels(hasCustomProps?: boolean): string {
	let xml = `${XML_DECL}${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>`
	if (hasCustomProps) {
		xml +=
			'\n\t\t<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
	}
	xml += '\n\t\t</Relationships>'
	return xml
}
