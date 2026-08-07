/**
 * ts-pptx: `_rels/.rels`
 *
 * Emit the package root relationships (app/core/presentation, plus custom-props
 * when present).
 */

import { CRLF, XML_DECL } from '../../constants-internal.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { OFFICE_REL, PACKAGE_REL_NS } from '../oxml/schema-uris.js'

function relationship(id: string, type: string, target: string): string {
	return voidEl('Relationship', { Id: id, Type: type, Target: target }, { openPrefix: '\n\t\t' })
}

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels(hasCustomProps?: boolean): string {
	return (
		XML_DECL +
		CRLF +
		el(
			'Relationships',
			{ xmlns: PACKAGE_REL_NS },
			[
				raw(relationship('rId1', OFFICE_REL + 'extended-properties', 'docProps/app.xml')),
				raw(relationship('rId2', PACKAGE_REL_NS + '/metadata/core-properties', 'docProps/core.xml')),
				raw(relationship('rId3', OFFICE_REL + 'officeDocument', 'ppt/presentation.xml')),
				hasCustomProps ? raw(relationship('rId4', OFFICE_REL + 'custom-properties', 'docProps/custom.xml')) : null,
			],
			// The closing tag is indented to child depth, not parent depth. That is
			// how this part has always been emitted; kept verbatim for byte-identity.
			{ closePrefix: '\n\t\t' }
		)
	)
}
