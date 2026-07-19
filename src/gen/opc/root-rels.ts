/**
 * PptxGenJS: `_rels/.rels`
 *
 * Emit the package root relationships (app/core/presentation, plus custom-props
 * when present).
 */

import { CRLF, XML_DECL } from '../../core-enums.js'
import { el, raw, voidEl } from '../oxml/el.js'

const SCHEMA_BASE = 'http://schemas.openxmlformats.org/'

function relationship(id: string, type: string, target: string): string {
	return voidEl('Relationship', { Id: id, Type: SCHEMA_BASE + type, Target: target }, { openPrefix: '\n\t\t' })
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
			{ xmlns: SCHEMA_BASE + 'package/2006/relationships' },
			[
				raw(relationship('rId1', 'officeDocument/2006/relationships/extended-properties', 'docProps/app.xml')),
				raw(relationship('rId2', 'package/2006/relationships/metadata/core-properties', 'docProps/core.xml')),
				raw(relationship('rId3', 'officeDocument/2006/relationships/officeDocument', 'ppt/presentation.xml')),
				hasCustomProps
					? raw(relationship('rId4', 'officeDocument/2006/relationships/custom-properties', 'docProps/custom.xml'))
					: null,
			],
			// The closing tag is indented to child depth, not parent depth. That is
			// how this part has always been emitted; kept verbatim for byte-identity.
			{ closePrefix: '\n\t\t' }
		)
	)
}
