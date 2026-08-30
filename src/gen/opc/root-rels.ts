/**
 * ts-pptx: `_rels/.rels`
 *
 * Emit the package root relationships (app/core/presentation, plus custom-props
 * when present).
 */

import { relationshipEl, relationshipsPart } from './rels.js'
import { OFFICE_REL, PACKAGE_REL_NS } from '../../ooxml/rel-types.js'

/** Each entry carries its own indent; the `<Relationships>` wrapper adds none of its own. */
const INDENTED = { fmt: { openPrefix: '\n\t\t' } }

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels(hasCustomProps?: boolean): string {
	const rels = [
		relationshipEl('rId1', OFFICE_REL + 'extended-properties', 'docProps/app.xml', INDENTED),
		relationshipEl('rId2', PACKAGE_REL_NS + '/metadata/core-properties', 'docProps/core.xml', INDENTED),
		relationshipEl('rId3', OFFICE_REL + 'officeDocument', 'ppt/presentation.xml', INDENTED),
	]
	if (hasCustomProps) {
		rels.push(relationshipEl('rId4', OFFICE_REL + 'custom-properties', 'docProps/custom.xml', INDENTED))
	}
	// The closing tag is indented to child depth, not parent depth. That is
	// how this part has always been emitted; kept verbatim for byte-identity.
	return relationshipsPart(rels, { closePrefix: '\n\t\t' })
}
