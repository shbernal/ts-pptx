/**
 * ts-pptx: `_rels/.rels`
 *
 * Emit the package root relationships (app/core/presentation, plus custom-props
 * when present).
 */

import { relationshipEl, relationshipsPart } from './rels.js'
import { CORE_PROPS_REL, CUSTOM_PROPS_REL, EXTENDED_PROPS_REL, OFFICE_DOCUMENT_REL } from '../../ooxml/rel-types.js'

/** Each entry carries its own indent; the `<Relationships>` wrapper adds none of its own. */
const INDENTED = { fmt: { openPrefix: '\n\t\t' } }

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels(hasCustomProps?: boolean): string {
	const rels = [
		relationshipEl('rId1', EXTENDED_PROPS_REL, 'docProps/app.xml', INDENTED),
		relationshipEl('rId2', CORE_PROPS_REL, 'docProps/core.xml', INDENTED),
		relationshipEl('rId3', OFFICE_DOCUMENT_REL, 'ppt/presentation.xml', INDENTED),
	]
	if (hasCustomProps) {
		rels.push(relationshipEl('rId4', CUSTOM_PROPS_REL, 'docProps/custom.xml', INDENTED))
	}
	// The closing tag is indented to child depth, not parent depth. That is
	// how this part has always been emitted; kept verbatim for byte-identity.
	return relationshipsPart(rels, { closePrefix: '\n\t\t' })
}
