/**
 * ts-pptx: `docProps/core.xml`
 *
 * Emit the core-properties part (title/subject/creator/revision and the
 * created/modified timestamps).
 */

import { XML_DECL } from '../../constants-internal.js'
import { el, raw } from '../oxml/el.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { warn } from '../../diagnostics.js'

/** Each property sits on its own indented line; the parent supplies the closing indent. */
const PROP = { openPrefix: '\n\t\t' }

/**
 * The five namespaces a `docProps/core.xml` declares, in the order Office writes them.
 * Shared with the embedded workbook a chart carries (`gen/chart/embed-xlsx.ts`), which is a
 * package of its own and so has a core-properties part of its own.
 */
export const CORE_PROPS_NS = {
	'xmlns:cp': OOXML_NS.cp,
	'xmlns:dc': OOXML_NS.dc,
	'xmlns:dcterms': OOXML_NS.dcterms,
	'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
	'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
}

/**
 * Now, as `dcterms:W3CDTF` without milliseconds — the precision Office writes.
 *
 * Call it **once** per part and reuse the value for `created` and `modified`: two calls make a
 * part whose two timestamps can disagree by a millisecond depending on when the build crossed a
 * tick, which is a difference no reader cares about and every byte-diff does.
 */
export function coreTimestamp(): string {
	return new Date().toISOString().replace(/\.\d\d\dZ/, 'Z')
}

/**
 * `cp:revision` as it is written: a whole number, the only form PowerPoint opens.
 *
 * Anything else warns and writes `1`, the revision a new deck starts at.
 */
function coreRevision(revision: string): string {
	if (/^\d+$/.test(revision)) return revision
	warn(
		'core/revision-not-a-whole-number',
		`revision ${JSON.stringify(revision)} is not a whole number, which PowerPoint cannot open; writing 1`
	)
	return '1'
}

/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} subject - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
export function makeXmlCore(title: string, subject: string, author: string, revision: string): string {
	const dcterms = { 'xsi:type': 'dcterms:W3CDTF' }
	// One reading of the clock for both stamps; see coreTimestamp.
	const now = coreTimestamp()
	return (
		XML_DECL +
		el(
			'cp:coreProperties',
			CORE_PROPS_NS,
			[
				raw(el('dc:title', null, title, PROP)),
				raw(el('dc:subject', null, subject, PROP)),
				raw(el('dc:creator', null, author, PROP)),
				raw(el('cp:lastModifiedBy', null, author, PROP)),
				// Escaped like every other property. It went in raw on the strength of the setter's note
				// that a revision is a whole number, which nothing enforced, so a `<` or `&` in one wrote a
				// `core.xml` that does not parse.
				raw(el('cp:revision', null, coreRevision(revision), PROP)),
				raw(el('dcterms:created', dcterms, raw(now), PROP)),
				raw(el('dcterms:modified', dcterms, raw(now), PROP)),
			],
			{ openPrefix: '\n\t', closePrefix: '\n\t' }
		)
	)
}
