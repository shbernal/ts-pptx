/**
 * ts-pptx: `docProps/core.xml`
 *
 * Emit the core-properties part (title/subject/creator/revision and the
 * created/modified timestamps).
 */

import { XML_DECL } from '../../constants-internal.js'
import { el, raw } from '../oxml/el.js'

/** Each property sits on its own indented line; the parent supplies the closing indent. */
const PROP = { openPrefix: '\n\t\t' }

/**
 * The five namespaces a `docProps/core.xml` declares, in the order Office writes them.
 * Shared with the embedded workbook a chart carries (`gen/chart/embed-xlsx.ts`), which is a
 * package of its own and so has a core-properties part of its own.
 */
export const CORE_PROPS_NS = {
	'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
	'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
	'xmlns:dcterms': 'http://purl.org/dc/terms/',
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
				// `revision` is interpolated unescaped today; raw() preserves that.
				raw(el('cp:revision', null, raw(revision), PROP)),
				raw(el('dcterms:created', dcterms, raw(now), PROP)),
				raw(el('dcterms:modified', dcterms, raw(now), PROP)),
			],
			{ openPrefix: '\n\t', closePrefix: '\n\t' }
		)
	)
}
