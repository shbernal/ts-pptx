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

const NS = {
	'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
	'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
	'xmlns:dcterms': 'http://purl.org/dc/terms/',
	'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
	'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
}

function timestamp(): string {
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
	return (
		XML_DECL +
		el(
			'cp:coreProperties',
			NS,
			[
				raw(el('dc:title', null, title, PROP)),
				raw(el('dc:subject', null, subject, PROP)),
				raw(el('dc:creator', null, author, PROP)),
				raw(el('cp:lastModifiedBy', null, author, PROP)),
				// `revision` is interpolated unescaped today; raw() preserves that.
				raw(el('cp:revision', null, raw(revision), PROP)),
				raw(el('dcterms:created', dcterms, raw(timestamp()), PROP)),
				raw(el('dcterms:modified', dcterms, raw(timestamp()), PROP)),
			],
			{ openPrefix: '\n\t', closePrefix: '\n\t' }
		)
	)
}
