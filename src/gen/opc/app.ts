/**
 * ts-pptx: `docProps/app.xml`
 *
 * Emit the extended-properties part (application, slide/notes counts, heading
 * pairs, titles of parts, company).
 */

import { CRLF, XML_DECL } from '../../core-enums-internal.js'
import type { PresSlideInternal } from '../../types/internal.js'
import { el, raw } from '../oxml/el.js'

/** This part is pretty-printed, one element per line, at three nesting depths. */
const INDENT_1 = '\n\t'
const INDENT_2 = '\n\t\t'
const INDENT_3 = '\n\t\t\t'
const PROP = { openPrefix: INDENT_1 }
const VECTOR = { openPrefix: INDENT_2, closePrefix: INDENT_2 }
const ITEM = { openPrefix: INDENT_3 }

const NS = {
	xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
	'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
}

/** One `<vt:variant>` of a heading pair: the section's name, then its count. */
function headingPair(name: string, count: number): string[] {
	return [
		el('vt:variant', null, raw(el('vt:lpstr', null, name)), ITEM),
		el('vt:variant', null, raw(el('vt:i4', null, count)), ITEM),
	]
}

/**
 * Creates `docProps/app.xml`
 * @param {PresSlideInternal[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
export function makeXmlApp(slides: PresSlideInternal[], company: string): string {
	const headingPairs = el(
		'vt:vector',
		{ size: 6, baseType: 'variant' },
		[...headingPair('Fonts Used', 2), ...headingPair('Theme', 1), ...headingPair('Slide Titles', slides.length)].map(
			raw
		),
		VECTOR
	)

	const titlesOfParts = el(
		'vt:vector',
		{ size: slides.length + 1 + 2, baseType: 'lpstr' },
		[
			raw(el('vt:lpstr', null, 'Arial', ITEM)),
			raw(el('vt:lpstr', null, 'Calibri', ITEM)),
			raw(el('vt:lpstr', null, 'Office Theme', ITEM)),
			// The slide titles share one line, so the indent is emitted once — and it is
			// emitted whether or not there are any titles to follow it.
			raw(INDENT_3),
			...slides.map((_slideObj, idx) => raw(el('vt:lpstr', null, `Slide ${idx + 1}`))),
		],
		VECTOR
	)

	return (
		XML_DECL +
		CRLF +
		el(
			'Properties',
			NS,
			[
				raw(el('TotalTime', null, 0, PROP)),
				raw(el('Words', null, 0, PROP)),
				raw(el('Application', null, 'Microsoft Office PowerPoint', PROP)),
				raw(el('PresentationFormat', null, 'On-screen Show (16:9)', PROP)),
				raw(el('Paragraphs', null, 0, PROP)),
				raw(el('Slides', null, slides.length, PROP)),
				raw(el('Notes', null, slides.length, PROP)),
				raw(el('HiddenSlides', null, 0, PROP)),
				raw(el('MMClips', null, 0, PROP)),
				raw(el('ScaleCrop', null, 'false', PROP)),
				raw(el('HeadingPairs', null, raw(headingPairs), { ...PROP, closePrefix: INDENT_1 })),
				raw(el('TitlesOfParts', null, raw(titlesOfParts), { ...PROP, closePrefix: INDENT_1 })),
				raw(el('Company', null, company, PROP)),
				raw(el('LinksUpToDate', null, 'false', PROP)),
				raw(el('SharedDoc', null, 'false', PROP)),
				raw(el('HyperlinksChanged', null, 'false', PROP)),
				raw(el('AppVersion', null, '16.0000', PROP)),
			],
			{ closePrefix: INDENT_1 }
		)
	)
}
