/**
 * ts-pptx: `ppt/presentation.xml` + presProps/viewProps
 *
 * Emit the presentation part (masters/notes/slide id lists, sizes, embedded-font
 * list, default text styles, sections) and the small presProps/viewProps parts.
 */

import { CRLF, XML_DECL } from '../../core-enums-internal.js'
import type { PresentationPropsInternal, SectionInternalProps } from '../../types/internal.js'
import { flattenEmbeddedFaces, serializeEmbeddedFontLst } from '../../embedded-fonts.js'
import { presentationFontRelStart } from './presentation-rels.js'
import { el, raw, voidEl } from '../oxml/el.js'

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PML = 'http://schemas.openxmlformats.org/presentationml/2006/main'

function defaultTextStyleLevel(idy: number): string {
	return el(
		`a:lvl${idy}pPr`,
		{ marL: (idy - 1) * 457200, algn: 'l', defTabSz: 914400, rtl: 0, eaLnBrk: 1, latinLnBrk: 0, hangingPunct: 1 },
		raw(
			el('a:defRPr', { sz: 1800, kern: 1200 }, [
				raw(el('a:solidFill', null, raw(voidEl('a:schemeClr', { val: 'tx1' })))),
				raw(voidEl('a:latin', { typeface: '+mn-lt' })),
				raw(voidEl('a:ea', { typeface: '+mn-ea' })),
				raw(voidEl('a:cs', { typeface: '+mn-cs' })),
			])
		)
	)
}

function sectionsExtLst(sections: SectionInternalProps[]): string {
	return el('p:extLst', null, [
		raw(
			el(
				'p:ext',
				{ uri: '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}' },
				raw(
					el(
						'p14:sectionLst',
						{ 'xmlns:p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main' },
						sections.map((sect) =>
							raw(
								el(
									'p14:section',
									{ name: sect.title, id: sect._id },
									raw(
										el(
											'p14:sldIdLst',
											null,
											sect._slides.map((slide) => raw(voidEl('p14:sldId', { id: slide._slideId })))
										)
									)
								)
							)
						)
					)
				)
			)
		),
		raw(
			el(
				'p:ext',
				{ uri: '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}' },
				raw(voidEl('p15:sldGuideLst', { 'xmlns:p15': 'http://schemas.microsoft.com/office/powerpoint/2012/main' }))
			)
		),
	])
}

/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {PresentationPropsInternal} pres - presentation
 * @return {string} XML
 */
export function makeXmlPresentation(pres: PresentationPropsInternal): string {
	// NOTE: the double space before saveSubsetFonts/embedTrueTypeFonts when
	// rtlMode is unset is a pre-existing template artifact (verified against
	// the byte-identity baseline) — the literal space around `${rtl}` survives
	// even when the interpolation is empty. el()'s attrs always normalize to
	// single-space separators, so this opening tag stays a raw template rather
	// than being forced through the builder.
	const openTag =
		`<p:presentation xmlns:a="${DML}" xmlns:r="${REL}" ` +
		// When fonts are embedded we carry WHOLE faces, so `embedTrueTypeFonts="1"` (so
		// PowerPoint honors the embed) and `saveSubsetFonts="0"` (we did not subset).
		// With no embedded fonts, keep the historical inert `saveSubsetFonts="1"`.
		`xmlns:p="${PML}" ${pres.rtlMode ? 'rtl="1"' : ''} ${(pres.embeddedFonts || []).some((font) => font.faces.some((face) => face.bytes)) ? 'embedTrueTypeFonts="1" saveSubsetFonts="0"' : 'saveSubsetFonts="1"'} autoCompressPictures="0"${pres.firstSlideNum !== 1 ? ` firstSlideNum="${pres.firstSlideNum}"` : ''}>`

	// SPEC (ECMA-376 Part 1 §19.2.1.26): sldMasterIdLst, then notesMasterIdLst
	// BEFORE sldIdLst — emitting notesMasterIdLst after sldIdLst (or after
	// sldSz/notesSz) violates the schema and is flagged by OpenXmlValidator as
	// Sch_UnexpectedElementContentExpectingComplex.
	const sldMasterIdLst = el('p:sldMasterIdLst', null, raw(voidEl('p:sldMasterId', { id: 2147483648, 'r:id': 'rId1' })))

	// NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic)
	const notesMasterIdLst = el(
		'p:notesMasterIdLst',
		null,
		raw(voidEl('p:notesMasterId', { 'r:id': `rId${pres.slides.length + 2}` }))
	)

	const sldIdLst = el(
		'p:sldIdLst',
		null,
		pres.slides.map((slide) => raw(voidEl('p:sldId', { id: slide._slideId, 'r:id': `rId${slide._rId}` })))
	)

	const sldSz = voidEl('p:sldSz', { cx: pres.presLayout.width, cy: pres.presLayout.height })
	const notesSz = voidEl('p:notesSz', { cx: pres.presLayout.height, cy: pres.presLayout.width })

	// Embedded fonts (CT_Presentation index 7 — after notesSz, before defaultTextStyle).
	// rIds continue past the fixed presentation rels and must match makeXmlPresentationRels.
	const fonts = pres.embeddedFonts || []
	const flat = flattenEmbeddedFaces(fonts, presentationFontRelStart(pres.slides))
	const rIdOf = new Map(flat.map((face) => [`${face.fontIndex}:${face.slot}`, face.rId]))
	const embeddedFontLst = serializeEmbeddedFontLst(fonts, (fontIndex, slot) => rIdOf.get(`${fontIndex}:${slot}`))

	const defaultTextStyle = el(
		'p:defaultTextStyle',
		null,
		Array.from({ length: 9 }, (_, i) => raw(defaultTextStyleLevel(i + 1)))
	)

	const extLst = pres.sections && pres.sections.length > 0 ? sectionsExtLst(pres.sections) : ''

	return (
		XML_DECL +
		CRLF +
		openTag +
		sldMasterIdLst +
		notesMasterIdLst +
		sldIdLst +
		sldSz +
		notesSz +
		embeddedFontLst +
		defaultTextStyle +
		extLst +
		'</p:presentation>'
	)
}

/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export function makeXmlPresProps(): string {
	return XML_DECL + CRLF + voidEl('p:presentationPr', { 'xmlns:a': DML, 'xmlns:r': REL, 'xmlns:p': PML })
}

/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export function makeXmlViewProps(): string {
	return (
		XML_DECL +
		CRLF +
		el('p:viewPr', { 'xmlns:a': DML, 'xmlns:r': REL, 'xmlns:p': PML }, [
			raw(
				el('p:normalViewPr', { horzBarState: 'maximized' }, [
					raw(voidEl('p:restoredLeft', { sz: 15611 })),
					raw(voidEl('p:restoredTop', { sz: 94610 })),
				])
			),
			raw(
				el(
					'p:slideViewPr',
					null,
					raw(
						el('p:cSldViewPr', { snapToGrid: 0, snapToObjects: 1 }, [
							raw(
								el('p:cViewPr', { varScale: 1 }, [
									raw(
										el('p:scale', null, [
											raw(voidEl('a:sx', { n: 136, d: 100 })),
											raw(voidEl('a:sy', { n: 136, d: 100 })),
										])
									),
									raw(voidEl('p:origin', { x: 216, y: 312 })),
								])
							),
							raw(voidEl('p:guideLst')),
						])
					)
				)
			),
			raw(
				el(
					'p:notesTextViewPr',
					null,
					raw(
						el('p:cViewPr', null, [
							raw(el('p:scale', null, [raw(voidEl('a:sx', { n: 1, d: 1 })), raw(voidEl('a:sy', { n: 1, d: 1 }))])),
							raw(voidEl('p:origin', { x: 0, y: 0 })),
						])
					)
				)
			),
			raw(voidEl('p:gridSpacing', { cx: 76200, cy: 76200 })),
		])
	)
}
