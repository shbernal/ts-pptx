/**
 * PptxGenJS: `[Content_Types].xml`
 *
 * Emit the package content-types part: Default entries for the media extensions
 * actually used by the deck (plus xlsx/font defaults when present) and Override
 * entries for every written part.
 */

import { CRLF, XML_DECL } from '../../core-enums.js'
import type { PresSlideInternal, SlideLayoutInternal, SlideRelChart, SlideRelMedia } from '../../core-interfaces.js'
import { avContentType } from '../../gen-utils.js'
import { type EmbeddedFont, FONT_DATA_CONTENT_TYPE, FONT_DATA_EXTENSION } from '../../embedded-fonts.js'

/**
 * Generate XML ContentType
 * @param {PresSlideInternal[]} slides - slides
 * @param {SlideLayoutInternal[]} slideLayouts - slide layouts
 * @param {PresSlideInternal} masterSlide - master slide
 * @returns XML
 */
export function makeXmlContTypes(
	slides: PresSlideInternal[],
	slideLayouts: SlideLayoutInternal[],
	masterSlide?: PresSlideInternal,
	hasCustomProps?: boolean,
	embeddedFonts?: EmbeddedFont[]
): string {
	let strXml = XML_DECL + CRLF
	strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
	strXml += '<Default Extension="xml" ContentType="application/xml"/>'
	strXml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'

	// STEP 1 - Emit Default Extension entries only for media types actually used by the deck.
	// Walk slides + slideLayouts + masterSlide _relsMedia[] and dedupe by extension.
	// Skip 'online' rels (no part written) and rels missing extn/type.
	const extnTypeMap = new Map<string, string>()
	const ctTargets: Array<{ _relsMedia?: SlideRelMedia[]; _relsChart?: SlideRelChart[] }> = []
	;(slides || []).forEach((s) => ctTargets.push(s))
	;(slideLayouts || []).forEach((l) => ctTargets.push(l))
	if (masterSlide) ctTargets.push(masterSlide)
	let ctHasChart = false
	ctTargets.forEach((target) => {
		;(target._relsMedia || []).forEach((rel) => {
			if (rel.type === 'online' || !rel.extn || !rel.type) return
			// A/V rel `type` is `${mtype}/${extn}` (e.g. `audio/mp3`); resolve the part's
			// Default content type to what PowerPoint authors (`audio/mpeg`). Image rels
			// already carry their final content type (imageContentType).
			const contentType = rel.type.startsWith('audio/')
				? avContentType(rel.extn, 'audio')
				: rel.type.startsWith('video/')
					? avContentType(rel.extn, 'video')
					: rel.type
			if (!extnTypeMap.has(rel.extn)) extnTypeMap.set(rel.extn, contentType)
		})
		if ((target._relsChart || []).length > 0) ctHasChart = true
	})
	extnTypeMap.forEach((type, extn) => {
		strXml += '<Default Extension="' + extn + '" ContentType="' + type + '"/>'
	})
	// Charts embed an xlsx workbook part; emit the Default only when at least one chart is present.
	if (ctHasChart) {
		strXml +=
			'<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
	}
	// Embedded fonts: one Default covers every `.fntdata` part (emitted only when fonts are embedded).
	if ((embeddedFonts || []).some((font) => font.faces.some((face) => face.bytes))) {
		strXml += `<Default Extension="${FONT_DATA_EXTENSION}" ContentType="${FONT_DATA_CONTENT_TYPE}"/>`
	}

	// STEP 2: Add presentation and slide master(s)/slide(s)
	strXml +=
		'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
	strXml +=
		'<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
	// Only one slideMaster part (`slideMaster1.xml`) is written; emit a single matching Override
	// rather than one per slide (which would dangle, since `slideMaster2..N.xml` do not exist).
	strXml +=
		'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
	slides.forEach((slide, idx) => {
		strXml += `<Override PartName="/ppt/slides/slide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
		// Add charts if any
		slide._relsChart.forEach((rel) => {
			strXml += `<Override PartName="${rel.Target}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
		})
	})

	// STEP 3: Core PPT
	strXml +=
		'<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
	strXml +=
		'<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
	strXml +=
		'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
	// notesMaster1.xml.rels references ../theme/theme2.xml; emit a matching Override so the part resolves
	strXml +=
		'<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
	strXml +=
		'<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'

	// STEP 4: Add Slide Layouts
	slideLayouts.forEach((layout, idx) => {
		strXml += `<Override PartName="/ppt/slideLayouts/slideLayout${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`
		;(layout._relsChart || []).forEach((rel) => {
			strXml +=
				' <Override PartName="' +
				rel.Target +
				'" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
		})
	})

	// STEP 5: Add notes slide(s)
	slides.forEach((_slide, idx) => {
		strXml += `<Override PartName="/ppt/notesSlides/notesSlide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
	})

	// STEP 5b: Comments — per-slide comment part Override for slides that have comments, plus the
	// single presentation-level commentAuthors part Override when the deck has any comments.
	let hasAnyComment = false
	slides.forEach((slide, idx) => {
		if ((slide._comments || []).length > 0) {
			hasAnyComment = true
			strXml += `<Override PartName="/ppt/comments/comment${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>`
		}
	})
	if (hasAnyComment) {
		strXml +=
			'<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>'
	}

	// STEP 6: Add rels
	masterSlide?._relsChart.forEach((rel) => {
		strXml +=
			' <Override PartName="' +
			rel.Target +
			'" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
	})
	// master _relsMedia extensions are already covered by the unified ctTargets walk above; no per-master Default block needed here.

	// LAST: Finish XML (Resume core)
	strXml +=
		' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
	strXml +=
		' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
	if (hasCustomProps) {
		strXml +=
			' <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
	}
	strXml += '</Types>'

	return strXml
}
