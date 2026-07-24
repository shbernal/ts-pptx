/**
 * ts-pptx: `[Content_Types].xml`
 *
 * Emit the package content-types part: Default entries for the media extensions
 * actually used by the deck (plus xlsx/font defaults when present) and Override
 * entries for every written part.
 */

import { CRLF, XML_DECL } from '../../core-enums-internal.js'
import type { PresSlideInternal, SlideLayoutInternal, SlideRelChart, SlideRelMedia } from '../../types/internal.js'
import { avContentType } from '../../media/content-type.js'
import { type EmbeddedFont, FONT_DATA_CONTENT_TYPE, FONT_DATA_EXTENSION } from '../../embedded-fonts.js'
import { el, raw, voidEl } from '../oxml/el.js'

/** Content-type prefixes; spelled out per part below so each entry stays greppable by its suffix. */
const OD = 'application/vnd.openxmlformats-officedocument.'
const PKG = 'application/vnd.openxmlformats-package.'
const CT_CHART = OD + 'drawingml.chart+xml'
// chartEx (cx:) parts use Microsoft content types, NOT the `OD` (openxmlformats) prefix. Each
// chartEx chart part also requires a chart-style + color-style sidecar part.
const CT_CHARTEX = 'application/vnd.ms-office.chartex+xml'
const CT_CHARTEX_STYLE = 'application/vnd.ms-office.chartstyle+xml'
const CT_CHARTEX_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml'
const CT_THEME = OD + 'theme+xml'

/**
 * Some Override entries have always been emitted with a leading space. It is insignificant
 * whitespace, but it is in the bytes this library has shipped for years, so it is reproduced
 * verbatim rather than normalized (see AGENTS.md "Verification": whitespace diffs are a STOP).
 */
const LEADING_SPACE = { openPrefix: ' ' }

function contentDefault(extension: string, contentType: string): string {
	return voidEl('Default', { Extension: extension, ContentType: contentType })
}

function override(partName: string, contentType: string, fmt?: { openPrefix: string }): string {
	return voidEl('Override', { PartName: partName, ContentType: contentType }, fmt)
}

/**
 * Content-type Override(s) for a chart rel. A classic chart is one Override; a chartEx chart is
 * three: the `chartEx{N}.xml` part plus its mandatory `style{N}.xml` and `colors{N}.xml` sidecars
 * (keyed to the same `globalId`).
 */
function chartOverrides(rel: SlideRelChart, fmt?: { openPrefix: string }): string[] {
	if (!rel.isChartEx) return [override(rel.Target, CT_CHART, fmt)]
	return [
		override(rel.Target, CT_CHARTEX, fmt),
		override(`/ppt/charts/style${rel.globalId}.xml`, CT_CHARTEX_STYLE, fmt),
		override(`/ppt/charts/colors${rel.globalId}.xml`, CT_CHARTEX_COLORS, fmt),
	]
}

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
	const parts: string[] = [contentDefault('xml', 'application/xml'), contentDefault('rels', PKG + 'relationships+xml')]

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
		parts.push(contentDefault(extn, type))
	})
	// Charts embed an xlsx workbook part; emit the Default only when at least one chart is present —
	// and only if an OLE object hasn't already contributed the same `xlsx` Default above (one
	// Extension may appear once).
	if (ctHasChart && !extnTypeMap.has('xlsx')) parts.push(contentDefault('xlsx', OD + 'spreadsheetml.sheet'))
	// Embedded fonts: one Default covers every `.fntdata` part (emitted only when fonts are embedded).
	if ((embeddedFonts || []).some((font) => font.faces.some((face) => face.bytes))) {
		parts.push(contentDefault(FONT_DATA_EXTENSION, FONT_DATA_CONTENT_TYPE))
	}

	// STEP 2: Add presentation and slide master(s)/slide(s)
	parts.push(override('/ppt/presentation.xml', OD + 'presentationml.presentation.main+xml'))
	parts.push(override('/ppt/notesMasters/notesMaster1.xml', OD + 'presentationml.notesMaster+xml'))
	// Only one slideMaster part (`slideMaster1.xml`) is written; emit a single matching Override
	// rather than one per slide (which would dangle, since `slideMaster2..N.xml` do not exist).
	parts.push(override('/ppt/slideMasters/slideMaster1.xml', OD + 'presentationml.slideMaster+xml'))
	slides.forEach((slide, idx) => {
		parts.push(override(`/ppt/slides/slide${idx + 1}.xml`, OD + 'presentationml.slide+xml'))
		// Add charts if any
		slide._relsChart.forEach((rel) => {
			parts.push(...chartOverrides(rel))
		})
	})

	// STEP 3: Core PPT
	parts.push(override('/ppt/presProps.xml', OD + 'presentationml.presProps+xml'))
	parts.push(override('/ppt/viewProps.xml', OD + 'presentationml.viewProps+xml'))
	parts.push(override('/ppt/theme/theme1.xml', CT_THEME))
	// notesMaster1.xml.rels references ../theme/theme2.xml; emit a matching Override so the part resolves
	parts.push(override('/ppt/theme/theme2.xml', CT_THEME))
	parts.push(override('/ppt/tableStyles.xml', OD + 'presentationml.tableStyles+xml'))

	// STEP 4: Add Slide Layouts
	slideLayouts.forEach((layout, idx) => {
		parts.push(override(`/ppt/slideLayouts/slideLayout${idx + 1}.xml`, OD + 'presentationml.slideLayout+xml'))
		;(layout._relsChart || []).forEach((rel) => {
			parts.push(...chartOverrides(rel, LEADING_SPACE))
		})
	})

	// STEP 5: Add notes slide(s)
	slides.forEach((_slide, idx) => {
		parts.push(override(`/ppt/notesSlides/notesSlide${idx + 1}.xml`, OD + 'presentationml.notesSlide+xml'))
	})

	// STEP 5b: Comments — per-slide comment part Override for slides that have comments, plus the
	// single presentation-level commentAuthors part Override when the deck has any comments.
	let hasAnyComment = false
	slides.forEach((slide, idx) => {
		if ((slide._comments || []).length > 0) {
			hasAnyComment = true
			parts.push(override(`/ppt/comments/comment${idx + 1}.xml`, OD + 'presentationml.comments+xml'))
		}
	})
	if (hasAnyComment) parts.push(override('/ppt/commentAuthors.xml', OD + 'presentationml.commentAuthors+xml'))

	// STEP 6: Add rels
	masterSlide?._relsChart.forEach((rel) => {
		parts.push(...chartOverrides(rel, LEADING_SPACE))
	})
	// master _relsMedia extensions are already covered by the unified ctTargets walk above; no per-master Default block needed here.

	// LAST: Finish XML (Resume core)
	parts.push(override('/docProps/core.xml', PKG + 'core-properties+xml', LEADING_SPACE))
	parts.push(override('/docProps/app.xml', OD + 'extended-properties+xml', LEADING_SPACE))
	if (hasCustomProps) parts.push(override('/docProps/custom.xml', OD + 'custom-properties+xml', LEADING_SPACE))

	return (
		XML_DECL +
		CRLF +
		el('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' }, parts.map(raw))
	)
}
