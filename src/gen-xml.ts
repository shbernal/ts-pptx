/**
 * PptxGenJS: XML Generation
 *
 * The bulk of the OOXML emitter: turns the in-memory slide/presentation model into
 * the many XML parts of a `.pptx` package (slides, masters, layouts, notes, theme,
 * `[Content_Types].xml`, `.rels`, app/core/custom props, table styles, view props).
 * Every export is a pure string builder — no I/O; the `pptxgen.ts` export flow calls
 * these and hands the strings to the zip writer. Charts live in `gen-charts.ts`.
 *
 * Contents — jump by grepping the `// ===== <region> =====` banners:
 *   - Value clamps & shape/geometry helpers   font/spacing clamps, crop, preset & custom geometry, cell borders
 *   - Slide serialization                     slideObjectToXml (the per-shape spTree builder) + its rels
 *   - Text body generation                    paragraph/run props, runs, math, genXmlTextBody, placeholders
 *   - Package-level parts                     [Content_Types].xml, root rels, app/core/custom props, presentation rels
 *   - Transitions & animations                slide timing, transitions, the p:anim* sequence builders
 *   - Slides, notes & layouts                 makeXmlSlide / notes parts / makeXmlLayout
 *   - Masters & text styles                   master txStyles defaults + makeXmlMaster
 *   - Slide/master rels & comments            per-part .rels and the comment author/thread parts
 *   - Theme, presentation & root files        theme, makeXmlPresentation, presProps, table styles, viewProps
 */

import { CRLF, LAYOUT_IDX_SERIES_BASE, REGEX_HEX_COLOR, TableStyle, XML_DECL } from './core-enums.js'
import type {
	BorderProps,
	CustomPropertyValue,
	PresentationPropsInternal,
	SlideRelChart,
	SlideRelMedia,
	MasterBulletProps,
	MasterTextStyleLevel,
	MasterTextStyleProps,
	PresSlideInternal,
	SlideLayoutInternal,
	TableStyleInternal,
	TableStyleRegionProps,
	ThemeColorScheme,
} from './core-interfaces.js'
import {
	avContentType,
	createColorElement,
	encodeXmlEntities,
	genXmlColorSelection,
	getUuid,
	inch2Emu,
	lineWidthToEmu,
	resolveBorderWidth,
} from './gen-utils.js'
import { HUNDREDTHS_PER_POINT, ptToHundredths } from './units.js'
import { slideTimingToXml } from './gen/anim/timing.js'
import { slideTransitionToXml } from './gen/anim/transition.js'
import { slideObjectRelationsToXml, slideObjectToXml } from './gen/slide/object.js'
// Notes parts live in gen/slide/notes.ts; re-exported so `genXml.*` keeps resolving in pptxgen.ts.
export {
	buildNotesSlideRels,
	getNotesFromSlide,
	makeXmlNotesMaster,
	makeXmlNotesMasterRel,
	makeXmlNotesSlide,
	makeXmlNotesSlideRel,
} from './gen/slide/notes.js'
// Comment parts live in gen/slide/comments.ts; re-exported for pptxgen.ts's `genXml.*` access.
export { makeXmlCommentAuthors, makeXmlComments, resolveCommentAuthors } from './gen/slide/comments.js'
export type { ResolvedComments } from './gen/slide/comments.js'
import { warn } from './log.js'
import {
	type EmbeddedFont,
	FONT_DATA_CONTENT_TYPE,
	FONT_DATA_EXTENSION,
	FONT_REL_TYPE,
	flattenEmbeddedFaces,
	serializeEmbeddedFontLst,
} from './embedded-fonts.js'

// XML-GEN: First 6 functions create the base /ppt files

// ===== Package-level parts =====

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

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels(hasCustomProps?: boolean): string {
	let xml = `${XML_DECL}${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>`
	if (hasCustomProps) {
		xml +=
			'\n\t\t<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
	}
	xml += '\n\t\t</Relationships>'
	return xml
}

/**
 * Creates `docProps/app.xml`
 * @param {PresSlideInternal[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
export function makeXmlApp(slides: PresSlideInternal[], company: string): string {
	return `${XML_DECL}${CRLF}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
	<TotalTime>0</TotalTime>
	<Words>0</Words>
	<Application>Microsoft Office PowerPoint</Application>
	<PresentationFormat>On-screen Show (16:9)</PresentationFormat>
	<Paragraphs>0</Paragraphs>
	<Slides>${slides.length}</Slides>
	<Notes>${slides.length}</Notes>
	<HiddenSlides>0</HiddenSlides>
	<MMClips>0</MMClips>
	<ScaleCrop>false</ScaleCrop>
	<HeadingPairs>
		<vt:vector size="6" baseType="variant">
			<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>2</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>1</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>${slides.length}</vt:i4></vt:variant>
		</vt:vector>
	</HeadingPairs>
	<TitlesOfParts>
		<vt:vector size="${slides.length + 1 + 2}" baseType="lpstr">
			<vt:lpstr>Arial</vt:lpstr>
			<vt:lpstr>Calibri</vt:lpstr>
			<vt:lpstr>Office Theme</vt:lpstr>
			${slides.map((_slideObj, idx) => `<vt:lpstr>Slide ${idx + 1}</vt:lpstr>`).join('')}
		</vt:vector>
	</TitlesOfParts>
	<Company>${encodeXmlEntities(company)}</Company>
	<LinksUpToDate>false</LinksUpToDate>
	<SharedDoc>false</SharedDoc>
	<HyperlinksChanged>false</HyperlinksChanged>
	<AppVersion>16.0000</AppVersion>
	</Properties>`
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
	return `${XML_DECL}
	<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
		<dc:title>${encodeXmlEntities(title)}</dc:title>
		<dc:subject>${encodeXmlEntities(subject)}</dc:subject>
		<dc:creator>${encodeXmlEntities(author)}</dc:creator>
		<cp:lastModifiedBy>${encodeXmlEntities(author)}</cp:lastModifiedBy>
		<cp:revision>${revision}</cp:revision>
		<dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d\d\dZ/, 'Z')}</dcterms:created>
		<dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d\d\dZ/, 'Z')}</dcterms:modified>
	</cp:coreProperties>`
}

const CUSTOM_PROPS_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

/**
 * Creates `docProps/custom.xml`
 * @param props - custom property name/value pairs
 * @returns XML
 */
export function makeXmlCustomProperties(props: Array<{ name: string; value: CustomPropertyValue }>): string {
	const propertiesXml = props
		.map(({ name, value }, idx) => {
			let valueXml: string
			if (typeof value === 'boolean') {
				valueXml = `<vt:bool>${value}</vt:bool>`
			} else if (value instanceof Date) {
				valueXml = `<vt:filetime>${value.toISOString().replace(/\.\d{3}Z$/, 'Z')}</vt:filetime>`
			} else if (typeof value === 'number') {
				valueXml = Number.isInteger(value) ? `<vt:i4>${value}</vt:i4>` : `<vt:r8>${value}</vt:r8>`
			} else {
				valueXml = `<vt:lpwstr>${encodeXmlEntities(String(value))}</vt:lpwstr>`
			}
			return `<property fmtid="${CUSTOM_PROPS_FMTID}" pid="${idx + 2}" name="${encodeXmlEntities(name)}">${valueXml}</property>`
		})
		.join('')
	return `${XML_DECL}${CRLF}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${propertiesXml}</Properties>`
}

/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {PresSlideInternal[]} slides - Presenation Slides
 * @returns XML
 */
/**
 * The first relationship id free for embedded-font rels in `presentation.xml.rels`,
 * i.e. one past the last fixed rel {@link makeXmlPresentationRels} emits. Shared by
 * the rels writer and {@link makeXmlPresentation} so the `embeddedFontLst` face
 * `r:id`s match the relationships that back them.
 *
 * Layout: rId1 = slideMaster, rId2..(N+1) = N slides, then notesMaster/presProps/
 * viewProps/theme1/tableStyles (5), then commentAuthors (1, only with comments).
 */
function presentationFontRelStart(slides: PresSlideInternal[]): number {
	const hasComments = (slides || []).some((slide) => (slide._comments || []).length > 0)
	return slides.length + 7 + (hasComments ? 1 : 0)
}

export function makeXmlPresentationRels(slides: PresSlideInternal[], embeddedFonts?: EmbeddedFont[]): string {
	let intRelNum = 1
	let strXml = XML_DECL + CRLF
	strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
	strXml +=
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
	for (let idx = 1; idx <= slides.length; idx++) {
		strXml += `<Relationship Id="rId${++intRelNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${idx}.xml"/>`
	}
	intRelNum++
	strXml +=
		`<Relationship Id="rId${intRelNum + 0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>` +
		`<Relationship Id="rId${intRelNum + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>` +
		`<Relationship Id="rId${intRelNum + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
		`<Relationship Id="rId${intRelNum + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
		`<Relationship Id="rId${intRelNum + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`
	// The presentation-level commentAuthors part is shared by every slide's comments, so it is
	// related once from the presentation (only when the deck has at least one comment).
	if ((slides || []).some((slide) => (slide._comments || []).length > 0)) {
		strXml += `<Relationship Id="rId${intRelNum + 5}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>`
	}
	// Embedded fonts: one `font` rel per face, ids continuing past the fixed rels above.
	for (const face of flattenEmbeddedFaces(embeddedFonts || [], presentationFontRelStart(slides))) {
		strXml += `<Relationship Id="rId${face.rId}" Type="${FONT_REL_TYPE}" Target="fonts/font${face.partIndex}.fntdata"/>`
	}
	strXml += '</Relationships>'

	return strXml
}

// XML-GEN: Functions that run 1-N times (once for each Slide)

// ===== Slides, notes & layouts =====

/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {PresSlideInternal} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlSlide(slide: PresSlideInternal): string {
	return (
		`${XML_DECL}${CRLF}` +
		'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
		`${slide?.hidden ? ' show="0"' : ''}>` +
		`${slideObjectToXml(slide)}` +
		'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
		slideTransitionToXml(slide) +
		slideTimingToXml(slide) +
		'</p:sld>'
	)
}

/**
 * Generates the XML layout resource from a layout object
 * @param {SlideLayoutInternal} layout - slide layout (master)
 * @return {string} XML
 */
export function makeXmlLayout(layout: SlideLayoutInternal): string {
	return `${XML_DECL}
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">
		${slideObjectToXml(layout)}
		<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
}

// ===== Masters & text styles =====
// Default per-level values mirroring the built-in Office master (used as the base that
// `MasterTextStyleProps` overrides are layered onto). `bu` describes the level's default
// bullet: 'none' -> <a:buNone/>, undefined -> no bullet element (otherStyle), or a glyph.
interface MasterLevelDefault {
	marL: number // EMU
	indent?: number // EMU (omitted when undefined)
	algn: string // OOXML algn value
	spcBefPct?: number // <a:spcBef> percent (×1000); omitted when undefined
	bu?: 'none' | { char: string; font: string } // default bullet; undefined => emit no bullet element
	sz: number // <a:defRPr@sz>
	font: 'mj' | 'mn' // major (heading) vs minor (body) theme font family
}
const MASTER_TITLE_DEFAULT: MasterLevelDefault = {
	marL: 0,
	algn: 'ctr',
	spcBefPct: 0,
	bu: 'none',
	sz: 4400,
	font: 'mj',
}
const MASTER_BODY_DEFAULTS: MasterLevelDefault[] = [
	{
		marL: 342900,
		indent: -342900,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 3200,
		font: 'mn',
	},
	{
		marL: 742950,
		indent: -285750,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '–', font: 'Arial' },
		sz: 2800,
		font: 'mn',
	},
	{
		marL: 1143000,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2400,
		font: 'mn',
	},
	{
		marL: 1600200,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '–', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2057400,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '»', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2514600,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2971800,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 3429000,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 3886200,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
]
const MASTER_OTHER_DEFAULTS: MasterLevelDefault[] = [
	0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600,
].map((marL) => ({ marL, algn: 'l', sz: 1800, font: 'mn' as const }))

function masterAlignAttr(align: MasterTextStyleLevel['align']): string {
	switch (align) {
		case 'left':
			return 'l'
		case 'right':
			return 'r'
		case 'center':
			return 'ctr'
		case 'justify':
			return 'just'
		default:
			return ''
	}
}

/** Build the bullet element for a master level: caller override wins over the level default. */
function masterBulletXml(
	bulletOverride: boolean | MasterBulletProps | undefined,
	base: MasterLevelDefault['bu']
): string {
	// Explicit override
	if (bulletOverride === false) return '<a:buNone/>'
	if (bulletOverride && typeof bulletOverride === 'object') {
		const font = bulletOverride.fontFace ? `<a:buFont typeface="${encodeXmlEntities(bulletOverride.fontFace)}"/>` : ''
		if (bulletOverride.type === 'number') {
			const type = bulletOverride.numberType || 'arabicPeriod'
			const startAt =
				typeof bulletOverride.numberStartAt === 'number' ? ` startAt="${Math.round(bulletOverride.numberStartAt)}"` : ''
			return `${font}<a:buAutoNum type="${type}"${startAt}/>`
		}
		// character bullet (default)
		const char = bulletOverride.characterCode ? `&#x${bulletOverride.characterCode};` : '•'
		const buFont = bulletOverride.fontFace ? font : '<a:buFont typeface="Arial" pitchFamily="34" charset="0"/>'
		return `${buFont}<a:buChar char="${char}"/>`
	}
	// No override (undefined / true): keep the level's default bullet
	if (base === 'none') return '<a:buNone/>'
	if (base && typeof base === 'object')
		return `<a:buFont typeface="${base.font}" pitchFamily="34" charset="0"/><a:buChar char="${base.char}"/>`
	return '' // otherStyle: no bullet element by default
}

/** Serialize one `<a:lvlNpPr>` from its default, layering an optional caller override. */
function masterLevelXml(levelNum: number, base: MasterLevelDefault, levelOverride: MasterTextStyleLevel = {}): string {
	const marL =
		typeof levelOverride.marginLeft === 'number' && !isNaN(levelOverride.marginLeft)
			? inch2Emu(levelOverride.marginLeft)
			: base.marL
	const indentEmu =
		typeof levelOverride.indent === 'number' && !isNaN(levelOverride.indent)
			? inch2Emu(levelOverride.indent)
			: base.indent
	const algn = (levelOverride.align && masterAlignAttr(levelOverride.align)) || base.algn
	const indentAttr = typeof indentEmu === 'number' ? ` indent="${indentEmu}"` : ''

	let xml = `<a:lvl${levelNum}pPr marL="${marL}"${indentAttr} algn="${algn}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
	if (typeof base.spcBefPct === 'number') xml += `<a:spcBef><a:spcPct val="${base.spcBefPct}"/></a:spcBef>`
	xml += masterBulletXml(levelOverride.bullet, base.bu)

	// defRPr
	let sz = base.sz
	if (typeof levelOverride.fontSize === 'number') {
		if (isNaN(levelOverride.fontSize) || levelOverride.fontSize <= 0)
			warn(
				`master textStyles fontSize "${levelOverride.fontSize}" is invalid; keeping default ${base.sz / HUNDREDTHS_PER_POINT}pt.`
			)
		else sz = ptToHundredths(levelOverride.fontSize)
	}
	const boldAttr = levelOverride.bold ? ' b="1"' : ''
	const italicAttr = levelOverride.italic ? ' i="1"' : ''
	const colorXml = levelOverride.color ? createColorElement(levelOverride.color) : '<a:schemeClr val="tx1"/>'
	const latinXml = levelOverride.fontFace
		? `<a:latin typeface="${encodeXmlEntities(levelOverride.fontFace)}"/>`
		: `<a:latin typeface="+${base.font}-lt"/>`
	xml += `<a:defRPr sz="${sz}"${boldAttr}${italicAttr} kern="1200"><a:solidFill>${colorXml}</a:solidFill>${latinXml}<a:ea typeface="+${base.font}-ea"/><a:cs typeface="+${base.font}-cs"/></a:defRPr>`
	xml += `</a:lvl${levelNum}pPr>`
	return xml
}

/** Clamp a caller-provided per-level override array to the 9 valid list levels, warning on overflow. */
function masterLevelOverrides(levels: MasterTextStyleLevel[] | undefined, group: string): MasterTextStyleLevel[] {
	if (!Array.isArray(levels)) return []
	if (levels.length > 9) warn(`master textStyles.${group} has ${levels.length} levels; only the first 9 are used.`)
	return levels.slice(0, 9)
}

/**
 * Build the `<p:txStyles>` block from caller overrides layered onto the Office master defaults.
 * Only invoked when `defineSlideMaster({ textStyles })` was set; the unconfigured deck keeps the
 * verbatim default literal in `makeXmlMaster` for byte-identical output.
 */
function makeXmlMasterTxStyles(textStyles: MasterTextStyleProps): string {
	const title = masterLevelXml(1, MASTER_TITLE_DEFAULT, textStyles.title)
	const bodyOverrides = masterLevelOverrides(textStyles.body, 'body')
	const body = MASTER_BODY_DEFAULTS.map((base, i) => masterLevelXml(i + 1, base, bodyOverrides[i])).join('')
	const otherOverrides = masterLevelOverrides(textStyles.other, 'other')
	const other = MASTER_OTHER_DEFAULTS.map((base, i) => masterLevelXml(i + 1, base, otherOverrides[i])).join('')
	return (
		'<p:txStyles>' +
		`<p:titleStyle>${title}</p:titleStyle>` +
		`<p:bodyStyle>${body}</p:bodyStyle>` +
		`<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${other}</p:otherStyle>` +
		'</p:txStyles>'
	)
}

/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {PresSlideInternal} slide - slide object that represents master slide layout
 * @param {SlideLayoutInternal[]} layouts - slide layouts
 * @return {string} XML
 */
export function makeXmlMaster(slide: PresSlideInternal, layouts: SlideLayoutInternal[]): string {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	const layoutDefs = layouts.map(
		(_layoutDef, idx) =>
			`<p:sldLayoutId id="${LAYOUT_IDX_SERIES_BASE + idx}" r:id="rId${slide._rels.length + idx + 1}"/>`
	)

	let strXml = XML_DECL + CRLF
	strXml +=
		'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
	strXml += slideObjectToXml(slide)
	strXml +=
		'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
	strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>'
	// CT_HeaderFooter/@sldNum defaults to true (ECMA-376). When a slide-number placeholder is
	// defined on the master we must NOT disable it here, otherwise slides that PowerPoint inserts
	// from this master inherit sldNum="0" and the master slide number disappears.
	strXml += `<p:hf${slide._slideNumberProps ? '' : ' sldNum="0"'} hdr="0" ftr="0" dt="0"/>`
	strXml += slide._txStyles
		? makeXmlMasterTxStyles(slide._txStyles)
		: '<p:txStyles>' +
			' <p:titleStyle>' +
			'  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
			' </p:titleStyle>' +
			' <p:bodyStyle>' +
			'  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
			'  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
			'  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
			'  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
			'  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
			'  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
			'  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
			'  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
			'  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
			' </p:bodyStyle>' +
			' <p:otherStyle>' +
			'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
			'  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
			'  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
			'  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
			'  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
			'  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
			'  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
			'  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
			'  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
			'  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
			' </p:otherStyle>' +
			'</p:txStyles>'
	strXml += '</p:sldMaster>'

	return strXml
}

// ===== Slide/master rels & comments =====

/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlSlideLayoutRel(layoutNumber: number, slideLayouts: SlideLayoutInternal[]): string {
	const slideLayout = slideLayouts[layoutNumber - 1]
	if (!slideLayout) throw new Error(`makeXmlSlideLayoutRel: no slide layout at index ${layoutNumber - 1}`)
	return slideObjectRelationsToXml(slideLayout, [
		{
			target: '../slideMasters/slideMaster1.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
		},
	])
}

/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {PresSlideInternal[]} slides
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlSlideRel(
	slides: PresSlideInternal[],
	slideLayouts: SlideLayoutInternal[],
	slideNumber: number
): string {
	const slide = slides[slideNumber - 1]
	if (!slide) throw new Error(`makeXmlSlideRel: no slide at index ${slideNumber - 1}`)
	const defaultRels = [
		{
			target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
		},
		{
			target: `../notesSlides/notesSlide${slideNumber}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
		},
	]
	// Only emit the comments rel for slides that actually carry comments (the comment part
	// is likewise only written for those slides); the rId is assigned after slideLayout/notesSlide.
	if ((slide._comments || []).length > 0) {
		defaultRels.push({
			target: `../comments/comment${slideNumber}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
		})
	}
	return slideObjectRelationsToXml(slide, defaultRels)
}

/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {PresSlideInternal} masterSlide - Slide object
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel(masterSlide: PresSlideInternal, slideLayouts: SlideLayoutInternal[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
		target: `../slideLayouts/slideLayout${idx + 1}.xml`,
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
	}))
	defaultRels.push({
		target: '../theme/theme1.xml',
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
	})

	return slideObjectRelationsToXml(masterSlide, defaultRels)
}

// ===== Theme, presentation & root files =====

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {PresSlideInternal[]} slides - srray of slides
 * @param {SlideLayoutInternal[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide(
	slides: PresSlideInternal[],
	slideLayouts: SlideLayoutInternal[],
	slideNumber: number
): number {
	for (let i = 0; i < slideLayouts.length; i++) {
		if (slideLayouts[i]?._name === slides[slideNumber - 1]?._slideLayout?._name) {
			return i + 1
		}
	}

	// IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
	// So all objects are in Layout1 and every slide that references it uses this layout.
	return 1
}

// XML-GEN: Last 5 functions create root /ppt files

/**
 * Theme `<a:clrScheme>` slots in OOXML document order, with their default Office color child.
 * `dk1`/`lt1` default to `sysClr` (windowText/window); the rest are `srgbClr`. A user override
 * for any slot is emitted as `<a:srgbClr>` (see `buildThemeClrScheme`).
 */
const THEME_CLR_SCHEME_DEFAULTS: ReadonlyArray<[keyof ThemeColorScheme, string]> = [
	['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'],
	['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'],
	['dk2', '<a:srgbClr val="44546A"/>'],
	['lt2', '<a:srgbClr val="E7E6E6"/>'],
	['accent1', '<a:srgbClr val="4472C4"/>'],
	['accent2', '<a:srgbClr val="ED7D31"/>'],
	['accent3', '<a:srgbClr val="A5A5A5"/>'],
	['accent4', '<a:srgbClr val="FFC000"/>'],
	['accent5', '<a:srgbClr val="5B9BD5"/>'],
	['accent6', '<a:srgbClr val="70AD47"/>'],
	['hlink', '<a:srgbClr val="0563C1"/>'],
	['folHlink', '<a:srgbClr val="954F72"/>'],
]

/**
 * Build the theme `<a:clrScheme>` block, applying any caller-supplied color overrides over the
 * default Office scheme. Invalid (non 6-digit-hex) overrides warn and keep the default rather
 * than emitting a degenerate color.
 * @param {ThemeColorScheme} [scheme] - per-slot hex overrides
 * @return {string} the `<a:clrScheme>...</a:clrScheme>` XML
 */
function buildThemeClrScheme(scheme?: ThemeColorScheme): string {
	const slots = THEME_CLR_SCHEME_DEFAULTS.map(([slot, defaultChild]) => {
		const override = scheme?.[slot]
		let child = defaultChild
		if (typeof override === 'string' && override.length > 0) {
			const hex = override.replace('#', '')
			if (REGEX_HEX_COLOR.test(hex)) child = `<a:srgbClr val="${hex.toUpperCase()}"/>`
			else
				warn(`makeXmlTheme: colorScheme.${slot} "${override}" is not a 6-digit hex color; keeping the Office default.`)
		}
		return `<a:${slot}>${child}</a:${slot}>`
	}).join('')
	return `<a:clrScheme name="Office">${slots}</a:clrScheme>`
}

/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
export function makeXmlTheme(pres: PresentationPropsInternal): string {
	const majorFont = pres.theme?.headFontFace
		? `<a:latin typeface="${pres.theme?.headFontFace}"/>`
		: '<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
	const minorFont = pres.theme?.bodyFontFace
		? `<a:latin typeface="${pres.theme?.bodyFontFace}"/>`
		: '<a:latin typeface="Calibri" panose="020F0502020204030204"/>'
	// East Asian (`<a:ea>`) and complex-script (`<a:cs>`) theme font slots. PowerPoint emits these
	// empty by default and resolves per-script via the `<a:font>` list that follows; setting them
	// lets CJK / complex-script runs fall back to a caller-chosen theme font.
	const majorEa = `<a:ea typeface="${pres.theme?.headFontFaceEA ?? ''}"/>`
	const minorEa = `<a:ea typeface="${pres.theme?.bodyFontFaceEA ?? ''}"/>`
	const majorCs = `<a:cs typeface="${pres.theme?.headFontFaceCS ?? ''}"/>`
	const minorCs = `<a:cs typeface="${pres.theme?.bodyFontFaceCS ?? ''}"/>`
	return `${XML_DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>${buildThemeClrScheme(pres.theme?.colorScheme)}<a:fontScheme name="Office"><a:majorFont>${majorFont}${majorEa}${majorCs}<a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont>${minorFont}${minorEa}${minorCs}<a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>`
}

/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {PresentationPropsInternal} pres - presentation
 * @return {string} XML
 */
export function makeXmlPresentation(pres: PresentationPropsInternal): string {
	let strXml =
		`${XML_DECL}${CRLF}` +
		'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		// When fonts are embedded we carry WHOLE faces, so `embedTrueTypeFonts="1"` (so
		// PowerPoint honors the embed) and `saveSubsetFonts="0"` (we did not subset).
		// With no embedded fonts, keep the historical inert `saveSubsetFonts="1"`.
		`xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${pres.rtlMode ? 'rtl="1"' : ''} ${(pres.embeddedFonts || []).some((font) => font.faces.some((face) => face.bytes)) ? 'embedTrueTypeFonts="1" saveSubsetFonts="0"' : 'saveSubsetFonts="1"'} autoCompressPictures="0"${pres.firstSlideNum !== 1 ? ` firstSlideNum="${pres.firstSlideNum}"` : ''}>`

	// STEP 1: Add slide master (SPEC: tag 1 under <presentation>)
	strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'

	// STEP 2: Add Notes Master (SPEC: tag 2 under <presentation>)
	// CT_Presentation child sequence (ECMA-376 Part 1 §19.2.1.26) requires
	// notesMasterIdLst to appear BEFORE sldIdLst. Emitting it after sldIdLst
	// (or after sldSz/notesSz) violates the schema and is flagged by
	// OpenXmlValidator as Sch_UnexpectedElementContentExpectingComplex.
	// (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
	strXml += `<p:notesMasterIdLst><p:notesMasterId r:id="rId${pres.slides.length + 2}"/></p:notesMasterIdLst>`

	// STEP 3: Add all Slides (SPEC: tag 3 under <presentation>)
	strXml += '<p:sldIdLst>'
	pres.slides.forEach((slide) => (strXml += `<p:sldId id="${slide._slideId}" r:id="rId${slide._rId}"/>`))
	strXml += '</p:sldIdLst>'

	// STEP 4: Add sizes
	strXml += `<p:sldSz cx="${pres.presLayout.width}" cy="${pres.presLayout.height}"/>`
	strXml += `<p:notesSz cx="${pres.presLayout.height}" cy="${pres.presLayout.width}"/>`

	// STEP 4b: Embedded fonts (CT_Presentation index 7 — after notesSz, before defaultTextStyle).
	// rIds continue past the fixed presentation rels and must match makeXmlPresentationRels.
	{
		const fonts = pres.embeddedFonts || []
		const flat = flattenEmbeddedFaces(fonts, presentationFontRelStart(pres.slides))
		const rIdOf = new Map(flat.map((face) => [`${face.fontIndex}:${face.slot}`, face.rId]))
		strXml += serializeEmbeddedFontLst(fonts, (fontIndex, slot) => rIdOf.get(`${fontIndex}:${slot}`))
	}

	// STEP 5: Add text styles
	strXml += '<p:defaultTextStyle>'
	for (let idy = 1; idy < 10; idy++) {
		strXml +=
			`<a:lvl${idy}pPr marL="${(idy - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
			'<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' +
			`</a:defRPr></a:lvl${idy}pPr>`
	}
	strXml += '</p:defaultTextStyle>'

	// STEP 6: Add Sections (if any)
	if (pres.sections && pres.sections.length > 0) {
		strXml += '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
		strXml += '<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
		pres.sections.forEach((sect) => {
			strXml += `<p14:section name="${encodeXmlEntities(sect.title)}" id="{${getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')}}"><p14:sldIdLst>`
			sect._slides.forEach((slide) => (strXml += `<p14:sldId id="${slide._slideId}"/>`))
			strXml += '</p14:sldIdLst></p14:section>'
		})
		strXml += '</p14:sectionLst></p:ext>'
		strXml +=
			'<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext>'
		strXml += '</p:extLst>'
	}

	// Done
	strXml += '</p:presentation>'
	return strXml
}

/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export function makeXmlPresProps(): string {
	return `${XML_DECL}${CRLF}<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`
}

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles(tableStyles: TableStyleInternal[] = []): string {
	const open = `${XML_DECL}${CRLF}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"`
	if (!tableStyles || tableStyles.length === 0) return `${open}/>`

	let strXml = `${open}>`
	tableStyles.forEach(({ guid, def }) => {
		strXml += `<a:tblStyle styleId="${guid}" styleName="${encodeXmlEntities(def.name)}">`
		// NOTE: regions MUST be emitted in CT_TableStyle schema order or PowerPoint reports the file as corrupt
		;(
			[
				['wholeTbl', def.wholeTbl],
				['band1H', def.band1H],
				['band2H', def.band2H],
				['band1V', def.band1V],
				['band2V', def.band2V],
				['lastCol', def.lastCol],
				['firstCol', def.firstCol],
				['lastRow', def.lastRow],
				['firstRow', def.firstRow],
			] as const
		).forEach(([name, region]) => {
			if (region) strXml += genXmlTableStyleRegion(name, region)
		})
		strXml += '</a:tblStyle>'
	})
	strXml += '</a:tblStyleLst>'
	return strXml
}

/**
 * Build one `CT_TablePartStyle` region (e.g. `firstRow`, `band1H`) for a custom table style.
 * Emits `tcTxStyle` (text) before `tcStyle` (cell fill/borders) per the schema sequence.
 * @param {string} name - region element name
 * @param {TableStyleRegionProps} region - region styling
 * @return {string} XML
 */
function genXmlTableStyleRegion(name: string, region: TableStyleRegionProps): string {
	let xml = `<a:${name}>`

	// A: tcTxStyle — text style (only when text formatting is requested)
	if (region.bold !== undefined || region.italic !== undefined || region.color) {
		const b = region.bold ? ' b="on"' : ''
		const i = region.italic ? ' i="on"' : ''
		xml += `<a:tcTxStyle${b}${i}><a:fontRef idx="minor"/>`
		xml += region.color ? createColorElement(region.color) : ''
		xml += '</a:tcTxStyle>'
	}

	// B: tcStyle — cell style: tcBdr (borders) then fill, in schema order
	if (region.border !== undefined || region.fill !== undefined) {
		xml += '<a:tcStyle>'
		if (region.border !== undefined) xml += genXmlTableStyleBorders(region.border)
		if (region.fill !== undefined) xml += `<a:fill>${genXmlColorSelection(region.fill)}</a:fill>`
		xml += '</a:tcStyle>'
	}

	xml += `</a:${name}>`
	return xml
}

/**
 * Build the `tcBdr` border block for a custom table style region.
 * A single `BorderProps` styles all four sides plus the interior grid lines; a
 * TRBL array styles only the four outer sides. Sides are emitted in schema order.
 * @param {BorderProps | BorderProps[]} border - border definition
 * @return {string} XML
 */
function genXmlTableStyleBorders(border: BorderProps | BorderProps[]): string {
	// NOTE: order MUST be left,right,top,bottom,insideH,insideV (CT_TableCellBorderStyle sequence)
	let sides: Array<[string, BorderProps | undefined]>
	if (Array.isArray(border)) {
		const [top, right, bottom, left] = border // TRBL input order
		sides = [
			['left', left],
			['right', right],
			['top', top],
			['bottom', bottom],
		]
	} else {
		sides = [
			['left', border],
			['right', border],
			['top', border],
			['bottom', border],
			['insideH', border],
			['insideV', border],
		]
	}

	let xml = '<a:tcBdr>'
	sides.forEach(([side, b]) => {
		if (!b) return
		xml += `<a:${side}>`
		if (b.type === 'none') {
			xml += '<a:ln><a:noFill/></a:ln>'
		} else {
			xml += `<a:ln w="${lineWidthToEmu(resolveBorderWidth(b, 1))}" cap="flat" cmpd="sng" algn="ctr">`
			xml += genXmlColorSelection({ color: b.color ?? '666666', transparency: b.transparency })
			xml += `<a:prstDash val="${b.type === 'dash' ? 'sysDash' : 'solid'}"/>`
			xml += '</a:ln>'
		}
		xml += `</a:${side}>`
	})
	xml += '</a:tcBdr>'
	return xml
}

/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export function makeXmlViewProps(): string {
	return `${XML_DECL}${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
}
