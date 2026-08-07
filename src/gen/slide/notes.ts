/**
 * ts-pptx: speaker-notes parts
 *
 * Everything for the notes side of a deck: reading a slide's notes text/runs,
 * building the notes-slide hyperlink rels, and emitting the notes-slide,
 * notes-master and their `.rels` parts.
 */

import { SlideObjectType } from '../../enums.js'
import { CRLF, SLDNUMFLDID, XML_DECL } from '../../constants-internal.js'
import type { TextProps } from '../../types/index.js'
import type { PresSlideInternal, SlideRel } from '../../types/internal.js'
import { warn } from '../../diagnostics.js'
import { genXmlTextRun } from '../drawingml/text-run.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { OFFICE_REL, PACKAGE_REL_NS } from '../oxml/schema-uris.js'

/**
 * Get text content of Notes from Slide
 * @param {PresSlideInternal} slide - the slide object to transform into XML
 * @return {string} notes text
 */
export function getNotesFromSlide(slide: PresSlideInternal): string {
	let notesText = ''

	slide._slideObjects.forEach((data) => {
		if (data._type === SlideObjectType.notes) notesText += data?.text && data.text[0] ? data.text[0].text : ''
	})

	return notesText.replace(/\r*\n/g, CRLF)
}

/**
 * Collect the speaker-notes runs for a slide (flattened across any number of `addNotes()` calls).
 * @param {PresSlideInternal} slide - the slide object
 * @return {TextProps[]} notes text runs in document order
 */
function getNotesRuns(slide: PresSlideInternal): TextProps[] {
	const runs: TextProps[] = []
	slide._slideObjects.forEach((obj) => {
		if (obj._type === SlideObjectType.notes && obj.text) runs.push(...obj.text)
	})
	return runs
}

/**
 * Build (and cache) the hyperlink relationships for a slide's notes part (`notesSlideN.xml.rels`).
 *
 * Notes rels use their own namespace, independent of `slide._rels` (which serialize to
 * `slideN.xml.rels`). The notes part always reserves rId1=notesMaster and rId2=slide, so
 * dynamic hyperlink rels are allocated starting at rId3. Each notes hyperlink run is tagged
 * with its `_rId` so the body serializer and the rels file agree.
 *
 * Idempotent: the result is cached on `slide._relsNotes` and reused by both callers.
 * Only external `url` hyperlinks are supported; `slide` targets are ignored with a warning.
 * @param {PresSlideInternal} slide - the slide object
 * @return {SlideRel[]} notes hyperlink relationships
 */
export function buildNotesSlideRels(slide: PresSlideInternal): SlideRel[] {
	if (slide._relsNotes) return slide._relsNotes

	const NOTES_REL_RESERVED = 2 // rId1=notesMaster, rId2=slide
	const rels: SlideRel[] = []
	let lastRid = NOTES_REL_RESERVED

	getNotesRuns(slide).forEach((run) => {
		const hyperlink = run.options?.hyperlink
		if (!hyperlink) return
		if (!hyperlink.url) {
			// Notes support external `url` links only. Drop unsupported (e.g. `slide`) targets so the
			// run serializer doesn't emit a dangling <a:hlinkClick> with no matching relationship.
			if (hyperlink.slide)
				warn('notes/hyperlink-slide-unsupported', 'notes hyperlinks support `url` only (ignoring `slide` target)')
			if (run.options) delete run.options.hyperlink
			return
		}

		lastRid++
		hyperlink._rId = lastRid
		rels.push({
			type: SlideObjectType.hyperlink,
			data: 'dummy',
			rId: lastRid,
			// `Target` is stored RAW; every emitter escapes it. See the note on `SlideRel.Target`.
			Target: hyperlink.url,
		})
	})

	slide._relsNotes = rels
	return rels
}

/**
 * Build the `<p:txBody>` paragraphs for the notes placeholder.
 * Runs are split into `<a:p>` paragraphs on newlines; each run is serialized with the standard
 * text-run generator so inline formatting and `<a:hlinkClick>` markup are emitted consistently.
 * @param {PresSlideInternal} slide - the slide object
 * @return {string} XML string of `<a:p>` paragraphs
 */
function genXmlNotesParagraphs(slide: PresSlideInternal): string {
	const paragraphs: TextProps[][] = [[]]

	getNotesRuns(slide).forEach((run) => {
		const segments = String(run.text ?? '').split('\n')
		segments.forEach((segment, idx) => {
			if (idx > 0) paragraphs.push([]) // a newline starts a new paragraph
			const text = segment.replace(/\r/g, '')
			if (text !== '') paragraphs[paragraphs.length - 1]?.push({ text, options: run.options || {} })
		})
	})

	return paragraphs
		.map((runs) =>
			el('a:p', null, [
				raw(runs.map((run) => genXmlTextRun(run)).join('')),
				raw(voidEl('a:endParaRPr', { lang: 'en-US', dirty: 0 })),
			])
		)
		.join('')
}

/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
export function makeXmlNotesMaster(): string {
	const xfrmOffExt = (x: number, y: number, cx: number, cy: number): string =>
		el('a:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])
	const prstGeomRect = el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))
	const spLocks = (attrs: Record<string, number>): string => el('p:cNvSpPr', null, raw(voidEl('a:spLocks', attrs)))
	const nvPr = (phAttrs: Record<string, string | number>): string => el('p:nvPr', null, raw(voidEl('p:ph', phAttrs)))
	const bodyPr = (attrs: Record<string, string | number>): string => voidEl('a:bodyPr', attrs)
	const lstStyleLvl1 = (algn: string): string =>
		el('a:lstStyle', null, raw(el('a:lvl1pPr', { algn }, raw(voidEl('a:defRPr', { sz: 1200 })))))
	const emptyPara = el('a:p', null, raw(voidEl('a:endParaRPr', { lang: 'en-US' })))

	const header = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 2, name: 'Header Placeholder 1' })),
				raw(spLocks({ noGrp: 1 })),
				raw(nvPr({ type: 'hdr', sz: 'quarter' })),
			])
		),
		raw(el('p:spPr', null, [raw(xfrmOffExt(0, 0, 2971800, 458788)), raw(prstGeomRect)])),
		raw(
			el('p:txBody', null, [
				raw(bodyPr({ vert: 'horz', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0 })),
				raw(lstStyleLvl1('l')),
				raw(emptyPara),
			])
		),
	])

	const date = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 3, name: 'Date Placeholder 2' })),
				raw(spLocks({ noGrp: 1 })),
				raw(nvPr({ type: 'dt', idx: 1 })),
			])
		),
		raw(el('p:spPr', null, [raw(xfrmOffExt(3884613, 0, 2971800, 458788)), raw(prstGeomRect)])),
		raw(
			el('p:txBody', null, [
				raw(bodyPr({ vert: 'horz', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0 })),
				raw(lstStyleLvl1('r')),
				raw(
					el('a:p', null, [
						raw(
							el('a:fld', { id: '{5282F153-3F37-0F45-9E97-73ACFA13230C}', type: 'datetimeFigureOut' }, [
								raw(voidEl('a:rPr', { lang: 'en-US' })),
								raw(el('a:t', null, '7/23/19')),
							])
						),
						raw(voidEl('a:endParaRPr', { lang: 'en-US' })),
					])
				),
			])
		),
	])

	const slideImg = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 4, name: 'Slide Image Placeholder 3' })),
				raw(spLocks({ noGrp: 1, noRot: 1, noChangeAspect: 1 })),
				raw(nvPr({ type: 'sldImg', idx: 2 })),
			])
		),
		raw(
			el('p:spPr', null, [
				raw(xfrmOffExt(685800, 1143000, 5486400, 3086100)),
				raw(prstGeomRect),
				raw(voidEl('a:noFill')),
				raw(el('a:ln', { w: 12700 }, raw(el('a:solidFill', null, raw(voidEl('a:prstClr', { val: 'black' })))))),
			])
		),
		raw(
			el('p:txBody', null, [
				raw(bodyPr({ vert: 'horz', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0, anchor: 'ctr' })),
				raw(voidEl('a:lstStyle')),
				raw(emptyPara),
			])
		),
	])

	const notesPlaceholder = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 5, name: 'Notes Placeholder 4' })),
				raw(spLocks({ noGrp: 1 })),
				raw(nvPr({ type: 'body', sz: 'quarter', idx: 3 })),
			])
		),
		raw(el('p:spPr', null, [raw(xfrmOffExt(685800, 4400550, 5486400, 3600450)), raw(prstGeomRect)])),
		raw(
			el('p:txBody', null, [
				raw(bodyPr({ vert: 'horz', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0 })),
				raw(voidEl('a:lstStyle')),
				raw(
					['Click to edit Master text styles', 'Second level', 'Third level', 'Fourth level', 'Fifth level']
						.map((text, lvl) =>
							el('a:p', null, [
								raw(voidEl('a:pPr', { lvl })),
								raw(el('a:r', null, [raw(voidEl('a:rPr', { lang: 'en-US' })), raw(el('a:t', null, text))])),
							])
						)
						.join('')
				),
			])
		),
	])

	const footer = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 6, name: 'Footer Placeholder 5' })),
				raw(spLocks({ noGrp: 1 })),
				raw(nvPr({ type: 'ftr', sz: 'quarter', idx: 4 })),
			])
		),
		raw(el('p:spPr', null, [raw(xfrmOffExt(0, 8685213, 2971800, 458787)), raw(prstGeomRect)])),
		raw(
			el('p:txBody', null, [
				raw(bodyPr({ vert: 'horz', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0, anchor: 'b' })),
				raw(lstStyleLvl1('l')),
				raw(emptyPara),
			])
		),
	])

	const slideNum = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 7, name: 'Slide Number Placeholder 6' })),
				raw(spLocks({ noGrp: 1 })),
				raw(nvPr({ type: 'sldNum', sz: 'quarter', idx: 5 })),
			])
		),
		raw(el('p:spPr', null, [raw(xfrmOffExt(3884613, 8685213, 2971800, 458787)), raw(prstGeomRect)])),
		raw(
			el('p:txBody', null, [
				raw(bodyPr({ vert: 'horz', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0, anchor: 'b' })),
				raw(lstStyleLvl1('r')),
				raw(
					el('a:p', null, [
						raw(
							el('a:fld', { id: '{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}', type: 'slidenum' }, [
								raw(voidEl('a:rPr', { lang: 'en-US' })),
								raw(el('a:t', null, '‹#›')),
							])
						),
						raw(voidEl('a:endParaRPr', { lang: 'en-US' })),
					])
				),
			])
		),
	])

	const spTree = el('p:spTree', null, [
		raw(
			el('p:nvGrpSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 1, name: '' })),
				raw(voidEl('p:cNvGrpSpPr')),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(
			el(
				'p:grpSpPr',
				null,
				raw(
					el('a:xfrm', null, [
						raw(voidEl('a:off', { x: 0, y: 0 })),
						raw(voidEl('a:ext', { cx: 0, cy: 0 })),
						raw(voidEl('a:chOff', { x: 0, y: 0 })),
						raw(voidEl('a:chExt', { cx: 0, cy: 0 })),
					])
				)
			)
		),
		raw(header),
		raw(date),
		raw(slideImg),
		raw(notesPlaceholder),
		raw(footer),
		raw(slideNum),
	])

	const extLst = el(
		'p:extLst',
		null,
		raw(
			el(
				'p:ext',
				{ uri: '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}' },
				raw(
					voidEl('p14:creationId', {
						'xmlns:p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
						val: 1024086991,
					})
				)
			)
		)
	)

	const cSld = el('p:cSld', null, [
		raw(el('p:bg', null, raw(el('p:bgRef', { idx: 1001 }, raw(voidEl('a:schemeClr', { val: 'bg1' })))))),
		raw(spTree),
		raw(extLst),
	])

	const clrMap = voidEl('p:clrMap', {
		bg1: 'lt1',
		tx1: 'dk1',
		bg2: 'lt2',
		tx2: 'dk2',
		accent1: 'accent1',
		accent2: 'accent2',
		accent3: 'accent3',
		accent4: 'accent4',
		accent5: 'accent5',
		accent6: 'accent6',
		hlink: 'hlink',
		folHlink: 'folHlink',
	})

	const notesStyleLevel = (n: number, marL: number): string =>
		el(
			`a:lvl${n}pPr`,
			{ marL, algn: 'l', defTabSz: 914400, rtl: 0, eaLnBrk: 1, latinLnBrk: 0, hangingPunct: 1 },
			raw(
				el('a:defRPr', { sz: 1200, kern: 1200 }, [
					raw(el('a:solidFill', null, raw(voidEl('a:schemeClr', { val: 'tx1' })))),
					raw(voidEl('a:latin', { typeface: '+mn-lt' })),
					raw(voidEl('a:ea', { typeface: '+mn-ea' })),
					raw(voidEl('a:cs', { typeface: '+mn-cs' })),
				])
			)
		)
	const notesStyle = el(
		'p:notesStyle',
		null,
		raw(
			[0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600]
				.map((marL, i) => notesStyleLevel(i + 1, marL))
				.join('')
		)
	)

	return (
		XML_DECL +
		CRLF +
		el(
			'p:notesMaster',
			{
				'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
				'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
			},
			[raw(cSld), raw(clrMap), raw(notesStyle)]
		)
	)
}

/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {PresSlideInternal} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlNotesSlide(slide: PresSlideInternal): string {
	// Allocate notes hyperlink rels first so run serialization can reference the correct rId
	buildNotesSlideRels(slide)

	const slideImgPh = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 2, name: 'Slide Image Placeholder 1' })),
				raw(el('p:cNvSpPr', null, raw(voidEl('a:spLocks', { noGrp: 1, noRot: 1, noChangeAspect: 1 })))),
				raw(el('p:nvPr', null, raw(voidEl('p:ph', { type: 'sldImg' })))),
			])
		),
		raw(voidEl('p:spPr')),
	])

	const notesPh = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 3, name: 'Notes Placeholder 2' })),
				raw(el('p:cNvSpPr', null, raw(voidEl('a:spLocks', { noGrp: 1 })))),
				raw(el('p:nvPr', null, raw(voidEl('p:ph', { type: 'body', idx: 1 })))),
			])
		),
		raw(voidEl('p:spPr')),
		raw(el('p:txBody', null, [raw(voidEl('a:bodyPr')), raw(voidEl('a:lstStyle')), raw(genXmlNotesParagraphs(slide))])),
	])

	const slideNumPh = el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 4, name: 'Slide Number Placeholder 3' })),
				raw(el('p:cNvSpPr', null, raw(voidEl('a:spLocks', { noGrp: 1 })))),
				raw(el('p:nvPr', null, raw(voidEl('p:ph', { type: 'sldNum', sz: 'quarter', idx: 10 })))),
			])
		),
		raw(voidEl('p:spPr')),
		raw(
			el('p:txBody', null, [
				raw(voidEl('a:bodyPr')),
				raw(voidEl('a:lstStyle')),
				raw(
					el('a:p', null, [
						raw(
							el('a:fld', { id: SLDNUMFLDID, type: 'slidenum' }, [
								raw(voidEl('a:rPr', { lang: 'en-US' })),
								raw(el('a:t', null, slide._slideNum)),
							])
						),
						raw(voidEl('a:endParaRPr', { lang: 'en-US' })),
					])
				),
			])
		),
	])

	const spTree = el('p:spTree', null, [
		raw(
			el('p:nvGrpSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 1, name: '' })),
				raw(voidEl('p:cNvGrpSpPr')),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(
			el(
				'p:grpSpPr',
				null,
				raw(
					el('a:xfrm', null, [
						raw(voidEl('a:off', { x: 0, y: 0 })),
						raw(voidEl('a:ext', { cx: 0, cy: 0 })),
						raw(voidEl('a:chOff', { x: 0, y: 0 })),
						raw(voidEl('a:chExt', { cx: 0, cy: 0 })),
					])
				)
			)
		),
		raw(slideImgPh),
		raw(notesPh),
		raw(slideNumPh),
	])

	const extLst = el(
		'p:extLst',
		null,
		raw(
			el(
				'p:ext',
				{ uri: '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}' },
				raw(
					voidEl('p14:creationId', {
						'xmlns:p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
						val: 1024086991,
					})
				)
			)
		)
	)

	const cSld = el('p:cSld', null, [raw(spTree), raw(extLst)])

	return (
		XML_DECL +
		CRLF +
		el(
			'p:notes',
			{
				'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
				'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
			},
			[raw(cSld), raw(el('p:clrMapOvr', null, raw(voidEl('a:masterClrMapping'))))]
		)
	)
}

/**
 * Generates XML string for a notes-slide relation file (`ppt/notesSlides/_rels/notesSlideN.xml.rels`).
 * rId1=notesMaster and rId2=slide are always reserved; any notes hyperlink rels follow (rId3+).
 * @param {PresSlideInternal} slide - the slide whose notes part is being related
 * @param {number} slideNumber - 1-indexed slide number the notes part belongs to
 * @return {string} XML
 */
export function makeXmlNotesSlideRel(slide: PresSlideInternal, slideNumber: number): string {
	// Flat: the hyperlink rels run together on one line, after the indented rId1/rId2 pair.
	const hlinkRels = buildNotesSlideRels(slide)
		.map((rel) =>
			voidEl('Relationship', {
				Id: `rId${rel.rId}`,
				Type: OFFICE_REL + 'hyperlink',
				Target: rel.Target,
				TargetMode: 'External',
			})
		)
		.join('')

	return (
		XML_DECL +
		el(
			'Relationships',
			{ xmlns: PACKAGE_REL_NS },
			[
				raw(
					voidEl('Relationship', {
						Id: 'rId1',
						Type: OFFICE_REL + 'notesMaster',
						Target: '../notesMasters/notesMaster1.xml',
					})
				),
				raw(
					voidEl('Relationship', {
						Id: 'rId2',
						Type: OFFICE_REL + 'slide',
						Target: `../slides/slide${slideNumber}.xml`,
					})
				),
				// Always a child, even when empty, so its `childPrefix` indent is still emitted.
				raw(hlinkRels),
			],
			// Two quirks kept verbatim: this part alone follows XML_DECL with a bare `\n`
			// rather than CRLF, and the closing tag hugs the last child with no prefix.
			{ openPrefix: '\n\t\t', childPrefix: '\n\t\t\t', closePrefix: '' }
		)
	)
}

/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
export function makeXmlNotesMasterRel(): string {
	return (
		XML_DECL +
		CRLF +
		el(
			'Relationships',
			{ xmlns: PACKAGE_REL_NS },
			raw(voidEl('Relationship', { Id: 'rId1', Type: OFFICE_REL + 'theme', Target: '../theme/theme2.xml' })),
			// The closing tag is indented to child depth, not parent depth — as emitted today.
			{ childPrefix: '\n\t\t', closePrefix: '\n\t\t' }
		)
	)
}
