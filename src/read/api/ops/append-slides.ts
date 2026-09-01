/**
 * ts-pptx: the body of {@link Presentation.appendSlides}
 *
 * The hybrid "generate onto an existing deck" path: a slide producer authors slides, they are
 * spliced into this package under fresh partnames, and every dependency each one carries —
 * images, embedded and online media, hyperlinks, speaker notes, charts, intra-batch slide links
 * — is wired into its `.rels`.
 *
 * Here rather than on the class for the reason `presentation.ts` states about the four import
 * bodies: the *contract* is the doc comment a caller reads and stays on the public method, while
 * the wiring is one job with a lot of parts. It was the largest thing left in that file. Each
 * loop that used to be an anonymous `for` inside one closure is now a named function carrying the
 * paragraph that explains the relationship layout it produces.
 *
 * The rel ids these functions allocate depend on the order they run in, so the two passes below
 * and the sequence inside pass 2 are both load-bearing.
 */

import { emuToInches } from '../../../units.js'
import { relativePartName } from '../../opc/partnames.js'
import type { Part } from '../../opc/part.js'
import type { Relationships } from '../../opc/relationships.js'
import { InvalidOptionError } from '../../../errors.js'
import {
	IMAGE_REL,
	NOTES_MASTER_REL,
	NOTES_SLIDE_CONTENT_TYPE,
	NOTES_SLIDE_REL,
	SLIDE_CONTENT_TYPE,
	SLIDE_LAYOUT_REL,
	SLIDE_REL,
} from '../../../ooxml/rel-types.js'
import type { Presentation } from '../presentation.js'
import type { Slide } from '../slide.js'
import type { AppendSlidesOptions, LayoutHandle, SlideSource } from '../presentation-types.js'
import { carryGeneratedEmbeddedFonts } from './embedded-fonts.js'
import { ensureNotesMasterFromXml } from './notes-master.js'
import { requireEqualSlideSize } from './slide-size.js'
import { pickDefined } from '../../../options-internal.js'

const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
// chartEx (Office 2016) charts are reached through a Microsoft rel, not the ECMA `chart` one,
// and their two mandatory style sidecars through a third pair — see `gen/chart/chartex-style.ts`.
const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx'
const CHART_STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
const CHART_COLOR_STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
// Microsoft 2007 `media` rel: paired with the ECMA audio/video rel (same Target),
// referenced by the slide body's <p14:media r:embed>.
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
// chartEx parts carry Microsoft content types, not the `openxmlformats` ones.
const CHARTEX_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml'
const CHART_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartstyle+xml'
const CHART_COLOR_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml'
const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

const textEncoder = new TextEncoder()

/** One extracted slide as pass 1 left it: the source body, the part it landed in, and its name. */
type PlacedSlide = {
	slide: Awaited<ReturnType<SlideSource['extractSlides']>>['slides'][number]
	part: Part
	partName: string
}

/**
 * Resolve `options.layout` — a layout *name* or a {@link LayoutHandle} — against this deck's
 * gallery, with no silent fallback: an unusable value is one of three named errors rather than a
 * default layout, because a slide bound to the wrong layout is a deck that looks authored.
 *
 * Split out because it is the part of the append most likely to be wanted by a second caller, and
 * because it is the only place `layout/ambiguous-name`, `layout/not-found` and
 * `layout/foreign-handle` are raised.
 *
 * @param deck - the destination deck, whose `layouts()` gallery is searched
 * @param layout - the caller's layout name or handle
 * @param api - the method name, opening each error message
 */
export function resolveLayoutTarget(
	deck: Presentation,
	layout: AppendSlidesOptions['layout'],
	api = 'appendSlides'
): LayoutHandle {
	const gallery = deck.layouts()
	if (typeof layout !== 'string') {
		if (!gallery.some((l) => l.partName === layout.partName)) {
			throw new InvalidOptionError(
				'layout/foreign-handle',
				`${api}: layout ${layout.partName} does not belong to this presentation`
			)
		}
		return layout
	}
	const matches = gallery.filter((l) => l.name === layout)
	if (matches.length > 1) {
		throw new InvalidOptionError(
			'layout/ambiguous-name',
			`${api}: layout name ${JSON.stringify(layout)} is ambiguous (${matches.length} layouts share it); pass a LayoutHandle from layouts() instead`
		)
	}
	const [only] = matches
	if (!only) {
		const names = gallery.map((l) => JSON.stringify(l.name)).join(', ')
		throw new InvalidOptionError(
			'layout/not-found',
			`${api}: no layout named ${JSON.stringify(layout)}; available: ${names || '(none)'}`
		)
	}
	return only
}

/** Image parts behind the slide body's `image` rels, each keeping the id the body was built with. */
function wireMedia(deck: Presentation, rels: Relationships, partName: string, slide: PlacedSlide['slide']): void {
	for (const m of slide.media) {
		const mediaPartName = deck.opc.reserveMediaPartName(m.extn)
		deck.opc.addPart(mediaPartName, m.contentType, m.bytes)
		rels.addWithId(`rId${m.rId}`, IMAGE_REL, relativePartName(partName, mediaPartName))
	}
}

/**
 * Embedded audio and video. One media part backs two rels (ECMA audio/video + MS-2007 media)
 * sharing its Target; the preview poster is a separate image part. `ensureDefault` runs before
 * `addPart` so the content type resolves via a Default extension entry (what PowerPoint authors)
 * rather than a per-part Override.
 */
function wireAvMedia(deck: Presentation, rels: Relationships, partName: string, slide: PlacedSlide['slide']): void {
	for (const av of slide.avMedia) {
		const mediaPartName = deck.opc.reserveMediaPartName(av.mediaExtn, 'media')
		deck.opc.contentTypes.ensureDefault(av.mediaExtn, av.mediaContentType)
		deck.opc.addPart(mediaPartName, av.mediaContentType, av.mediaBytes)
		const mediaTarget = relativePartName(partName, mediaPartName)
		rels.addWithId(`rId${av.mediaRid}`, av.mtype === 'audio' ? AUDIO_REL : VIDEO_REL, mediaTarget)
		rels.addWithId(`rId${av.msMediaRid}`, MS_MEDIA_REL, mediaTarget)

		const previewPartName = deck.opc.reserveMediaPartName(av.previewExtn)
		deck.opc.contentTypes.ensureDefault(av.previewExtn, av.previewContentType)
		deck.opc.addPart(previewPartName, av.previewContentType, av.previewBytes)
		rels.addWithId(`rId${av.previewRid}`, IMAGE_REL, relativePartName(partName, previewPartName))
	}
}

/**
 * Online (external-link) video: two External rels share the link Target — the ECMA video rel and
 * the MS-2007 media rel — with no media binary part and no content-type entry. The poster image is
 * wired by {@link wireMedia}, which runs first.
 */
function wireOnlineMedia(rels: Relationships, slide: PlacedSlide['slide']): void {
	for (const ov of slide.onlineMedia) {
		rels.addWithId(`rId${ov.mediaRid}`, VIDEO_REL, ov.link, 'External')
		rels.addWithId(`rId${ov.msMediaRid}`, MS_MEDIA_REL, ov.link, 'External')
	}
}

/** External hyperlink rels, each keeping the id the slide body references it by. */
function wireHyperlinks(rels: Relationships, slide: PlacedSlide['slide']): void {
	for (const h of slide.hyperlinks) {
		rels.addWithId(`rId${h.rId}`, HYPERLINK_REL, h.target, 'External')
	}
}

/**
 * Speaker notes. The notes part carries its own rel namespace, independent of the slide's:
 * rId1 = notesMaster, rId2 = the slide it annotates, hyperlinks from rId3 — the order the
 * generator's body was serialized against, so these are added by explicit id rather than left to
 * auto-numbering.
 */
function wireNotes(
	deck: Presentation,
	rels: Relationships,
	partName: string,
	slide: PlacedSlide['slide'],
	sourceNotesMaster: { xml: string; themeXml: string } | undefined | null
): void {
	if (!slide.notes) return
	const notesPartName = deck.opc.reservePartNameLike('/ppt/notesSlides/notesSlide1.xml')
	deck.opc.addPart(notesPartName, NOTES_SLIDE_CONTENT_TYPE, textEncoder.encode(slide.notes.xml))
	rels.add(NOTES_SLIDE_REL, relativePartName(partName, notesPartName))

	const notesRels = deck.opc.relationshipsFor(notesPartName)
	const notesMasterPartName = sourceNotesMaster ? ensureNotesMasterFromXml(deck, sourceNotesMaster) : null
	if (notesMasterPartName) {
		notesRels.addWithId('rId1', NOTES_MASTER_REL, relativePartName(notesPartName, notesMasterPartName))
	}
	notesRels.addWithId('rId2', SLIDE_REL, relativePartName(notesPartName, partName))
	for (const h of slide.notes.hyperlinks) {
		notesRels.addWithId(`rId${h.rId}`, HYPERLINK_REL, h.target, 'External')
	}
}

/**
 * Chart part + its embedded workbook, each under a fresh name. Both chart flavours reference the
 * workbook through the chart part's own rId1, so the chart `.rels` is rebuilt here against the
 * reserved workbook partname.
 *
 * A chartEx chart differs in every other coordinate: the `chartEx{N}.xml` name family, the MS
 * chartex content type, the MS `chartEx` rel off the slide, and two mandatory sidecars at
 * rId2/rId3 — without which PowerPoint reports the deck as corrupt (schema-valid but unopenable).
 * `c.chartEx` carries both sidecar bodies; the generator built them, so nothing here reaches into
 * the emitters.
 */
function wireCharts(deck: Presentation, rels: Relationships, partName: string, slide: PlacedSlide['slide']): void {
	for (const c of slide.charts) {
		const isChartEx = c.chartEx !== undefined
		const chartPartName = deck.opc.reservePartNameLike(
			isChartEx ? '/ppt/charts/chartEx1.xml' : '/ppt/charts/chart1.xml'
		)
		deck.opc.addPart(
			chartPartName,
			isChartEx ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE,
			textEncoder.encode(c.chartXml)
		)
		const embeddingPartName = deck.opc.reservePartNameLike('/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
		deck.opc.contentTypes.ensureDefault('xlsx', XLSX_CONTENT_TYPE)
		deck.opc.addPart(embeddingPartName, XLSX_CONTENT_TYPE, c.embeddingBytes)
		const chartRels = deck.opc.relationshipsFor(chartPartName)
		chartRels.addWithId('rId1', PACKAGE_REL, relativePartName(chartPartName, embeddingPartName))
		if (c.chartEx) {
			// rId2 = colors, rId3 = style: the order the write path authors
			// (`buildChartExRelsXml` in gen/chart/embed-xlsx.ts).
			const colorsPartName = deck.opc.reservePartNameLike('/ppt/charts/colors1.xml')
			deck.opc.addPart(colorsPartName, CHART_COLOR_STYLE_CONTENT_TYPE, textEncoder.encode(c.chartEx.colorsXml))
			chartRels.addWithId('rId2', CHART_COLOR_STYLE_REL, relativePartName(chartPartName, colorsPartName))

			const stylePartName = deck.opc.reservePartNameLike('/ppt/charts/style1.xml')
			deck.opc.addPart(stylePartName, CHART_STYLE_CONTENT_TYPE, textEncoder.encode(c.chartEx.styleXml))
			chartRels.addWithId('rId3', CHART_STYLE_REL, relativePartName(chartPartName, stylePartName))
		}
		rels.addWithId(`rId${c.rId}`, isChartEx ? CHARTEX_REL : CHART_REL, relativePartName(partName, chartPartName))
	}
}

/**
 * Intra-batch slide-to-slide links: a `slide:N` in the source is repointed at the Nth appended
 * slide's new partname. A link out of the batch has no counterpart to point at and throws — which
 * is why pass 1 places every slide before any of this runs.
 */
function wireSlideLinks(
	rels: Relationships,
	partName: string,
	slide: PlacedSlide['slide'],
	partBySourceNumber: Map<number, string>,
	index: number
): void {
	for (const link of slide.slideLinks) {
		const targetPartName = partBySourceNumber.get(link.sourceSlideNumber)
		if (!targetPartName) {
			throw new InvalidOptionError(
				'import/unresolved-slide-link',
				`appendSlides: slide ${index} links to source slide ${link.sourceSlideNumber}, which is not among the appended slides`
			)
		}
		rels.addWithId(`rId${link.rId}`, SLIDE_REL, relativePartName(partName, targetPartName))
	}
}

/**
 * Append a slide producer's slides onto `deck`. The caller-facing contract is the doc comment on
 * {@link Presentation.appendSlides}.
 */
export async function appendSlides(
	deck: Presentation,
	source: SlideSource,
	options: AppendSlidesOptions
): Promise<Slide[]> {
	// 1. Resolve the target layout partname (explicit; no silent fallback).
	const target = resolveLayoutTarget(deck, options.layout)

	// 2. Author + extract; enforce equal slide size (no geometry rescale in v1).
	const extracted = await source.extractSlides(pickDefined(options, ['onMediaError']))
	// The extracted size is always known, so it is lifted into a `SlideSize` rather than kept
	// as a loose pair — that is what let this call site keep a second formatter alive.
	requireEqualSlideSize(
		deck.slideSize,
		{
			widthEmu: extracted.widthEmu,
			heightEmu: extracted.heightEmu,
			widthIn: emuToInches(extracted.widthEmu),
			heightIn: emuToInches(extracted.heightEmu),
		},
		'appendSlides'
	)

	// Any existing slide partname seeds the fresh-partname family; fall back to a
	// literal seed for a slide-less template shell (reservePartNameLike parses the
	// string, it does not require the part to exist).
	const slideTemplate = deck.slides[0]?.partName ?? '/ppt/slides/slide1.xml'

	// Pass 1: reserve + add every slide body first, so internal slide-to-slide
	// links (which may point forward) can resolve to any appended slide. Adding
	// each part immediately claims its name — reservePartNameLike returns max+1
	// from the existing parts, so the next reservation sees it. (addPart registers
	// the slide's Override content type.)
	const placed: PlacedSlide[] = extracted.slides.map((slide) => {
		const partName = deck.opc.reservePartNameLike(slideTemplate)
		const part = deck.opc.addPart(partName, SLIDE_CONTENT_TYPE, textEncoder.encode(slide.xml))
		return { slide, part, partName }
	})

	// 1-based source slide number -> the appended slide's new partname.
	const partBySourceNumber = new Map<number, string>(placed.map((p, i) => [i + 1, p.partName]))

	// Pass 2: build each slide's .rels and wire it into presentation.xml. Media,
	// hyperlinks, charts, and slide-links keep the body's rId (addWithId); the
	// layout rel is added last via add() so its auto-id cannot collide.
	const added: Slide[] = []
	placed.forEach(({ slide, part, partName }, i) => {
		const rels = deck.opc.relationshipsFor(partName)
		wireMedia(deck, rels, partName, slide)
		wireAvMedia(deck, rels, partName, slide)
		wireOnlineMedia(rels, slide)
		wireHyperlinks(rels, slide)
		wireNotes(deck, rels, partName, slide, extracted.notesMaster)
		wireCharts(deck, rels, partName, slide)
		wireSlideLinks(rels, partName, slide, partBySourceNumber, i)
		rels.add(SLIDE_LAYOUT_REL, relativePartName(partName, target.partName))

		// Wire into presentation.xml (rel + p:sldId) at the requested position.
		const at = options.at === undefined ? undefined : options.at + i
		added.push(deck.insertSlidePart(part, at))
	})

	// Carry the generator's presentation-level embedded fonts (pptx.embedFont) into
	// this deck, so author-side embedded fonts survive the append onto a template.
	carryGeneratedEmbeddedFonts(deck, extracted.embeddedFonts || [])

	return added
}
