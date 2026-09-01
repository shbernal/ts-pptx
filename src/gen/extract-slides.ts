/**
 * ts-pptx: authored slides as injectable descriptors
 *
 * The serialization pass behind `Presentation.extractSlides()`: it runs the same
 * pre-serialization work a write does (`./prepare`), serializes each slide body, and resolves
 * everything the body *references* — image media, audio/video, online video, hyperlinks,
 * charts, notes — into descriptors the read side's `appendSlides` can splice into a loaded
 * deck without touching that deck's masters, layouts or theme.
 *
 * Nothing here resolves a descriptor into a package part. Reserving parts and rebuilding the
 * rel graph is the append path's job, on the destination deck's own numbering; this side only
 * says what each body refers to and by which `rId`.
 */

import { SlideObjectType } from '../enums.js'
import type { PresentationPropsInternal, PresSlideInternal } from '../types/internal.js'
import type { ExtractedSlide, ExtractedSlides } from '../read/api/presentation-types.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import type { FontMetricsRegistry } from '../measure/font-metrics.js'
import { InternalError } from '../errors.js'
import { isHyperlinkRel } from './utils.js'
import { msMediaRid, previewRid } from './define/media.js'
import { decodeBase64ToBytes } from '../media/base64.js'
import { avContentType, imageContentType } from '../media/content-type.js'
import { makeXmlCharts } from './chart/chart-xml.js'
import { makeXmlChartEx } from './chart/chartex-xml.js'
import { makeChartExColorsXml, makeChartExStyleXml } from './chart/chartex-style.js'
import { buildEmbeddedWorksheet } from './chart/embed-xlsx.js'
import { bakeSlideContent, encodeMediaForTargets } from './prepare.js'
import { makeXmlSlide } from './slide/slide.js'
import { buildNotesSlideRels, makeXmlNotesMaster, makeXmlNotesSlide } from './slide/notes.js'
import { makeXmlTheme } from './pres/theme.js'

/**
 * The slice of a deck's state an extraction reads — the same three fields `PackageSource`
 * carries for a write, minus the custom properties, which reach no slide. `presentation`
 * supplies the slides, the canvas size and the embedded fonts, and is what the notes master's
 * theme is built from.
 */
export interface ExtractSource {
	readonly runtime: RuntimeAdapter
	readonly presentation: PresentationPropsInternal
	readonly fontMetrics: FontMetricsRegistry
}

/** One slide's `_relsMedia`, indexed by `rId`, for the media descriptors that resolve rels by id. */
type RelsByRid = ReadonlyMap<number, PresSlideInternal['_relsMedia'][number]>

/**
 * Embedded audio/video. addMedia (`gen/define/media.ts`) pushes three consecutive
 * _relsMedia entries per item off `mediaRid`: the ECMA audio/video rel
 * (rId=mediaRid), the MS-2007 `media` rel (mediaRid+1, sharing one Target),
 * and the preview image rel (mediaRid+2). The body references all three; we
 * surface the descriptor so appendSlides can reproduce the rel graph. Online
 * media (external link) is excluded — it has a different rel shape and no part.
 */
function avMediaOf(slide: PresSlideInternal, relByRid: RelsByRid): ExtractedSlide['avMedia'] {
	return slide._slideObjects
		.filter((obj) => obj._type === SlideObjectType.media && obj.mtype !== 'online' && typeof obj.mediaRid === 'number')
		.map((obj) => {
			const mtype: 'audio' | 'video' = obj.mtype === 'audio' ? 'audio' : 'video'
			const mediaRid = obj.mediaRid as number
			const mediaRel = relByRid.get(mediaRid)
			const previewRel = relByRid.get(previewRid(mediaRid))
			if (!mediaRel || !previewRel) return null
			const mediaBytes = decodeBase64ToBytes(typeof mediaRel.data === 'string' ? mediaRel.data : '')
			const previewBytes = decodeBase64ToBytes(typeof previewRel.data === 'string' ? previewRel.data : '')
			if (!mediaBytes || !previewBytes) return null
			const mediaExtn = (mediaRel.extn || mediaRel.Target.split('.').pop() || 'mp4').toLowerCase()
			const previewExtn = (previewRel.extn || previewRel.Target.split('.').pop() || 'png').toLowerCase()
			return {
				mtype,
				mediaRid,
				msMediaRid: msMediaRid(mediaRid),
				previewRid: previewRid(mediaRid),
				mediaBytes,
				mediaExtn,
				mediaContentType: avContentType(mediaExtn, mtype),
				previewBytes,
				previewExtn,
				previewContentType: imageContentType(previewExtn),
			}
		})
		.filter((m): m is NonNullable<typeof m> => m !== null)
}

/**
 * Plain image media, as decoded bytes keyed by the `rId` the body uses.
 *
 * `avPreviewRids` are the poster frames of {@link avMediaOf}'s items: they live in
 * `_relsMedia` as image rels too, and their descriptor already carries them, so they are
 * excluded here rather than surfaced twice.
 */
function imageMediaOf(slide: PresSlideInternal, avPreviewRids: ReadonlySet<number>): ExtractedSlide['media'] {
	return slide._relsMedia
		.filter((rel) => rel.type.toLowerCase().includes('image') && !avPreviewRids.has(rel.rId))
		.map((rel) => {
			// Normalize the base64 payload's data-URI prefix: a bare payload gets a whole prefix,
			// one that carries a media type but no encoding gets the encoding.
			let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''
			if (!data.includes(',')) data = 'image/png;base64,' + data
			else if (!data.includes(';')) data = 'image/png;' + data
			const bytes = decodeBase64ToBytes(data)
			const extn = (rel.extn || rel.Target.split('.').pop() || 'png').toLowerCase()
			return bytes ? { rId: rel.rId, bytes, extn, contentType: imageContentType(extn) } : null
		})
		.filter((m): m is NonNullable<typeof m> => m !== null)
}

/**
 * Online (external-link) video. addMedia type:'online' pushes three rels off
 * mediaRid: the ECMA video rel (mediaRid, External), the MS-2007 media rel
 * (mediaRid+1, External, sharing the link Target), and the poster image rel
 * (mediaRid+2, carried by {@link imageMediaOf} as a normal image). No media binary
 * part exists; appendSlides reproduces only the two external rels + poster.
 */
function onlineMediaOf(slide: PresSlideInternal, relByRid: RelsByRid): ExtractedSlide['onlineMedia'] {
	return slide._slideObjects
		.filter((obj) => obj._type === SlideObjectType.media && obj.mtype === 'online' && typeof obj.mediaRid === 'number')
		.map((obj) => {
			const mediaRid = obj.mediaRid as number
			const videoRel = relByRid.get(mediaRid)
			if (!videoRel || typeof videoRel.Target !== 'string' || !videoRel.Target) return null
			return { mediaRid, msMediaRid: msMediaRid(mediaRid), link: videoRel.Target }
		})
		.filter((m): m is NonNullable<typeof m> => m !== null)
}

/** External hyperlink rels, keyed by the `rId` the body uses. */
function hyperlinksOf(slide: PresSlideInternal): ExtractedSlide['hyperlinks'] {
	return slide._rels
		.filter((rel) => isHyperlinkRel(rel) && rel.data !== 'slide')
		.map((rel) => ({ rId: rel.rId, target: rel.Target }))
}

/**
 * Internal slide-to-slide links: `rel.data === 'slide'`, and `rel.Target` is the 1-based
 * source slide number (see addText's hyperlink.slide handling).
 *
 * A Target that does not read as a slide number throws rather than passing `NaN` on to the
 * append path's rel-graph rebuild, where it surfaces as "links to source slide NaN" a whole
 * subsystem later. `addText` writes this Target itself, so the condition is a defect on this
 * side rather than bad input — an {@link InternalError}, like the chart router's own.
 */
function slideLinksOf(slide: PresSlideInternal): ExtractedSlide['slideLinks'] {
	return slide._rels
		.filter((rel) => isHyperlinkRel(rel) && rel.data === 'slide')
		.map((rel) => {
			const sourceSlideNumber = Number(rel.Target)
			if (!Number.isInteger(sourceSlideNumber) || sourceSlideNumber < 1)
				throw new InternalError(
					'slide/link-target-not-a-number',
					`Internal slide link rId${rel.rId} targets "${String(rel.Target)}", which is not a slide number`
				)
			return { rId: rel.rId, sourceSlideNumber }
		})
}

/**
 * Charts: the chart part XML plus its embedded workbook bytes. The chart part's own `.rels`
 * (workbook reference) is rebuilt on injection.
 *
 * chartEx charts are a different part (`makeXmlChartEx`) behind a different rel type and
 * content type, and PowerPoint reports one as corrupt without its style/colors sidecars
 * (see gen/chart/chartex-style.ts). Both ride in the descriptor's `chartEx` slot, which
 * is what tells `appendSlides` the two shapes apart: it cannot be inferred from the XML,
 * and building one as a classic chart is what used to produce a `<c:chartSpace>` with
 * axes and no plot behind a slide still pointing at it through `<cx:chart>`.
 */
function chartsOf(slide: PresSlideInternal): ExtractedSlide['charts'] {
	return (slide._relsChart || []).map((rel) => {
		const base = { rId: rel.rId, embeddingBytes: buildEmbeddedWorksheet(rel) }
		return rel.isChartEx
			? {
					...base,
					chartXml: makeXmlChartEx(rel),
					chartEx: { styleXml: makeChartExStyleXml(), colorsXml: makeChartExColorsXml() },
				}
			: { ...base, chartXml: makeXmlCharts(rel) }
	})
}

/**
 * Speaker notes, or `undefined` when the slide has none. makeXmlNotesSlide calls
 * buildNotesSlideRels itself (and caches on the slide), so the rels are read back afterwards
 * rather than rebuilt — the body and the rels file must agree on every hyperlink rId. Notes
 * rels reserve rId1=notesMaster and rId2=slide, both of which appendSlides wires itself.
 */
function notesOf(slide: PresSlideInternal): ExtractedSlide['notes'] {
	const hasNotes = slide._slideObjects.some((obj) => obj._type === SlideObjectType.notes)
	if (!hasNotes) return undefined
	return {
		xml: makeXmlNotesSlide(slide),
		hyperlinks: buildNotesSlideRels(slide)
			.filter((rel) => typeof rel.Target === 'string' && rel.Target)
			.map((rel) => ({ rId: rel.rId, target: rel.Target })),
	}
}

/**
 * Serialize a deck's authored slides as injectable descriptors, without producing a package.
 * @param source - the deck state to read; see {@link ExtractSource}
 * @param opts.onMediaError - what a media source that cannot be read does, as on `write()`
 */
export async function extractSlides(
	source: ExtractSource,
	opts: { onMediaError?: 'throw' | 'placeholder' } = {}
): Promise<ExtractedSlides> {
	const onMediaError = opts.onMediaError ?? 'throw'
	const { presentation } = source
	const deckSlides = presentation.slides

	// STEP 1+2: The same pre-serialization pass `buildPackageParts` runs — encode media,
	// backfill placeholders, bake measured fit — so extracted bodies match a normal write
	// by construction rather than by keeping two copies in step. See `gen/prepare.ts`.
	// Only slides here: this emits no layout or master parts.
	await encodeMediaForTargets(deckSlides, source.runtime, onMediaError)
	bakeSlideContent(deckSlides, source.fontMetrics)

	// STEP 3: Serialize each slide body and resolve what it references.
	const slides: ExtractedSlide[] = deckSlides.map((slide) => {
		const relByRid = new Map(slide._relsMedia.map((rel) => [rel.rId, rel] as const))
		const avMedia = avMediaOf(slide, relByRid)
		const notes = notesOf(slide)
		return {
			xml: makeXmlSlide(slide),
			media: imageMediaOf(slide, new Set(avMedia.map((item) => item.previewRid))),
			hyperlinks: hyperlinksOf(slide),
			charts: chartsOf(slide),
			slideLinks: slideLinksOf(slide),
			avMedia,
			onlineMedia: onlineMediaOf(slide, relByRid),
			...(notes ? { notes } : {}),
		}
	})

	// Presentation-level embedded fonts (pptx.embedFont) ride alongside the slides
	// so appendSlides can carry them into the destination deck; same model the
	// write path serializes (see src/embedded-fonts.ts), passed through unchanged.
	// A notes slide must bind to a notes master, and a destination template commonly has
	// none (a deck authored without notes carries no notesMaster part). Ship ours so the
	// append path can install one; it is discarded when the destination already has one.
	// theme2.xml is what notesMaster1.xml.rels resolves to on the normal write path.
	const notesMaster = slides.some((s) => s.notes)
		? { xml: makeXmlNotesMaster(), themeXml: makeXmlTheme(presentation) }
		: undefined

	return {
		widthEmu: presentation.presLayout.width,
		heightEmu: presentation.presLayout.height,
		slides,
		embeddedFonts: presentation.embeddedFonts,
		...(notesMaster ? { notesMaster } : {}),
	}
}
