/**
 * Package assembly: turn an authored presentation's internal state into the full set of
 * OOXML package parts and hand the bytes to the ZIP writer. This is the write-side
 * "packaging" layer — `[Content_Types].xml`, the `_rels` graph, docProps, theme, the
 * per-slide/layout/master parts, comments, and chart/media rels — split out of the
 * authoring class (`TsPptx`) so that class stays a façade over slide authoring.
 *
 * The entry point `writePackage` takes a structural {@link PackageSource} that the
 * authoring class satisfies; it does not depend on the class itself, so the same pipeline
 * can be driven from any assembled deck state.
 */
import { ZipWriter } from '../zip.js'
import type { CustomPropertyValue, WriteProps } from '../types/index.js'
import type { PresentationPropsInternal, PresSlideInternal, SlideLayoutInternal } from '../types/internal.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import type { FontMetricsRegistry } from '../measure/font-metrics.js'
import { flattenEmbeddedFaces } from '../embedded-fonts.js'
import { getNewRelId } from '../gen/utils.js'
import { decodeBase64ToBytes } from '../media/base64.js'
import { audioExtensionForSubtype } from '../media/content-type.js'
import { createExcelWorksheet } from '../gen/chart/embed-xlsx.js'
import { bakeSlideContent, encodeMediaForTargets } from '../gen/prepare.js'
import { makeXmlApp } from '../gen/opc/app.js'
import { makeXmlContTypes } from '../gen/opc/content-types.js'
import { makeXmlCore } from '../gen/opc/core.js'
import { makeXmlCustomProperties } from '../gen/opc/custom-props.js'
import { makeXmlRootRels } from '../gen/opc/root-rels.js'
import { makeXmlPresentationRels } from '../gen/pres/presentation-rels.js'
import { makeXmlPresentation, makeXmlPresProps, makeXmlViewProps } from '../gen/pres/presentation.js'
import { makeXmlTableStyles } from '../gen/pres/table-styles.js'
import { makeXmlTheme } from '../gen/pres/theme.js'
import { makeXmlCommentAuthors, makeXmlComments, resolveCommentAuthors } from '../gen/slide/comments.js'
import { makeXmlLayout } from '../gen/slide/layout.js'
import { makeXmlMaster, makeXmlMasterRel } from '../gen/slide/master.js'
import {
	makeXmlNotesMaster,
	makeXmlNotesMasterRel,
	makeXmlNotesSlide,
	makeXmlNotesSlideRel,
} from '../gen/slide/notes.js'
import { makeXmlSlide, makeXmlSlideLayoutRel, makeXmlSlideRel } from '../gen/slide/slide.js'

/**
 * The slice of an authored presentation the packager reads. The authoring class satisfies
 * this structurally: `presentation` is its internal props view (slides, layouts, master,
 * embedded fonts, and metadata), and the remaining fields are state the part builders need
 * that the props view does not carry.
 */
export interface PackageSource {
	readonly runtime: RuntimeAdapter
	readonly presentation: PresentationPropsInternal
	readonly customProperties: Array<{ name: string; value: CustomPropertyValue }>
	readonly fontMetrics: FontMetricsRegistry
}

/**
 * Media extensions whose bytes are already entropy-coded, so running the ZIP's
 * DEFLATE pass over them costs CPU for a negligible size gain. For these we set
 * the per-entry ZIP compression to STORE while leaving XML parts on DEFLATE.
 * In image/video-heavy decks media dominates the byte count, so this is the
 * dominant cost when writing large presentations.
 * Formats that genuinely benefit from DEFLATE (bmp, wav, tiff, emf, wmf, svg)
 * are deliberately excluded so they keep inheriting the global compression.
 */
const ALREADY_COMPRESSED_MEDIA_EXTN = new Set([
	'jpg',
	'jpeg',
	'png',
	'gif',
	'webp',
	'heic',
	'heif',
	'avif',
	'mp4',
	'm4v',
	'mov',
	'avi',
	'mpg',
	'mpeg',
	'wmv',
	'webm',
	'mkv',
	'mp3',
	'm4a',
	'aac',
	'ogg',
	'oga',
])

/**
 * Extensions whose payload is itself a ZIP archive — the embedded OPC packages an OLE object can
 * carry (`addOleObject`). Deflating a zip inside a zip buys nothing, so these are STOREd like the
 * already-compressed media above. No media/image rel ever uses one of these extensions, so decks
 * without an OLE object are unaffected.
 */
const ZIP_CONTAINER_EXTN = new Set(['xlsx', 'xlsm', 'docx', 'docm', 'pptx', 'pptm'])

/**
 * Register an audio media part + relationship for each slide-transition start sound
 * (`transition.sound` with `data`/`path`), stamping the assigned relationship id onto
 * `transition._sndRId` for the `p:sndAc/p:snd r:embed`. Runs before media encoding so
 * the bytes are loaded; idempotent (skips a sound already registered) so re-export is
 * safe. The stop-previous form (`sound.stopPrevious`) needs no part and is skipped.
 */
function registerTransitionSounds(slides: PresSlideInternal[]): void {
	slides.forEach((slide) => {
		const transition = slide.transition
		const sound = transition?.sound
		if (!sound || sound.stopPrevious || typeof transition._sndRId === 'number') return
		if (!sound.data && !sound.path) return

		// Derive the file extension from the data-URI mime, else the path, defaulting to wav.
		// The mime's subtype is not itself an extension for the spellings PowerPoint actually
		// uses (`audio/x-wav` for a transition sound), so it goes through the mapping rather
		// than into the filename raw — see `audioExtensionForSubtype`.
		const dataMime = /audio\/([\w.-]+)[;,]/.exec(sound.data ?? '')
		const pathFile = sound.path ? ((sound.path.split('/').pop() ?? '').split('?')[0] ?? '') : ''
		const extn = dataMime
			? audioExtensionForSubtype(dataMime[1] ?? '')
			: (pathFile.split('.').pop() ?? 'wav').toLowerCase()

		const rId = getNewRelId(slide)
		const mediaSlideKey =
			slide._slideNum == null ? 'sm' : slide._slideNum >= 1000 ? `sl-${slide._slideNum}` : slide._slideNum
		slide._relsMedia.push({
			path: sound.path ?? `preencoded.${extn}`,
			type: `audio/${extn}`,
			extn,
			data: sound.data ?? '',
			rId,
			Target: `../media/audio-${mediaSlideKey}-${slide._relsMedia.length + 1}.${extn}`,
		})
		transition._sndRId = rId
	})
}

/**
 * Create all chart and media rels for this Presentation
 * @param {PresSlideInternal | SlideLayoutInternal} slide - slide with rels
 * @param {ZipWriter} zip - zip writer
 * @param {Promise<string>[]} chartPromises - promise array
 */
function createChartMediaRels(
	slide: PresSlideInternal | SlideLayoutInternal,
	zip: ZipWriter,
	chartPromises: Promise<string>[]
): void {
	slide._relsChart.forEach((rel) => chartPromises.push(createExcelWorksheet(rel, zip)))
	slide._relsMedia.forEach((rel) => {
		if (rel.type !== 'online' && rel.type !== 'hyperlink') {
			// A: Loop vars
			let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

			// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
			if (!data.includes(',') && !data.includes(';')) data = 'image/png;base64,' + data
			else if (!data.includes(',')) data = 'image/png;base64,' + data
			else if (!data.includes(';')) data = 'image/png;' + data

			// C: Add media. fflate needs decoded bytes (no base64 convenience), so
			// decode the payload here. Already-compressed formats (JPEG/PNG/video/…)
			// gain ~nothing from DEFLATE, so STORE them to avoid wasted compression
			// CPU on large decks; other parts inherit global compression.
			const bytes = decodeBase64ToBytes(data)
			if (!bytes) return
			const extn = (rel.extn || rel.Target.split('.').pop() || '').toLowerCase()
			zip.add(rel.Target.replace('..', 'ppt'), bytes, {
				store: ALREADY_COMPRESSED_MEDIA_EXTN.has(extn) || ZIP_CONTAINER_EXTN.has(extn),
			})
		}
	})
}

/**
 * One emitted OOXML package part, before zipping: its slash-path, already-encoded bytes, and
 * whether it is added with `store` (DEFLATE skipped — already-compressed media/fonts). This is
 * the build-side seam {@link buildPackageParts} returns and {@link zipPackageParts} consumes;
 * the `store` hint is an fflate-era zip optimization kept internal so re-zipping stays
 * byte-identical. A public parts API exposes only `{ path, data }`.
 */
export interface InternalPackagePart {
	readonly path: string
	readonly data: Uint8Array
	readonly store: boolean
}

/**
 * Assemble every package part for `source` and return them in emission order, without zipping.
 * This runs the transition-sound registration, media encode, cross-deck media de-dup,
 * chart-part-id assignment, placeholder backfill, and measured-fit passes before the synchronous
 * XML pass reads slide state. The bytes are the same the ZIP writer would compress; splitting the
 * assembly from the zip lets the byte-identity harness (and a future parts API) read parts
 * directly. Only `onMediaError` is meaningful here — compression/output shape are zip concerns
 * handled by {@link zipPackageParts}.
 */
export async function buildPackageParts(
	source: PackageSource,
	props: { onMediaError?: WriteProps['onMediaError'] }
): Promise<InternalPackagePart[]> {
	const pres = source.presentation
	const arrChartPromises: Promise<string>[] = []
	const zip = new ZipWriter()

	// STEP 0: Register transition-sound media parts/rels before encoding picks them up.
	registerTransitionSounds(pres.slides)

	// STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
	const onMediaError = props.onMediaError ?? 'throw'
	const mediaTargets = [...pres.slides, ...pres.slideLayouts, pres.masterSlide]

	// STEP 2: Wait for media (if any) then generate the PPTX file
	return await encodeMediaForTargets(mediaTargets, source.runtime, onMediaError).then(async () => {
		// PERF: Collapse identical media to a single package part across the entire deck.
		// Each target (slide/layout/master) namespaces its media `Target` by slide, so the
		// same image used on multiple slides — or loaded from the same path — otherwise
		// embeds one copy per use. By now `encodeSlideMediaRels` has populated every
		// `rel.data`, so we can point later duplicates at the first occurrence's `Target`
		// (slide `.rels` reference media by rId, and sharing a part across slides is valid
		// OOXML). This subsumes the per-slide path/data de-dup for cross-slide reuse and
		// also covers background images.
		const canonicalMediaTargets = new Map<string, string>()
		for (const target of [...pres.slides, ...pres.slideLayouts, pres.masterSlide]) {
			for (const rel of target._relsMedia || []) {
				if (rel.type === 'online' || rel.type === 'hyperlink' || typeof rel.data !== 'string' || !rel.data) continue
				// OLE payloads are exempt: PowerPoint gives every embedded object its own part, and
				// collapsing two identical ones would make editing either rewrite the other's source.
				if (rel.oleRelType) continue
				// Key on extension + bytes so identical content with differing part
				// extensions is never merged into one mistyped file.
				const key = (rel.extn || '') + '\0' + rel.data
				const canonical = canonicalMediaTargets.get(key)
				if (canonical) rel.Target = canonical
				else canonicalMediaTargets.set(key, rel.Target)
			}
		}

		// DETERMINISM: Assign chart part filenames from a per-presentation counter here,
		// at write time, so two identical decks built in one process produce byte-identical
		// packages. Chart parts share one `ppt/charts/` namespace across slides, layouts, and
		// the master, so the id must be package-wide; `addChartDefinition` only sets a
		// target-local placeholder. This is the authoritative assignment consumed by content
		// types, slide rels, and the chart/embedding parts below — all emitted after this pass.
		// A never-reset module global previously drove this (same input, different bytes).
		// See backlog fork-chart-counter-nondeterminism.
		let chartPartIdx = 0
		for (const target of [...pres.slides, ...pres.slideLayouts, pres.masterSlide]) {
			for (const rel of target._relsChart || []) {
				const chartId = ++chartPartIdx
				rel.globalId = chartId
				// chartEx charts share the `ppt/charts/` namespace but use the `chartEx{N}.xml` name.
				// The single shared counter keeps every chart part name globally unique regardless of
				// prefix, so classic and chartEx parts never collide.
				const chartBase = rel.isChartEx ? `chartEx${chartId}` : `chart${chartId}`
				rel.fileName = `${chartBase}.xml`
				rel.Target = `/ppt/charts/${chartBase}.xml`
			}
		}

		// A: Backfill inherited layout placeholders, then bake a real fontScale onto
		// `fit:'shrink'` text boxes when font metrics are registered — both before the
		// sync XML pass reads them. Shared with `extractSlides` (see `gen/prepare.ts`).
		bakeSlideContent(pres.slides, source.fontMetrics)

		// B: Add all required files. fflate keys on full slash-paths and emits no
		// directory entries, so there is no folder scaffolding to set up (and no
		// stray empty-directory entries to guard against on minimal decks).
		const hasCustomProps = source.customProperties.length > 0
		zip.add(
			'[Content_Types].xml',
			makeXmlContTypes(pres.slides, pres.slideLayouts, pres.masterSlide, hasCustomProps, pres.embeddedFonts)
		)
		zip.add('_rels/.rels', makeXmlRootRels(hasCustomProps))
		zip.add('docProps/app.xml', makeXmlApp(pres.slides, pres.company))
		zip.add('docProps/core.xml', makeXmlCore(pres.title, pres.subject, pres.author, pres.revision))
		if (hasCustomProps) {
			zip.add('docProps/custom.xml', makeXmlCustomProperties(source.customProperties))
		}
		zip.add('ppt/_rels/presentation.xml.rels', makeXmlPresentationRels(pres.slides, pres.embeddedFonts))
		// Embedded font parts (raw whole faces). Fonts are already compact binary, so STORE
		// (no DEFLATE) like already-compressed media. Part index matches the rels Target above.
		for (const face of flattenEmbeddedFaces(pres.embeddedFonts, 1)) {
			zip.add(`ppt/fonts/font${face.partIndex}.fntdata`, face.bytes, { store: true })
		}
		zip.add('ppt/theme/theme1.xml', makeXmlTheme(pres))
		// emit a separate theme2.xml part so notesMaster1.xml.rels resolves
		zip.add('ppt/theme/theme2.xml', makeXmlTheme(pres))
		zip.add('ppt/presentation.xml', makeXmlPresentation(pres))
		zip.add('ppt/presProps.xml', makeXmlPresProps())
		zip.add('ppt/tableStyles.xml', makeXmlTableStyles())
		zip.add('ppt/viewProps.xml', makeXmlViewProps())

		// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
		pres.slideLayouts.forEach((layout, idx) => {
			zip.add(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, makeXmlLayout(layout))
			zip.add(
				`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`,
				makeXmlSlideLayoutRel(idx + 1, pres.slideLayouts)
			)
		})
		pres.slides.forEach((slide, idx) => {
			zip.add(`ppt/slides/slide${idx + 1}.xml`, makeXmlSlide(slide))
			zip.add(`ppt/slides/_rels/slide${idx + 1}.xml.rels`, makeXmlSlideRel(pres.slides, pres.slideLayouts, idx + 1))
			// Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
			zip.add(`ppt/notesSlides/notesSlide${idx + 1}.xml`, makeXmlNotesSlide(slide))
			zip.add(`ppt/notesSlides/_rels/notesSlide${idx + 1}.xml.rels`, makeXmlNotesSlideRel(slide, idx + 1))
		})
		zip.add('ppt/slideMasters/slideMaster1.xml', makeXmlMaster(pres.masterSlide, pres.slideLayouts))
		zip.add('ppt/slideMasters/_rels/slideMaster1.xml.rels', makeXmlMasterRel(pres.masterSlide, pres.slideLayouts))
		zip.add('ppt/notesMasters/notesMaster1.xml', makeXmlNotesMaster())
		zip.add('ppt/notesMasters/_rels/notesMaster1.xml.rels', makeXmlNotesMasterRel())

		// C.1: Comments — resolve the deck-wide author registry once, then emit the shared
		// commentAuthors part plus a per-slide comment part for each slide that has comments.
		const resolvedComments = resolveCommentAuthors(pres.slides)
		if (resolvedComments.authors.length > 0) {
			zip.add('ppt/commentAuthors.xml', makeXmlCommentAuthors(resolvedComments.authors))
			pres.slides.forEach((slide, idx) => {
				if ((slide._comments || []).length > 0) {
					zip.add(`ppt/comments/comment${idx + 1}.xml`, makeXmlComments(slide, resolvedComments.meta))
				}
			})
		}

		// D: Create all Rels (images, media, chart data)
		pres.slideLayouts.forEach((layout) => {
			createChartMediaRels(layout, zip, arrChartPromises)
		})
		pres.slides.forEach((slide) => {
			createChartMediaRels(slide, zip, arrChartPromises)
		})
		createChartMediaRels(pres.masterSlide, zip, arrChartPromises)

		// E: Wait for the chart-embed Promises (if any), then snapshot the accumulated
		// parts in emission order. Zipping is deferred to `zipPackageParts`.
		return await Promise.all(arrChartPromises).then(() => zip.entries())
	})
}

/**
 * Zip an ordered list of package parts into the output shape `props.outputType` selects. Re-adds
 * each part to a fresh {@link ZipWriter} in order — preserving each part's `store` hint — so the
 * archive is byte-identical to assembling straight into one writer. Compression and output type
 * are the only zip-level knobs; part assembly happened in {@link buildPackageParts}.
 */
async function zipPackageParts(
	parts: InternalPackagePart[],
	props: WriteProps
): Promise<string | ArrayBuffer | Blob | Uint8Array> {
	const zip = new ZipWriter()
	for (const part of parts) zip.add(part.path, part.data, { store: part.store })

	const compression = props.compression !== false
	if (props.outputType === 'STREAM') {
		// A: stream file
		return await zip.generate('nodebuffer', { compression })
	} else if (props.outputType) {
		// B: Node [fs]: Output type user option or default
		return await zip.generate(props.outputType, { compression })
	} else {
		// C: Browser: Output blob as app/ms-pptx
		return await zip.generate('blob', { compression })
	}
}

/**
 * Assemble every package part for `source` and zip it into the output shape `props.outputType`
 * selects. Thin composition of {@link buildPackageParts} (the assembly pipeline) and
 * {@link zipPackageParts} (the zip pass), kept as the stable entry point the authoring class calls.
 */
export async function writePackage(
	source: PackageSource,
	props: WriteProps
): Promise<string | ArrayBuffer | Blob | Uint8Array> {
	return await zipPackageParts(await buildPackageParts(source, props), props)
}
