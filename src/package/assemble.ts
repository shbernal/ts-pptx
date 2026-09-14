/**
 * Package assembly: turn an authored presentation's internal state into the full set of
 * OOXML package parts and hand the bytes to the ZIP writer. This is the write-side
 * "packaging" layer — `[Content_Types].xml`, the `_rels` graph, docProps, theme, the
 * per-slide/layout/master parts, and media rels — split out of the authoring class
 * (`TsPptx`) so that class stays a façade over slide authoring. The parts a construct
 * family adds on top of that skeleton — charts, comments, notes — come from the part
 * contributors it is handed, so nothing here names a family.
 *
 * The entry point `writePackage` takes a structural {@link PackageSource} that the
 * authoring class satisfies; it does not depend on the class itself, so the same pipeline
 * can be driven from any assembled deck state.
 */
import { ZipWriter } from '../zip.js'
import type { CustomPropertyValue, WriteProps } from '../types/index.js'
import type { PresentationPropsInternal, PresSlideInternal, SlideRelMedia } from '../types/internal.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import type { FontMetricsRegistry } from '../measure/font-metrics.js'
import { flattenEmbeddedFaces } from '../embedded-fonts.js'
import { getNewRelId } from '../gen/utils.js'
import { pushMediaRel } from '../gen/define/image-rel.js'
import { decodeBase64ToBytes } from '../media/base64.js'
import { audioExtensionForSubtype } from '../media/content-type.js'
import { TRANSITION_TYPES } from '../ooxml/st-enums.js'
import {
	backfillPlaceholders,
	bakeMeasuredFit,
	encodeMediaForTargets,
	requireSlideLinksInDeck,
} from '../gen/prepare.js'
import { makeXmlApp } from '../gen/opc/app.js'
import { makeXmlContTypes } from '../gen/opc/content-types.js'
import { makeXmlCore } from '../gen/opc/core.js'
import { makeXmlCustomProperties } from '../gen/opc/custom-props.js'
import { makeXmlRootRels } from '../gen/opc/root-rels.js'
import { makeXmlPresentationRels } from '../gen/pres/presentation-rels.js'
import { makeXmlPresentation, makeXmlPresProps, makeXmlViewProps } from '../gen/pres/presentation.js'
import { makeXmlTableStyles } from '../gen/pres/table-styles.js'
import { makeXmlTheme } from '../gen/pres/theme.js'
import { makeXmlLayout } from '../gen/slide/layout.js'
import type { RendererTable } from '../gen/slide/objects/shared.js'
import { makeXmlMaster, makeXmlMasterRel } from '../gen/slide/master.js'
import { makeXmlSlide, makeXmlSlideLayoutRel, makeXmlSlideRel } from '../gen/slide/slide.js'
import { collectContentTypes, orderedContributors, type PartContributor, type PartTarget } from './parts/shared.js'
import {
	fontPath,
	PRESENTATION_PATH,
	relsPath,
	slideLayoutPath,
	SLIDE_MASTER_PATH,
	slidePath,
} from '../gen/opc/part-paths.js'

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
	/**
	 * Which renderer emits each shape family, handed down to every part that carries a shape tree
	 * (slides, layouts, the master). It travels with the deck state rather than being imported by
	 * the shape walk so that the set of families a program links is the caller's to decide — see
	 * `RendererTable` in `gen/slide/objects/shared.ts`.
	 */
	readonly renderers: RendererTable
	/**
	 * Which construct families add parts to the package — the chart parts and their embedded
	 * workbooks, the comment parts, the notes slides. Travels with the deck state for the same
	 * reason `renderers` does: naming the emitters here would link every family's part builders
	 * into every program that writes a deck. See `PartContributor` in `parts/shared.ts`.
	 */
	readonly partContributors: readonly PartContributor[]
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

/** The audio relationship a slide's transition sound was registered under, and the sound it was registered for. */
interface RegisteredTransitionSound {
	readonly rel: SlideRelMedia
	readonly data: string | undefined
	readonly path: string | undefined
}

/**
 * Each slide's registered transition sound.
 *
 * Kept per slide, never on the transition: one transition object can be assigned to several slides,
 * and a relationship id means something only inside the slide part that declares it. Kept across
 * writes as well, so a deck written twice registers each sound once. An entry is reused only while
 * the slide still holds its relationship and still has the same sound assigned.
 */
const registeredTransitionSounds = new WeakMap<PresSlideInternal, RegisteredTransitionSound>()

/**
 * Register an audio media part + relationship for each slide-transition start sound
 * (`transition.sound` with `data`/`path`), and return the relationship id each slide's
 * `p:sndAc/p:snd r:embed` refers to. Runs before media encoding so the bytes are loaded. Re-export
 * is safe: a slide whose sound is already registered reuses that relationship. The stop-previous
 * form (`sound.stopPrevious`) needs no part and is skipped.
 *
 * Registered here, at write time, rather than when `slide.transition` is assigned. A slide's rel
 * ids and media part names are handed out in registration order, and a transition is often set
 * before the slide's pictures and media are added; registering the sound then would move every
 * later rel id and part name on the slide.
 */
function registerTransitionSounds(slides: PresSlideInternal[]): Map<PresSlideInternal, number> {
	const soundRIds = new Map<PresSlideInternal, number>()
	slides.forEach((slide) => {
		const transition = slide.transition
		// A transition the emitter will not write has no sound action to register a part for.
		if (!transition || !(TRANSITION_TYPES as readonly string[]).includes(transition.type)) return
		const sound = transition.sound
		if (!sound || sound.stopPrevious) return
		if (!sound.data && !sound.path) return

		const registered = registeredTransitionSounds.get(slide)
		if (
			registered &&
			slide._relsMedia.includes(registered.rel) &&
			registered.data === sound.data &&
			registered.path === sound.path
		) {
			soundRIds.set(slide, registered.rel.rId)
			return
		}

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
		const rel = pushMediaRel(slide, {
			kind: 'audio',
			extn,
			type: `audio/${extn}`,
			path: sound.path,
			data: sound.data,
			rId,
		})
		registeredTransitionSounds.set(slide, { rel, data: sound.data, path: sound.path })
		soundRIds.set(slide, rId)
	})
	return soundRIds
}

/**
 * Write the media parts one target's rels point at. Images are how every deck carries a picture,
 * so this stays on the core path; the chart half of what used to be one function is a part
 * contributor (`parts/chart.ts`) and runs just before it, per target, so part order is unchanged.
 * @param {PartTarget} slide - slide, layout or master carrying the rels
 * @param {ZipWriter} zip - zip writer
 */
function createMediaParts(slide: PartTarget, zip: ZipWriter): void {
	slide._relsMedia.forEach((rel) => {
		if (rel.type !== 'online' && rel.type !== 'hyperlink') {
			// A: fflate needs decoded bytes (no base64 convenience), so decode the payload
			// here. `decodeBase64ToBytes` takes raw base64 and a `data:` URI alike, which is
			// every shape `rel.data` comes in: a loaded rel has been through `toMediaDataUri`,
			// and a caller-supplied `addImage({ data })` may be either. A prefix-correcting
			// block used to sit here and is gone -- it only ever prepended a label the decoder
			// then stripped, so it could not change a single byte of any payload.
			const data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

			// B: Already-compressed formats (JPEG/PNG/video/…) gain ~nothing from DEFLATE, so
			// STORE them to avoid wasted compression CPU on large decks; other parts inherit
			// global compression.
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
interface InternalPackagePart {
	readonly path: string
	readonly data: Uint8Array
	readonly store: boolean
}

/**
 * Assemble every package part for `source` and return them in emission order, without zipping.
 * This runs the transition-sound registration, placeholder backfill, media encode, cross-deck
 * media de-dup, chart-part-id assignment, and measured-fit passes before the synchronous
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
	// One pool for every contributor's async part work, awaited once at the end: a family that
	// awaited its own would make the write serial.
	const partPromises: Promise<unknown>[] = []
	const contributors = orderedContributors(source.partContributors)
	const zip = new ZipWriter()

	// A link to a slide the deck does not have is refused before anything is registered or built.
	requireSlideLinksInDeck(pres.slides)

	// STEP 0: Register transition-sound media parts/rels before encoding picks them up.
	const transitionSoundRIds = registerTransitionSounds(pres.slides)
	// STEP 0b: Seed each slide with the layout placeholders it leaves empty, for the same reason:
	// a seeded placeholder with an image fill registers media, and encoding has to see it. Shared
	// with `extractSlides` (see `gen/prepare.ts`).
	backfillPlaceholders(pres.slides)

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
		for (const target of mediaTargets) {
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
		// A never-reset module global in `gen/define/chart.ts` previously drove this — same
		// input, different bytes. `gen/chart/chartex-xml.ts` derives its series GUIDs from
		// the id assigned here, so both depend on this pass staying authoritative.
		let chartPartIdx = 0
		for (const target of mediaTargets) {
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

		// A: Bake a real fontScale onto `fit:'shrink'` text boxes when font metrics are registered,
		// before the sync XML pass reads them. Shared with `extractSlides` (see `gen/prepare.ts`).
		bakeMeasuredFit(pres.slides, source.fontMetrics)

		// B: Add all required files. fflate keys on full slash-paths and emits no
		// directory entries, so there is no folder scaffolding to set up (and no
		// stray empty-directory entries to guard against on minimal decks).
		const hasCustomProps = source.customProperties.length > 0
		zip.add(
			'[Content_Types].xml',
			makeXmlContTypes({
				slides: pres.slides,
				slideLayouts: pres.slideLayouts,
				masterSlide: pres.masterSlide,
				hasCustomProps,
				embeddedFonts: pres.embeddedFonts,
				contributions: collectContentTypes(contributors, pres),
			})
		)
		zip.add('_rels/.rels', makeXmlRootRels(hasCustomProps))
		zip.add('docProps/app.xml', makeXmlApp(pres.slides, pres.company))
		zip.add('docProps/core.xml', makeXmlCore(pres.title, pres.subject, pres.author, pres.revision))
		if (hasCustomProps) {
			zip.add('docProps/custom.xml', makeXmlCustomProperties(source.customProperties))
		}
		zip.add(relsPath(PRESENTATION_PATH), makeXmlPresentationRels(pres.slides, pres.embeddedFonts))
		// Embedded font parts (raw whole faces). Fonts are already compact binary, so STORE
		// (no DEFLATE) like already-compressed media. Part index matches the rels Target above.
		for (const face of flattenEmbeddedFaces(pres.embeddedFonts, 1)) {
			zip.add(fontPath(face.partIndex), face.bytes, { store: true })
		}
		zip.add('ppt/theme/theme1.xml', makeXmlTheme(pres))
		// emit a separate theme2.xml part so notesMaster1.xml.rels resolves
		zip.add('ppt/theme/theme2.xml', makeXmlTheme(pres))
		zip.add(PRESENTATION_PATH, makeXmlPresentation(pres))
		zip.add('ppt/presProps.xml', makeXmlPresProps())
		zip.add('ppt/tableStyles.xml', makeXmlTableStyles())
		zip.add('ppt/viewProps.xml', makeXmlViewProps())

		// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
		pres.slideLayouts.forEach((layout, idx) => {
			zip.add(slideLayoutPath(idx + 1), makeXmlLayout(layout, source.renderers))
			zip.add(relsPath(slideLayoutPath(idx + 1)), makeXmlSlideLayoutRel(idx + 1, pres.slideLayouts))
		})
		pres.slides.forEach((slide, idx) => {
			zip.add(slidePath(idx + 1), makeXmlSlide(slide, source.renderers, transitionSoundRIds.get(slide)))
			zip.add(relsPath(slidePath(idx + 1)), makeXmlSlideRel(pres.slides, pres.slideLayouts, idx + 1))
			contributors.forEach((contributor) => contributor.parts?.withEachSlide?.(slide, idx + 1, zip))
		})
		zip.add(SLIDE_MASTER_PATH, makeXmlMaster(pres.masterSlide, pres.slideLayouts, source.renderers))
		zip.add(relsPath(SLIDE_MASTER_PATH), makeXmlMasterRel(pres.masterSlide, pres.slideLayouts))
		contributors.forEach((contributor) => contributor.parts?.afterMaster?.(pres, zip))

		// D: Create all Rels (images, media, chart data). Per target the contributors go first, then
		// the media pass, which is the interleaving the zip has always had.
		const addTargetRelParts = (target: PartTarget): void => {
			contributors.forEach((contributor) => {
				const pending = contributor.parts?.fromTargetRels?.(target, zip)
				if (pending) partPromises.push(...pending)
			})
			createMediaParts(target, zip)
		}
		pres.slideLayouts.forEach(addTargetRelParts)
		pres.slides.forEach(addTargetRelParts)
		addTargetRelParts(pres.masterSlide)

		// E: Wait for the contributed async parts (if any), then snapshot the accumulated
		// parts in emission order. Zipping is deferred to `zipPackageParts`.
		return await Promise.all(partPromises).then(() => zip.entries())
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
	if (props.outputType) {
		// A: Output type selected by the caller or runtime adapter.
		return await zip.generate(props.outputType, { compression })
	} else {
		// B: Browser/neutral default: output a Blob as app/ms-pptx.
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
