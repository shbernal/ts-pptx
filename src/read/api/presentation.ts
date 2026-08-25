/**
 * Read-model entry point: `Presentation` wraps an `OpcPackage` and exposes a
 * navigable, typed view of the deck (slides → shapes → text), backed by the
 * live DOM so the same nodes can later be mutated.
 *
 * The class holds the model and the deck-level plumbing every operation needs — the part
 * index, the slide list, the copy registry that lets repeated imports from one source share
 * what they already copied, and the `p:sldIdLst` wiring. The operations themselves live in
 * `ops/`, one job per module, taking this deck as an argument. The slide-import machinery
 * used to sit here as ~280 lines of private methods and made that split hard to see: the
 * read model's own surface was outnumbered by one feature's internals.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { relativePartName, relsPartNameFor } from '../opc/partnames.js'
import {
	OOXML_NS,
	attr,
	createElement,
	firstChild,
	getElements,
	getOrAddChild,
	intValue,
	setAttr,
} from '../oxml/dom.js'
import { EMBEDDED_FONT_SLOTS } from '../../embedded-fonts.js'
import { flattenShape } from './ops/flatten.js'
import { Slide } from './slide.js'
import { SlideMaster } from './chrome.js'
import { wrapShapeElement, type AnyShape } from './shapes.js'
import { carryShapeAnimations } from './animation.js'
import {
	commentSchema,
	readCommentAuthors,
	readModernCommentAuthors,
	type CommentAuthor,
	type CommentSchema,
	type ModernCommentAuthor,
} from './comments.js'
import {
	readCoreProperties,
	readCustomProperties,
	type CoreProperties,
	type CustomProperty,
} from './document-properties.js'
import { readTagsForPart, type Tag } from './tags.js'
import type {
	AppendSlidesOptions,
	FromTemplateOptions,
	ImportShapeOptions,
	EmbeddedFontInfo,
	ImportSlideMastersOptions,
	ImportSlideOptions,
	ImportSlidesRequest,
	ImportedSlideMaster,
	LayoutHandle,
	SlideSize,
	SlideSource,
} from './presentation-types.js'
import { cSldName, nthShapeChild } from '../oxml/slide-dom.js'
import { computeRescale, rescaleSpTree, type RescaleTransform } from './ops/rescale.js'
import { carryTableStyles } from './ops/table-styles.js'
import { promoteMasters } from './ops/master-registry.js'
import { checkSelectionCopyable, copyPart, copySlidePart, newOwnedScope, type ImportContext } from './ops/part-copy.js'
// Deck-mutation operations. They live beside the model rather than on it: each is a whole job
// (prune a part fringe, carry notes, merge embedded fonts, rescale onto a new canvas) that reads
// and writes the package through the deck's public surface, and none of them is something a
// caller navigates *to*.
import { rewriteCarriedRels } from './ops/carried-rels.js'
import { carryEmbeddedFonts, carryGeneratedEmbeddedFonts } from './ops/embedded-fonts.js'
import { sourceFlattenContext } from './ops/flatten-context.js'
import { importSlidePreserve, importSlideRestyle } from './ops/import-slide.js'
import { carryNotes, ensureNotesMasterFromXml } from './ops/notes-master.js'
import { layoutPartNamesOf, slideMasterPartNames } from './ops/part-index.js'
import { duplicateOwnedTargets } from './ops/page-owned.js'
import { pruneIfOrphan } from './ops/prune.js'
import { rescaleImportedGeometry } from './ops/rescale-import.js'
import {
	IMAGE_REL,
	NOTES_MASTER_REL,
	NOTES_SLIDE_REL,
	OFFICE_DOCUMENT_REL,
	SLIDE_LAYOUT_REL,
	SLIDE_REL,
} from '../../ooxml/rel-types.js'
import { InternalError, InvalidOptionError, PackageReadError, UnsupportedFeatureError } from '../../errors.js'

const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
// Microsoft 2007 `media` rel: paired with the ECMA audio/video rel (same Target),
// referenced by the slide body's <p14:media r:embed>.
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const NOTES_SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/** Content type of the main part in an editable `.pptx` package. */
const PRESENTATION_MAIN_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
/** Content type of the main part in a `.potx` template package — flipped to {@link PRESENTATION_MAIN_CONTENT_TYPE} by {@link Presentation.fromTemplate}. */
const PRESENTATION_TEMPLATE_MAIN_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'

const textEncoder = new TextEncoder()

/**
 * `p:sldIdLst`'s document-order successors in `CT_Presentation` (ECMA-376):
 * everything that may legally follow it, so an inserted `p:sldIdLst` lands in the
 * right position when a template omitted it (zero slides).
 */
const PRESENTATION_SLD_ID_LST_SUCCESSORS = [
	'p:sldSz',
	'p:notesSz',
	'p:smartTags',
	'p:embeddedFontLst',
	'p:custShowLst',
	'p:photoAlbum',
	'p:custDataLst',
	'p:kinsoku',
	'p:defaultTextStyle',
	'p:modifyVerifier',
	'p:extLst',
]
/** ST_SlideId minimum (ECMA-376): slide ids live in [256, 2147483647]. */
const MIN_SLIDE_ID = 256

export class Presentation {
	#presentationPart: Part | undefined
	/**
	 * Per-source copy registry for {@link importSlide}: source `OpcPackage` →
	 * (source partname → partname allocated in this package). Lets parts shared
	 * across imports from the same source deck (layout, master, theme, media) be
	 * copied once and reused on later calls.
	 */
	#importRegistry = new Map<OpcPackage, Map<string, string>>()
	/**
	 * Parts whose geometry {@link importSlide}'s `rescale` has already rewritten, so a
	 * layout/master shared across repeated imports from one source is not scaled twice.
	 */
	#rescaledParts = new Set<string>()

	private constructor(readonly opc: OpcPackage) {}

	/** Open a `.pptx` from bytes and wrap it as a navigable `Presentation`. */
	static async load(input: OpcInput): Promise<Presentation> {
		return new Presentation(await OpcPackage.load(input))
	}

	/** Wrap an already-loaded OPC package (e.g. from the lower-level API). */
	static fromPackage(opc: OpcPackage): Presentation {
		return new Presentation(opc)
	}

	/**
	 * Open a PowerPoint template (`.pptx` or `.potx`) and return it as an empty
	 * deck shell ready to author onto: its slide masters, layouts, and theme are
	 * kept **byte-identical**, while any sample slides the template carried are
	 * stripped so only the shared chrome remains.
	 *
	 * Use it to build a fresh deck on a corporate template without rebuilding the
	 * masters in code: discover the bindable layouts with {@link layouts}, then
	 * author slides with a generator sized to match the template and graft them in
	 * with {@link appendSlides} (which enforces an equal slide size). Saving yields
	 * an editable `.pptx` that reuses the template's authored chrome verbatim.
	 *
	 * ```ts
	 * const deck = await Presentation.fromTemplate(templateBytes) // .pptx or .potx
	 * deck.layouts().map(l => l.name)                             // discover layouts
	 * await deck.appendSlides(pptx, { layout: 'Title and Content' })
	 * const out = await deck.save()                               // editable .pptx
	 * ```
	 *
	 * A `.potx` package declares its main part with the template content type; by
	 * default that override is flipped to the editable presentation type so the
	 * saved output opens as a normal deck. Pass `keepTemplateContentType: true` to
	 * preserve the template type. A `.pptx` input needs no flip, and a template
	 * that already carries zero slides makes the strip a no-op.
	 */
	static async fromTemplate(input: OpcInput, options: FromTemplateOptions = {}): Promise<Presentation> {
		const pres = new Presentation(await OpcPackage.load(input))

		// Normalize a .potx main part to the editable presentation content type so
		// the saved package opens as a deck, not a template. The officeDocument
		// relationship resolves the main part regardless of its content type, so a
		// .potx already loads; only the [Content_Types].xml override needs flipping.
		if (!options.keepTemplateContentType) {
			const mainPart = pres.presentationPart
			if (mainPart.contentType === PRESENTATION_TEMPLATE_MAIN_CONTENT_TYPE) {
				pres.opc.contentTypes.ensureRegistered(mainPart.partName, PRESENTATION_MAIN_CONTENT_TYPE)
			}
		}

		// Strip any sample slides to a master/layout-only shell. removeSlide never
		// prunes shared chrome, so masters/layouts/theme stay byte-identical.
		while (pres.slides.length > 0) pres.removeSlide(0)

		return pres
	}

	/** The main presentation part (`/ppt/presentation.xml`), resolved via the package `officeDocument` relationship. */
	get presentationPart(): Part {
		if (this.#presentationPart) return this.#presentationPart
		const packageRels = this.opc.relationshipsFor('/')
		const officeDocument = packageRels.byType(OFFICE_DOCUMENT_REL)
		const officeDocumentRel = officeDocument[0]
		if (officeDocument.length !== 1 || !officeDocumentRel) {
			throw new PackageReadError(
				'package/office-document-relationship-invalid',
				`Expected exactly one officeDocument relationship, found ${officeDocument.length}`
			)
		}
		const partName = packageRels.resolveTarget(officeDocumentRel.id)
		const part = this.opc.part(partName)
		if (!part)
			throw new PackageReadError(
				'package/relationship-target-missing',
				`officeDocument relationship targets a missing part: ${partName}`
			)
		this.#presentationPart = part
		return part
	}

	/** The slides in presentation order (resolved from `p:sldIdLst` + the presentation's relationships). */
	get slides(): Slide[] {
		const root = this.presentationPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (!sldIdLst) return []
		const rels = this.opc.relationshipsFor(this.presentationPart.partName)
		const slides: Slide[] = []
		let index = 0
		for (const sldId of getElements(sldIdLst, 'p:sldId')) {
			const relId = attr(sldId, 'r:id')
			if (!relId) continue
			const partName = rels.resolveTarget(relId)
			const part = this.opc.part(partName)
			if (!part)
				throw new PackageReadError(
					'package/relationship-target-missing',
					`Slide relationship ${relId} targets a missing part: ${partName}`
				)
			slides.push(new Slide(this, part, intValue(attr(sldId, 'id')) ?? 0, index++))
		}
		return slides
	}

	/** Slide dimensions (`p:sldSz`), or `null` if the presentation declares none. */
	get slideSize(): SlideSize | null {
		const root = this.presentationPart.dom.documentElement
		const sldSz = root && firstChild(root, 'p:sldSz')
		if (!sldSz) return null
		const widthEmu = intValue(attr(sldSz, 'cx'))
		const heightEmu = intValue(attr(sldSz, 'cy'))
		if (widthEmu === null || heightEmu === null) return null
		return { widthEmu, heightEmu, widthIn: emuToInches(widthEmu), heightIn: emuToInches(heightEmu) }
	}

	/**
	 * The deck's embedded font families (`p:embeddedFontLst` in `presentation.xml`),
	 * `[]` when it embeds none. Each entry names the typeface and resolves every
	 * embedded face's `r:id` to the absolute partname of its `.fntdata` binary — the
	 * read counterpart to the write-side `pptx.embedFont` / `importSlide({ embedFonts })`
	 * carry. An entry whose `p:font` has no `@typeface`, or a face whose `r:id` is
	 * missing or dangling, is skipped (faithful degradation, no throw). Read-only.
	 */
	get embeddedFonts(): EmbeddedFontInfo[] {
		const root = this.presentationPart.dom.documentElement
		const lst = root && firstChild(root, 'p:embeddedFontLst')
		if (!lst) return []
		const rels = this.opc.relationshipsFor(this.presentationPart.partName)
		const fonts: EmbeddedFontInfo[] = []
		for (const entry of getElements(lst, 'p:embeddedFont')) {
			const font = firstChild(entry, 'p:font')
			const typeface = font && attr(font, 'typeface')
			if (!typeface) continue
			const faces: EmbeddedFontInfo['faces'] = []
			for (const slot of EMBEDDED_FONT_SLOTS) {
				const face = firstChild(entry, `p:${slot}`)
				const relId = face && attr(face, 'r:id')
				if (!relId || !rels.get(relId)) continue
				faces.push({ slot, partName: rels.resolveTarget(relId) })
			}
			fonts.push({ typeface, panose: font ? attr(font, 'panose') : null, faces })
		}
		return fonts
	}

	/**
	 * The deck-wide **legacy** comment-author registry (`p:cmAuthorLst` in
	 * `ppt/commentAuthors.xml`), `[]` when the deck has no comments. Each slide's
	 * {@link Slide.comments} resolves its `@authorId` against this list. The 2018
	 * modern comment authors (`ppt/authors.xml`) are a separate part, not decoded here.
	 */
	get commentAuthors(): CommentAuthor[] {
		return readCommentAuthors(this.opc, this.presentationPart.partName)
	}

	/**
	 * The deck-wide **modern** (2018) comment-author registry (`p188:authorLst` in
	 * `ppt/authors.xml`), `[]` when the deck carries no modern comments. Unlike the
	 * legacy {@link commentAuthors}, each entry is keyed by a GUID `id` (plus
	 * `userId`/`providerId`); each slide's {@link Slide.modernComments} resolves its
	 * `@authorId` against this list. Read-only — the writer emits legacy comments.
	 */
	get modernCommentAuthors(): ModernCommentAuthor[] {
		return readModernCommentAuthors(this.opc, this.presentationPart.partName)
	}

	/**
	 * Which comment schema this deck uses: `'modern'` when it carries any 2018
	 * `modernComment_*` part (read via {@link Slide.modernComments} /
	 * {@link modernCommentAuthors}), `'legacy'` when it carries classic
	 * `commentN.xml` parts (read via {@link Slide.comments} / {@link commentAuthors}),
	 * or `'none'`. The two schemas do not coexist in practice, so this tells a
	 * consumer which accessor to read without probing both.
	 */
	get commentSchema(): CommentSchema {
		return commentSchema(this.opc)
	}

	/**
	 * The deck's core document properties (`docProps/core.xml`): title, subject,
	 * creator, keywords, revision, and the created/modified/lastPrinted timestamps
	 * (kept as raw W3CDTF strings). `{}` when the deck carries no core-properties
	 * part. The read counterpart of the write-side `pptx.title`/`subject`/`author`
	 * (→ `creator`)/`revision` setters.
	 */
	get coreProperties(): CoreProperties {
		return readCoreProperties(this.opc)
	}

	/**
	 * The deck's user-defined custom document properties (`docProps/custom.xml`) as
	 * `{ name, value }` pairs, each value typed from its `vt:` element (string,
	 * number, boolean, or a raw filetime string). `[]` when the deck carries no
	 * custom-properties part. The read counterpart of `pptx.setCustomProperty(...)`.
	 */
	get customProperties(): CustomProperty[] {
		return readCustomProperties(this.opc)
	}

	/**
	 * The deck-level programmatic tags (`p:custDataLst/p:tags` on `presentation.xml`,
	 * resolved to `ppt/tags/tagN.xml`) as `{ name, val }` string pairs, `[]` when the
	 * deck carries none. These are add-in/host metadata with no visible rendering and
	 * no writer — read-only, preserved byte-for-byte on round-trip. Per-slide tags are
	 * {@link Slide.tags}.
	 */
	get tags(): Tag[] {
		return readTagsForPart(this.opc, this.presentationPart.partName)
	}

	/**
	 * Duplicate the slide at `index` and insert the copy at `options.at` (deck
	 * order; `0` = first), defaulting to appending at the end when `at` is omitted
	 * or out of range. Returns the new slide. The new slide part copies the source
	 * bytes verbatim and shares the source's deck-wide relationship targets
	 * (layout, images, media) by copying its `.rels`; a new presentation→slide
	 * relationship and a `p:sldId` entry are wired up. Marks the presentation part
	 * dirty.
	 *
	 * What the source page *owns* is copied rather than shared: its notes slide,
	 * charts, SmartArt diagrams and OLE embeddings, each with the subtree under it
	 * (a chart's embedded workbook comes along; the image inside its user-shapes
	 * drawing stays shared). PowerPoint refuses to open a deck where two slides
	 * resolve to one chart or diagram, so this is not tidiness — see
	 * `page-owned.ts` for the rule and the evidence behind it.
	 */
	cloneSlide(index: number, options: { at?: number } = {}): Slide {
		const source = this.slides[index]
		if (!source) throw new InvalidOptionError('slide/index-out-of-range', `No slide at index ${index} to clone`)
		const opc = this.opc
		const sourcePart = source.part

		// 1. Copy the slide part bytes verbatim into a fresh slide partname.
		const newPartName = opc.reservePartNameLike(sourcePart.partName)
		const newPart = opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)

		// 2. Copy the slide's relationships (targets resolve identically — same dir).
		//    Through the live relationship set, not the `.rels` part bytes: a page
		//    this session imported or edited holds its rels in memory until the deck
		//    is saved, and copying the bytes gave such a clone no relationships at
		//    all — a slide whose `r:id`s resolved to nothing.
		const sourceRels = opc.relationshipsFor(sourcePart.partName)
		const cloneRels = opc.relationshipsFor(newPartName)
		for (const rel of sourceRels) cloneRels.addWithId(rel.id, rel.type, rel.target, rel.targetMode)

		// 3. Take copies of the parts the page owned, leaving the shared ones shared.
		duplicateOwnedTargets(opc, sourcePart.partName, newPartName)

		// 4. Wire the new slide into the presentation (rel + p:sldId entry) at `at`.
		return this.#insertSlidePart(newPart, options.at)
	}

	/**
	 * Remove the slide at `index` (deck order) and return its former partname. The
	 * `p:sldId` entry and the presentation→slide relationship are dropped, the slide
	 * part and its `.rels` are deleted, and any part the slide *privately* owned
	 * (its notes slide, slide-only media, charts/embeddings) that no remaining part
	 * references is pruned too — recursively. Shared deck chrome (layout, master,
	 * theme, …) is never pruned, so the deck stays renderable; removing every slide
	 * leaves a valid master/layout-only package (a template shell).
	 *
	 * Untouched parts stay byte-identical, matching the package fidelity contract.
	 * Throws when there is no slide at `index`.
	 */
	removeSlide(index: number): string {
		const slide = this.slides[index]
		if (!slide) throw new InvalidOptionError('slide/index-out-of-range', `No slide at index ${index} to remove`)
		const partName = slide.partName

		// The slide's internal targets, captured before its rels are dropped, so the
		// parts it privately owned can be pruned afterwards.
		const slideRels = this.opc.relationshipsFor(partName)
		const formerTargets = [...slideRels]
			.filter((rel) => rel.targetMode !== 'External')
			.map((rel) => slideRels.resolveTarget(rel.id))

		// Unwire from presentation.xml: remove the matching p:sldId and the rel.
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const root = presPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (sldIdLst) {
			for (const sldId of getElements(sldIdLst, 'p:sldId')) {
				const relId = attr(sldId, 'r:id')
				if (relId && presRels.get(relId) && presRels.resolveTarget(relId) === partName) {
					sldIdLst.removeChild(sldId)
					presRels.remove(relId)
					break
				}
			}
		}
		presPart.markDirty()

		// Drop the slide part and its .rels, then prune the parts it privately owned.
		this.opc.removePart(relsPartNameFor(partName))
		this.opc.removePart(partName)
		for (const target of formerTargets) pruneIfOrphan(this, target)

		return partName
	}

	/**
	 * Append a copy of `source.slides[index]` to this presentation and return it.
	 *
	 * Unlike {@link cloneSlide} (same-deck duplicate), this copies a slide across
	 * a package boundary: it brings the connected sub-graph the slide depends on —
	 * its `slideLayout` → `slideMaster` → `theme`, plus any media, charts, and
	 * embeddings — into this package under fresh partnames, rewriting every
	 * partname, relationship id, and content-type registration so the result is a
	 * self-consistent OPC package. Parts of this (target) package that are not
	 * touched stay byte-identical, matching `cloneSlide`'s fidelity contract.
	 *
	 * Only the layout(s) actually used by imported slides are copied; the imported
	 * master's `p:sldLayoutIdLst` is pruned to exactly those, mirroring how
	 * PowerPoint's "Reuse Slides" brings a slide across. Parts shared by repeated
	 * imports from the same source deck are copied once and reused — every part
	 * except the page itself. Importing the same source slide twice yields two
	 * independent copies over one shared layout/master/theme, which is what lets a
	 * deck show one page twice (verbatim, then edited).
	 *
	 * With `{ theme: 'preserve' }` the slide's source theme is instead *flattened*
	 * into the slide XML and the slide is bound to this deck's existing
	 * master/layout; with `{ theme: 'restyle' }` the slide is bound to this deck's
	 * master/layout with its theme references left symbolic, so it re-brands to the
	 * destination palette — see {@link ImportSlideOptions}.
	 *
	 * v1 limitations: by default the source slide size must equal this
	 * presentation's — pass `{ rescale: 'fit' | 'stretch' }` to rescale the imported
	 * geometry onto this deck's canvas instead (geometry only, not fonts/line
	 * widths). Source notes are dropped unless you pass `{ importNotes: true }`.
	 */
	importSlide(source: Presentation, index: number, options: ImportSlideOptions = {}): Slide {
		const sourceSlide = source.slides[index]
		if (!sourceSlide) throw new InvalidOptionError('slide/index-out-of-range', `No slide at index ${index} to import`)

		// 1. Pre-flight: slide sizes must match unless the caller opts into a rescale.
		const target = this.slideSize
		const incoming = source.slideSize
		const sizesDiffer =
			!target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu
		if (sizesDiffer && !options.rescale) {
			const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
			throw new InvalidOptionError(
				'import/slide-size-mismatch',
				`importSlide requires equal slide sizes (pass { rescale: 'fit' | 'stretch' } to rescale); target is ${fmt(target)}, source is ${fmt(incoming)}`
			)
		}
		if (sizesDiffer && options.rescale && (!target || !incoming)) {
			throw new InvalidOptionError(
				'import/slide-size-unknown',
				'importSlide rescale requires both decks to declare a slide size (p:sldSz)'
			)
		}

		// 2. Copy the slide and its dependencies. 'preserve' flattens the theme into
		//    the slide and attaches it to this deck's master; 'restyle' attaches it
		//    to this deck's master with theme refs left symbolic (re-brand); 'copy'
		//    brings the source theme subgraph across wholesale.
		const importCtx = this.#importContext(source.opc)
		const newPartName =
			options.theme === 'preserve'
				? importSlidePreserve(this, importCtx, source, sourceSlide, options.carryMasterGraphics === true)
				: options.theme === 'restyle'
					? importSlideRestyle(
							this,
							importCtx,
							source,
							sourceSlide,
							options.carryMasterGraphics === true,
							options.remapLiterals === true
						)
					: copySlidePart(importCtx, sourceSlide.partName)
		const newPart = this.opc.part(newPartName)
		if (!newPart)
			throw new InternalError('import/part-went-missing', `Imported slide part went missing: ${newPartName}`)

		// 2b. Rescale the imported geometry to this deck's canvas when sizes differ.
		if (sizesDiffer && options.rescale && target && incoming) {
			rescaleImportedGeometry(
				this,
				this.#rescaledParts,
				newPartName,
				options.theme,
				incoming,
				target,
				options.rescale === true ? 'fit' : options.rescale
			)
		}

		// 3. Wire the new slide into the presentation (rel + p:sldId entry) at `at`.
		const slide = this.#insertSlidePart(newPart, options.at)

		// 4. Optionally carry the source slide's speaker notes. The slide copy above
		//    drops the notesSlide rel (both copyPart and importSlideRebind do); this
		//    re-adds it wired to the new slide and merged onto a single notesMaster.
		if (options.importNotes) carryNotes(this, source, importCtx, sourceSlide.partName, newPartName)

		// 5. Optionally carry the source deck's embedded fonts (presentation-level, so
		//    a separate traversal from the slide-part copy chain above).
		if (options.embedFonts) carryEmbeddedFonts(this, source, importCtx)

		return slide
	}

	/**
	 * Import selected pages from one or more loaded source presentations as one
	 * batch. Each imported page lands at its `outputIndex` in the complete
	 * destination slide list after the batch, and the returned array is parallel
	 * to `requests` — `result[i]` is the page `requests[i]` asked for, whatever
	 * order the output positions were given in.
	 *
	 * Everything is checked before a single byte of this deck moves: every request
	 * names an existing source page, final output positions are unique, each source
	 * has this deck's slide size, and a read-only dry run of the copy proves every
	 * part it would reach is present. A batch therefore either applies in full or
	 * leaves this deck byte-identical, where a per-page loop of
	 * {@link importSlide} can leave a half-stitched deck behind.
	 *
	 * The batch also decides what a `slide → slide` link means: an internal link
	 * on a selected page must target another **selected** page (or one this deck
	 * already contains via an earlier import from that source), and is rewritten
	 * to the fresh partname — importing page 3 of 10 does not drag pages 1–2
	 * across as dependencies, and never strands the link. This mirrors
	 * `appendSlides`' `import/unresolved-slide-link`, which the write side already
	 * enforces for generator decks.
	 *
	 * One request is one output page, so naming the same source page in several
	 * requests is how you ask for several independent copies of it — the page part
	 * is the one thing an import never shares, exactly as in {@link importSlide}.
	 * Everything under it (layout, master, theme, media) is still copied once and
	 * shared. Where such a page is the target of a `slide → slide` link, the link
	 * resolves to one of its copies: pages duplicated together are copied in
	 * lockstep and link to their round-mates, and a link into a page requested only
	 * once always lands on that single copy.
	 *
	 * Speaker notes travel per request: `{ importNotes: true }` carries that page's
	 * `notesSlide` part across, wired to the new page and bound to a single
	 * `notesMaster` under the same 0..1 rule {@link importSlide} and
	 * {@link appendSlides} follow — the destination's own master wins when it has
	 * one. Notes are part of the up-front dry run too, so a batch that would fail
	 * carrying them is refused with the deck still byte-identical. A page named in
	 * several requests gets its own copy of its notes each time, as of everything
	 * else that page owns.
	 *
	 * Scope: pages come across under `'copy'` theme semantics (their own layout →
	 * master → theme subgraph, shared parts deduped via the copy registry).
	 * Embedded fonts are not carried, and sizes must match, since there is no batch
	 * spelling for `embedFonts` or `rescale`. Use {@link importSlide} when you need
	 * either of those.
	 */
	importSlides(requests: readonly ImportSlidesRequest[]): Slide[] {
		// 1. Validate everything up front: indexes exist, selections and output
		//    positions are unique, sizes match. No part is touched until all pass.
		const resolved = requests.map((request, requestIndex) => {
			if (!Number.isInteger(request.sourceIndex) || request.sourceIndex < 0) {
				throw new InvalidOptionError(
					'slide/index-out-of-range',
					`importSlides: sourceIndex must be a non-negative integer; received ${request.sourceIndex}`
				)
			}
			if (!Number.isInteger(request.outputIndex) || request.outputIndex < 0) {
				throw new InvalidOptionError(
					'import/output-index-out-of-range',
					`importSlides: outputIndex must be a non-negative integer; received ${request.outputIndex}`
				)
			}
			const sourceSlide = request.source.slides[request.sourceIndex]
			if (!sourceSlide) {
				throw new InvalidOptionError(
					'slide/index-out-of-range',
					`importSlides: no slide at index ${request.sourceIndex} to import`
				)
			}
			return { ...request, requestIndex, sourceSlide }
		})
		if (resolved.length === 0) return []

		// Step 4 inserts into this element; failing on it here rather than there
		// keeps the wiring loop, like the copy, unable to stop half-way.
		if (!this.presentationPart.dom.documentElement) {
			throw new PackageReadError(
				'package/part-has-no-root',
				'presentation.xml has no document element to insert slides into'
			)
		}

		const finalSlideCount = this.slides.length + resolved.length
		// Which pages each source is being asked for. A page may appear in several
		// requests: the set is what the dry run walks, and the per-request output
		// parts are allocated in step 2. `notesPages` is the subset whose notes are
		// coming too, which the dry run has to walk past the dropped notes rel.
		const selectedPages = new Map<OpcPackage, Set<string>>()
		const notesPages = new Map<OpcPackage, Set<string>>()
		for (const request of resolved) {
			let pages = selectedPages.get(request.source.opc)
			if (!pages) {
				pages = new Set()
				selectedPages.set(request.source.opc, pages)
			}
			pages.add(request.sourceSlide.partName)
			if (!request.importNotes) continue
			let withNotes = notesPages.get(request.source.opc)
			if (!withNotes) {
				withNotes = new Set()
				notesPages.set(request.source.opc, withNotes)
			}
			withNotes.add(request.sourceSlide.partName)
		}

		const target = this.slideSize
		const outputIndexes = new Set<number>()
		for (const request of resolved) {
			if (request.outputIndex >= finalSlideCount) {
				throw new InvalidOptionError(
					'import/output-index-out-of-range',
					`importSlides: outputIndex ${request.outputIndex} is outside the final slide list of ${finalSlideCount} slides`
				)
			}
			if (outputIndexes.has(request.outputIndex)) {
				throw new InvalidOptionError(
					'import/output-index-conflict',
					`importSlides: outputIndex ${request.outputIndex} is requested more than once`
				)
			}
			outputIndexes.add(request.outputIndex)
			const incoming = request.source.slideSize
			if (!target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu) {
				const fmt = (size: SlideSize | null): string => (size ? `${size.widthEmu}×${size.heightEmu} EMU` : 'unknown')
				throw new InvalidOptionError(
					'import/slide-size-mismatch',
					`importSlides requires equal slide sizes; target is ${fmt(target)}, source is ${fmt(incoming)}`
				)
			}
		}

		// 1b. Dry-run the copy against each source, still reading only source
		//     packages: every part the traversal will reach exists and parses, and
		//     no selected page links outside the selection. Once this passes the
		//     copy below has no reachable throw, which is what lets a rejected
		//     batch leave this deck byte-identical instead of half-stitched.
		//     `copyMaster` is read once, before anything moves: a destination that
		//     already has a notesMaster keeps it, so no source master is copied at
		//     all, and one that has none takes the first carried master — after which
		//     the rest bind to it. Walking every source's master when the deck has
		//     none is deliberately the strict side of that: it can only reject a
		//     source deck whose own notes master is already broken.
		const copyMaster = this.opc.relationshipsFor(this.presentationPart.partName).byType(NOTES_MASTER_REL).length === 0
		for (const [sourceOpc, pages] of selectedPages) {
			checkSelectionCopyable(sourceOpc, this.#importContext(sourceOpc).registry, pages, {
				pages: notesPages.get(sourceOpc) ?? new Set(),
				copyMaster,
			})
		}

		// 2. Materialize each request's output page now, so the copy traversals can
		//    wire slide→slide relationships to their pre-allocated destinations.
		//    One request is one output page, so a source page asked for twice gets
		//    two reservations, in request order.
		//    The reservation list is typed non-empty, so round 0 needs no fallback.
		const destinationsBySource = new Map<OpcPackage, Map<string, [string, ...string[]]>>()
		const planned = resolved.map((request) => {
			let destinations = destinationsBySource.get(request.source.opc)
			if (!destinations) {
				destinations = new Map()
				destinationsBySource.set(request.source.opc, destinations)
			}
			const sourcePart = request.sourceSlide.part
			const newPartName = this.opc.reservePartNameLike(request.sourceSlide.partName)
			const destPart = this.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
			const reserved = destinations.get(request.sourceSlide.partName)
			if (reserved) reserved.push(newPartName)
			else destinations.set(request.sourceSlide.partName, [newPartName])
			return { ...request, destPart }
		})

		// 3. Copy each selected page and its dependency subgraph (theme/master/
		//    layout/media/…), with links constrained to the selection.
		//
		//    `copyPart`'s plan holds one destination per source page, so a page
		//    requested N times is copied in N rounds: round K materializes every
		//    page that has a Kth reservation, and names each other page's *first*
		//    copy so a jump link out of the round still lands on a page of this
		//    batch. A page duplicated alongside another therefore links to its
		//    round-mate, and a link into a single-copy page resolves to that one
		//    copy from every round. Rounds after the first re-materialize only the
		//    pages they name: everything else is a registry hit `copyPart` returns
		//    unchanged.
		for (const [sourceOpc, destinations] of destinationsBySource) {
			const base = this.#importContext(sourceOpc)
			const rounds = Math.max(...[...destinations.values()].map((reserved) => reserved.length))
			for (let round = 0; round < rounds; round++) {
				const plan = new Map<string, string>()
				for (const [sourcePartName, reserved] of destinations) plan.set(sourcePartName, reserved[round] ?? reserved[0])
				const ctx: ImportContext = { ...base, selection: { destinations: plan } }
				for (const [sourcePartName, reserved] of destinations) {
					if (round < reserved.length) void copyPart(ctx, sourcePartName)
				}
			}
		}

		// 3b. Carry the notes of the pages that asked for them, in request order, so
		//     the deck's single notesMaster comes from the first such page — the same
		//     order the dry run assumed. The copy above dropped every notesSlide rel,
		//     so this is the only thing that re-adds one, and a page named twice gets
		//     a notes part per copy.
		for (const request of planned) {
			if (!request.importNotes) continue
			carryNotes(
				this,
				request.source,
				this.#importContext(request.source.opc),
				request.sourceSlide.partName,
				request.destPart.partName
			)
		}

		// 4. Wire into p:sldIdLst at each requested final position. Ascending order
		//    makes the raw final index the correct insertion point at every step:
		//    earlier inserts all sit before it, so they shift it by exactly the
		//    number of entries the final position already counts. The result is
		//    written back by request index, so `result[i]` is `requests[i]`'s page
		//    whatever order the positions were given in.
		const added: Slide[] = []
		for (const request of [...planned].sort((left, right) => left.outputIndex - right.outputIndex)) {
			added[request.requestIndex] = this.#insertSlidePart(request.destPart, request.outputIndex)
		}
		return added
	}

	/**
	 * Graft slide master(s) from another open package into this one and return what
	 * was copied. Unlike {@link importSlide} — which brings a master across only as
	 * the dependency of an imported *slide* and prunes it to the one layout that
	 * slide uses — this copies a master together with its **whole** layout family
	 * and attaches it to no slide: the master and its layouts land in this deck's
	 * layout gallery (PowerPoint's *Insert ▸ New Slide* / *Layout* picker) without
	 * changing any existing slide.
	 *
	 * It is the "ship a brand template's layouts into a generated deck" capability,
	 * kept brand-agnostic here: the caller supplies the source `.pptx`. Each grafted
	 * master is wired into `p:sldMasterIdLst` (so renderers treat it as active) and
	 * its `p:sldLayoutIdLst` is rebuilt to list exactly the copied layouts; the
	 * connected theme/media/tag parts come across under fresh partnames, and parts
	 * shared with earlier imports from the same source are reused (the copy
	 * registry), so a re-call is idempotent. Untouched parts of this package stay
	 * byte-identical, matching {@link importSlide}'s fidelity contract.
	 *
	 * `options.masters` / `options.layouts` narrow what is grafted; by default every
	 * master and every layout comes across. The source and destination slide sizes
	 * must match unless `options.requireEqualSize` is `false` (see
	 * {@link ImportSlideMastersOptions}).
	 *
	 * Presentation-level parts are carried only on request: embedded fonts via
	 * `options.embedFonts` and table styles via `options.tableStyles`. By default a
	 * grafted master is appended after the deck's existing masters; `options.primary`
	 * moves the grafted masters to the front so the deck presents as their theme (see
	 * {@link ImportSlideMastersOptions}). The v1 limitation mirroring
	 * {@link importSlide} is that geometry is not rescaled.
	 */
	importSlideMasters(source: Presentation, options: ImportSlideMastersOptions = {}): ImportedSlideMaster[] {
		if (options.requireEqualSize !== false) {
			const target = this.slideSize
			const incoming = source.slideSize
			if (!target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu) {
				const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
				throw new InvalidOptionError(
					'import/slide-size-mismatch',
					`importSlideMasters requires equal slide sizes (pass { requireEqualSize: false } to override); target is ${fmt(target)}, source is ${fmt(incoming)}`
				)
			}
		}

		const pickMaster = options.masters ?? (() => true)
		const pickLayout = options.layouts ?? (() => true)

		const ctx = this.#importContext(source.opc)
		const imported: ImportedSlideMaster[] = []
		slideMasterPartNames(source).forEach((masterPartName, masterIndex) => {
			if (!pickMaster(cSldName(source.opc.part(masterPartName)), masterIndex)) return

			// Copy the (lean) master first: copyPart registers it in p:sldMasterIdLst
			// and clears its layout list, then each copied layout re-links itself in.
			const newMasterPartName = copyPart(ctx, masterPartName)

			const layoutPartNames: string[] = []
			layoutPartNamesOf(source, masterPartName).forEach((layoutPartName, layoutIndex) => {
				if (!pickLayout(cSldName(source.opc.part(layoutPartName)), layoutIndex)) return
				layoutPartNames.push(copyPart(ctx, layoutPartName))
			})

			imported.push({ partName: newMasterPartName, layoutPartNames })
		})

		// Optionally carry the source deck's presentation-level styling parts. Both are
		// separate traversals from the master/layout copy chain above, and both are
		// whole-deck: neither part records which font/style belongs to which master.
		if (options.embedFonts) carryEmbeddedFonts(this, source, this.#importContext(source.opc))
		if (options.tableStyles) carryTableStyles(this, source.opc)
		if (options.primary)
			promoteMasters(
				this,
				imported.map((m) => m.partName)
			)

		return imported
	}

	/**
	 * The deck's slide masters, in `p:sldMasterIdLst` order, as modeled
	 * {@link SlideMaster}s — the typed read model over the shared chrome (each
	 * master's colour map, theme, placeholders, and the layouts built on it). Walk
	 * `pres.masters()[i].layouts` for the rich layout model, or use {@link layouts}
	 * for the flat {@link LayoutHandle} gallery {@link appendSlides} binds to.
	 * Read-only: it copies nothing and leaves the package byte-identical.
	 */
	masters(): SlideMaster[] {
		return slideMasterPartNames(this)
			.map((partName) => this.opc.part(partName))
			.filter((part): part is Part => part !== undefined)
			.map((part) => new SlideMaster(this.opc, part))
	}

	/**
	 * The deck's slide layouts, in master then layout order — the gallery a new
	 * slide can bind to. Each {@link LayoutHandle} addresses one layout for
	 * {@link appendSlides}; the `name` is its `p:cSld@name`. Read-only enumeration:
	 * it copies nothing and leaves the package byte-identical.
	 */
	layouts(): LayoutHandle[] {
		const out: LayoutHandle[] = []
		slideMasterPartNames(this).forEach((masterPartName, masterIndex) => {
			layoutPartNamesOf(this, masterPartName).forEach((layoutPartName, layoutIndex) => {
				out.push({
					partName: layoutPartName,
					name: cSldName(this.opc.part(layoutPartName)),
					masterPartName,
					masterIndex,
					layoutIndex,
				})
			})
		})
		return out
	}

	/**
	 * Append generator-produced slides onto this deck, binding each to an existing
	 * layout, and return the new {@link Slide}s. This is the hybrid
	 * "generate-onto-existing" path: the deck's masters, layouts, theme — and every
	 * other untouched part — stay **byte-identical** (only `presentation.xml`, its
	 * `.rels`, `[Content_Types].xml`, and the freshly-added slide/media parts
	 * change), because the existing chrome is never regenerated.
	 *
	 * `source` is any slide producer (a `TsPptx` instance); its authored slides
	 * are serialized via {@link SlideSource.extractSlides} and spliced in under
	 * fresh partnames, with each slide's `slideLayout` relationship pointed at the
	 * layout named by `options.layout` and its image/hyperlink relationships rebuilt
	 * (preserving the body's relationship ids). Insert position follows
	 * `options.at` (see {@link AppendSlidesOptions}).
	 *
	 * Charts and internal slide-to-slide hyperlinks are carried across: chart parts
	 * (chart XML + `.rels` + embedded workbook) are injected under fresh names, and a
	 * `slide:N` link is repointed at the Nth appended slide's new partname.
	 *
	 * The generator's presentation-level embedded fonts (`pptx.embedFont`) are also
	 * carried into this deck and merged into its `p:embeddedFontLst`, de-duped by
	 * typeface + face slot — so author-side embedded fonts survive the append onto a
	 * template that may itself already embed fonts.
	 *
	 * Embedded audio/video is carried too: the media part backs the ECMA audio/video rel
	 * and the MS-2007 `media` rel sharing one Target, plus a separate preview image part,
	 * with the media content type registered as a Default extension entry (what PowerPoint
	 * authors). Online (external-link) video rides as two External rels over the link.
	 *
	 * Speaker notes ride across too: a slide authored with `addNotes` gets a `notesSlide`
	 * part wired back to it, with its hyperlink rels preserved. A notes slide must bind to
	 * a notes master, and a template commonly has none, so the generator's is installed
	 * (with the theme its `.rels` needs) — but only when this deck has none of its own, so
	 * an existing notes master and its styling always win.
	 *
	 * Limitations:
	 * - An internal link to a source slide outside the appended batch throws (its
	 *   target has no counterpart in the destination).
	 * - Appended slides are concrete absolute-positioned content with no placeholder
	 *   inheritance from the bound layout; the binding governs theme/`clrMap`
	 *   resolution and the "based on" link, not placeholder geometry. Author with
	 *   concrete colours — any `schemeClr` re-resolves against the destination theme.
	 * - Source and destination slide sizes must match (no geometry rescale).
	 */
	async appendSlides(source: SlideSource, options: AppendSlidesOptions): Promise<Slide[]> {
		// 1. Resolve the target layout partname (explicit; no silent fallback).
		const gallery = this.layouts()
		let target: LayoutHandle
		if (typeof options.layout === 'string') {
			const matches = gallery.filter((l) => l.name === options.layout)
			if (matches.length > 1) {
				throw new InvalidOptionError(
					'layout/ambiguous-name',
					`appendSlides: layout name ${JSON.stringify(options.layout)} is ambiguous (${matches.length} layouts share it); pass a LayoutHandle from layouts() instead`
				)
			}
			const [only] = matches
			if (!only) {
				const names = gallery.map((l) => JSON.stringify(l.name)).join(', ')
				throw new InvalidOptionError(
					'layout/not-found',
					`appendSlides: no layout named ${JSON.stringify(options.layout)}; available: ${names || '(none)'}`
				)
			}
			target = only
		} else {
			const handle = options.layout
			if (!gallery.some((l) => l.partName === handle.partName)) {
				throw new InvalidOptionError(
					'layout/foreign-handle',
					`appendSlides: layout ${handle.partName} does not belong to this presentation`
				)
			}
			target = handle
		}

		// 2. Author + extract; enforce equal slide size (no geometry rescale in v1).
		const extracted = await source.extractSlides({ onMediaError: options.onMediaError })
		const size = this.slideSize
		if (!size || size.widthEmu !== extracted.widthEmu || size.heightEmu !== extracted.heightEmu) {
			const fmt = (w: number, h: number): string => `${w}×${h} EMU`
			throw new InvalidOptionError(
				'import/slide-size-mismatch',
				`appendSlides requires equal slide sizes; target is ${size ? fmt(size.widthEmu, size.heightEmu) : 'unknown'}, source is ${fmt(extracted.widthEmu, extracted.heightEmu)}`
			)
		}

		// Any existing slide partname seeds the fresh-partname family; fall back to a
		// literal seed for a slide-less template shell (reservePartNameLike parses the
		// string, it does not require the part to exist).
		const slideTemplate = this.slides[0]?.partName ?? '/ppt/slides/slide1.xml'

		// Pass 1: reserve + add every slide body first, so internal slide-to-slide
		// links (which may point forward) can resolve to any appended slide. Adding
		// each part immediately claims its name — reservePartNameLike returns max+1
		// from the existing parts, so the next reservation sees it. (addPart registers
		// the slide's Override content type.)
		const placed = extracted.slides.map((slide) => {
			const partName = this.opc.reservePartNameLike(slideTemplate)
			const part = this.opc.addPart(partName, SLIDE_CONTENT_TYPE, textEncoder.encode(slide.xml))
			return { slide, part, partName }
		})

		// 1-based source slide number -> the appended slide's new partname.
		const partBySourceNumber = new Map<number, string>(placed.map((p, i) => [i + 1, p.partName]))

		// Pass 2: build each slide's .rels and wire it into presentation.xml. Media,
		// hyperlinks, charts, and slide-links keep the body's rId (addWithId); the
		// layout rel is added last via add() so its auto-id cannot collide.
		const added: Slide[] = []
		placed.forEach(({ slide, part, partName }, i) => {
			const rels = this.opc.relationshipsFor(partName)
			for (const m of slide.media) {
				const mediaPartName = this.opc.reserveMediaPartName(m.extn)
				this.opc.addPart(mediaPartName, m.contentType, m.bytes)
				rels.addWithId(`rId${m.rId}`, IMAGE_REL, relativePartName(partName, mediaPartName))
			}
			for (const av of slide.avMedia) {
				// One media part backs two rels (ECMA audio/video + MS-2007 media) sharing
				// its Target; the preview poster is a separate image part. ensureDefault
				// runs before addPart so the content type resolves via a Default extension
				// entry (what PowerPoint authors) rather than a per-part Override.
				const mediaPartName = this.opc.reserveMediaPartName(av.mediaExtn, 'media')
				this.opc.contentTypes.ensureDefault(av.mediaExtn, av.mediaContentType)
				this.opc.addPart(mediaPartName, av.mediaContentType, av.mediaBytes)
				const mediaTarget = relativePartName(partName, mediaPartName)
				rels.addWithId(`rId${av.mediaRid}`, av.mtype === 'audio' ? AUDIO_REL : VIDEO_REL, mediaTarget)
				rels.addWithId(`rId${av.msMediaRid}`, MS_MEDIA_REL, mediaTarget)

				const previewPartName = this.opc.reserveMediaPartName(av.previewExtn)
				this.opc.contentTypes.ensureDefault(av.previewExtn, av.previewContentType)
				this.opc.addPart(previewPartName, av.previewContentType, av.previewBytes)
				rels.addWithId(`rId${av.previewRid}`, IMAGE_REL, relativePartName(partName, previewPartName))
			}
			for (const ov of slide.onlineMedia) {
				// Online (external-link) video: two External rels share the link Target — the
				// ECMA video rel and the MS-2007 media rel — with no media binary part and no
				// content-type entry. The poster image is wired by the `slide.media` loop above.
				rels.addWithId(`rId${ov.mediaRid}`, VIDEO_REL, ov.link, 'External')
				rels.addWithId(`rId${ov.msMediaRid}`, MS_MEDIA_REL, ov.link, 'External')
			}
			for (const h of slide.hyperlinks) {
				rels.addWithId(`rId${h.rId}`, HYPERLINK_REL, h.target, 'External')
			}
			if (slide.notes) {
				// Speaker notes. The notes part carries its own rel namespace, independent of
				// the slide's: rId1 = notesMaster, rId2 = the slide it annotates, hyperlinks
				// from rId3 — the order the generator's body was serialized against, so these
				// are added by explicit id rather than left to auto-numbering.
				const notesPartName = this.opc.reservePartNameLike('/ppt/notesSlides/notesSlide1.xml')
				this.opc.addPart(notesPartName, NOTES_SLIDE_CONTENT_TYPE, textEncoder.encode(slide.notes.xml))
				rels.add(NOTES_SLIDE_REL, relativePartName(partName, notesPartName))

				const notesRels = this.opc.relationshipsFor(notesPartName)
				const notesMasterPartName = extracted.notesMaster ? ensureNotesMasterFromXml(this, extracted.notesMaster) : null
				if (notesMasterPartName) {
					notesRels.addWithId('rId1', NOTES_MASTER_REL, relativePartName(notesPartName, notesMasterPartName))
				}
				notesRels.addWithId('rId2', SLIDE_REL, relativePartName(notesPartName, partName))
				for (const h of slide.notes.hyperlinks) {
					notesRels.addWithId(`rId${h.rId}`, HYPERLINK_REL, h.target, 'External')
				}
			}
			for (const c of slide.charts) {
				// Chart part + its embedded workbook, each under a fresh name. The chart
				// XML references the workbook through the chart part's own rId1, so the
				// chart .rels is rebuilt here against the reserved workbook partname.
				const chartPartName = this.opc.reservePartNameLike('/ppt/charts/chart1.xml')
				this.opc.addPart(chartPartName, CHART_CONTENT_TYPE, textEncoder.encode(c.chartXml))
				const embeddingPartName = this.opc.reservePartNameLike('/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
				this.opc.contentTypes.ensureDefault('xlsx', XLSX_CONTENT_TYPE)
				this.opc.addPart(embeddingPartName, XLSX_CONTENT_TYPE, c.embeddingBytes)
				this.opc
					.relationshipsFor(chartPartName)
					.addWithId('rId1', PACKAGE_REL, relativePartName(chartPartName, embeddingPartName))
				rels.addWithId(`rId${c.rId}`, CHART_REL, relativePartName(partName, chartPartName))
			}
			for (const link of slide.slideLinks) {
				const targetPartName = partBySourceNumber.get(link.sourceSlideNumber)
				if (!targetPartName) {
					throw new InvalidOptionError(
						'import/unresolved-slide-link',
						`appendSlides: slide ${i} links to source slide ${link.sourceSlideNumber}, which is not among the appended slides`
					)
				}
				rels.addWithId(`rId${link.rId}`, SLIDE_REL, relativePartName(partName, targetPartName))
			}
			rels.add(SLIDE_LAYOUT_REL, relativePartName(partName, target.partName))

			// Wire into presentation.xml (rel + p:sldId) at the requested position.
			const at = options.at === undefined ? undefined : options.at + i
			added.push(this.#insertSlidePart(part, at))
		})

		// Carry the generator's presentation-level embedded fonts (pptx.embedFont) into
		// this deck, so author-side embedded fonts survive the append onto a template.
		carryGeneratedEmbeddedFonts(this, extracted.embeddedFonts || [])

		return added
	}

	/**
	 * Copy one shape — an autoshape, picture, table/chart graphic frame, connector,
	 * or group — from `source.shapes[shapeIndex]` onto `target`, returning the new
	 * {@link Shape}. `target` must be a slide of *this* presentation; `source` may
	 * belong to any open presentation.
	 *
	 * The lifted subtree is copied self-consistently: every media/chart/embedding it
	 * depends on is dragged into this package (deduped against earlier imports from
	 * the same source), its `r:embed`/`r:id`/… are rewritten to fresh host-slide
	 * relationships, and its drawing ids (including a group's children) are reassigned
	 * so they cannot collide with the host. With `theme: 'preserve'` (default) the
	 * shape's theme references are baked to literals against the *source* theme so it
	 * renders the same on a foreign host; `restyle` leaves them symbolic to re-brand;
	 * `copy` brings the XML across untouched — see {@link ImportShapeOptions}.
	 *
	 * Differing slide sizes need `{ rescale }` (see {@link ImportShapeOptions.rescale});
	 * a lifted `preserve` placeholder is baked self-contained and demoted to a plain
	 * shape (see {@link ImportShapeOptions.theme}), so it neither re-inherits from nor
	 * collides with the host. A shape's build animation lives in the slide-scoped
	 * `p:timing`, not in its subtree, so the shape lands static unless
	 * `{ carryAnimation: true }` opts in (see {@link ImportShapeOptions.carryAnimation}).
	 */
	importShape(target: Slide, source: Slide, shapeIndex: number, options: ImportShapeOptions = {}): AnyShape {
		const [shape] = this.importShapes(target, source, [shapeIndex], options)
		if (!shape)
			throw new InvalidOptionError(
				'shape/index-out-of-range',
				`importShape: source slide has no shape at index ${shapeIndex}`
			)
		return shape
	}

	/**
	 * Batch form of {@link importShape}: copy several shapes from one source slide
	 * onto `target` in the given order. Media/chart/embedding parts shared by the
	 * lifted shapes (and by earlier imports from the same source deck) are copied
	 * once via the copy registry, and shared images resolve to a single host-slide
	 * relationship. Returns the new {@link Shape}s in `shapeIndices` order.
	 */
	importShapes(target: Slide, source: Slide, shapeIndices: number[], options: ImportShapeOptions = {}): AnyShape[] {
		if (target.presentation !== this)
			throw new InvalidOptionError('slide/foreign-target', 'importShape: target slide must belong to this presentation')

		// Pre-flight: slide sizes must match unless { rescale } opts into scaling the
		// lifted geometry onto this canvas (computed once, applied per shape below).
		const targetSize = this.slideSize
		const sourceSize = source.presentation.slideSize
		const sizesMatch =
			!!targetSize &&
			!!sourceSize &&
			targetSize.widthEmu === sourceSize.widthEmu &&
			targetSize.heightEmu === sourceSize.heightEmu
		let transform: RescaleTransform | null = null
		if (!sizesMatch) {
			if (!options.rescale || !targetSize || !sourceSize) {
				const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
				throw new InvalidOptionError(
					'import/slide-size-mismatch',
					`importShape requires equal slide sizes (or { rescale }); target is ${fmt(targetSize)}, source is ${fmt(sourceSize)}`
				)
			}
			transform = computeRescale(sourceSize, targetSize, options.rescale === 'stretch' ? 'stretch' : 'fit')
		}

		// Resolve + validate every index up front so a bad batch throws before mutating.
		const sourceShapes = source.shapes
		const sourceElements = shapeIndices.map((i) => {
			const shape = sourceShapes[i]
			if (!shape)
				throw new InvalidOptionError(
					'shape/index-out-of-range',
					`No shape at index ${i} on the source slide (it has ${sourceShapes.length})`
				)
			return shape.element_
		})

		const spTree = target.shapeTree()
		if (!spTree)
			throw new PackageReadError(
				'slide/no-shape-tree',
				`importShape: target slide ${target.partName} has no shape tree`
			)
		const targetDoc = spTree.ownerDocument
		if (!targetDoc)
			throw new InternalError('oxml/node-has-no-document', 'importShape: target slide DOM has no owner document')

		const theme = options.theme ?? 'preserve'
		const sourceOpc = source.presentation.opc
		const sourceRels = sourceOpc.relationshipsFor(source.partName)
		const targetRels = this.opc.relationshipsFor(target.partName)
		// One rel-id map across the batch so shapes sharing a source image share a rel.
		const relIdMap = new Map<string, string>()
		// preserve: build the source theme context once; copy/restyle need none.
		const ctx = theme === 'preserve' ? sourceFlattenContext(sourceOpc, source.partName) : null
		const importCtx = this.#importContext(sourceOpc)

		// Anchor for z-order: the existing shape currently at `at` (insert before it,
		// preserving batch order), else append before any trailing p:extLst.
		const extLst = firstChild(spTree, 'p:extLst')
		const anchor = options.at == null ? extLst : (nthShapeChild(spTree, options.at) ?? extLst)

		const result: AnyShape[] = []
		for (const shapeEl of sourceElements) {
			const imported = targetDoc.importNode(shapeEl, true)

			// Drag media/charts/embeddings across and rewrite refs to fresh host rels.
			// A scope per shape: media are shared through `relIdMap`, but the chart or
			// diagram under this frame is its own — importing one chart shape twice must
			// not point both frames at one chart part (see `page-owned.ts`).
			rewriteCarriedRels(imported, importCtx, sourceRels, target.partName, targetRels, relIdMap, newOwnedScope())

			// preserve: bake the source theme onto the subtree. The flatten passes match
			// descendants (not the root), so wrap the shape in a throwaway container.
			if (ctx) {
				const holder = createElement(targetDoc, 'p:spTree')
				holder.appendChild(imported)
				flattenShape(holder, ctx)
			}

			// Rescale geometry onto this canvas (after flatten, so a placeholder's just-baked
			// inherited a:xfrm is scaled too). rescaleSpTree wants a p:spTree, so wrap the shape.
			if (transform) {
				const holder = createElement(targetDoc, 'p:spTree')
				holder.appendChild(imported)
				rescaleSpTree(holder, transform)
			}

			// Give the shape and any group children collision-free host ids, recording the
			// source id → new id map so a carried build animation can be remapped onto it.
			let nextId = target.nextShapeId()
			const spidMap = new Map<number, number>()
			for (const cNvPr of imported.getElementsByTagNameNS(OOXML_NS.p, 'cNvPr')) {
				const oldId = intValue(attr(cNvPr, 'id'))
				if (oldId !== null) spidMap.set(oldId, nextId)
				setAttr(cNvPr, 'id', String(nextId++))
			}

			// Insert into the host tree (this reparents it out of any holder).
			spTree.insertBefore(imported, anchor)

			// Carry the shape's slide-scoped build animation (opt-in): append its effect
			// click-group(s) + <p:bldP> into the destination timing, remapped to the new id.
			if (options.carryAnimation) {
				const sourceRoot = source.part.dom.documentElement
				const targetRoot = target.part.dom.documentElement
				if (sourceRoot && targetRoot) carryShapeAnimations(sourceRoot, targetRoot, spidMap)
			}

			const shape = wrapShapeElement(imported, target)
			if (!shape)
				throw new UnsupportedFeatureError(
					'shape/element-unsupported',
					`importShape: unsupported shape element <${imported.localName}>`
				)
			if (options.left != null) shape.left = options.left
			if (options.top != null) shape.top = options.top
			if (options.width != null) shape.width = options.width
			if (options.height != null) shape.height = options.height
			result.push(shape)
		}

		target.part.markDirty()
		return result
	}

	/**
	 * Open an import out of `source`: this deck as the destination, paired with the
	 * copy registry for that package (created on first use). The registry is held on
	 * the class rather than in the context because it must outlive any one call —
	 * that is what makes a second import from the same source reuse the layout,
	 * master, theme, and media it already copied instead of duplicating them.
	 */
	#importContext(source: OpcPackage): ImportContext {
		let registry = this.#importRegistry.get(source)
		if (!registry) {
			registry = new Map()
			this.#importRegistry.set(source, registry)
		}
		return { dest: this, source, registry }
	}

	/**
	 * Wire a new slide part into `p:sldIdLst` (rel + `p:sldId`) at zero-based
	 * position `at` and return it. `p:sldIdLst` order *is* deck order, so the
	 * insertion point is the only bookkeeping needed. An `at` that is omitted,
	 * negative, or `>=` the current slide count appends (the prior behaviour).
	 */
	#insertSlidePart(newPart: Part, at?: number): Slide {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)

		// A slide part belongs to `p:sldIdLst` exactly once. Two `p:sldId` entries
		// naming one part is a package PowerPoint refuses to open (0x80070570) and
		// the read model cannot see: `slides` counts the list, so the deck reports
		// the slide it does not have and the failure surfaces only when someone
		// opens the file. Every caller here hands over a part it just materialized,
		// so this cannot fire on caller input — only on the import machinery
		// handing back a part it should have copied (issue #18). Refusing at the
		// call is the difference between a thrown error and a bad deliverable.
		for (const rel of presRels) {
			if (rel.type !== SLIDE_REL || rel.targetMode === 'External') continue
			if (presRels.resolveTarget(rel.id) !== newPart.partName) continue
			throw new InternalError(
				'slide/part-already-in-deck',
				`Slide part ${newPart.partName} is already in this deck's slide list; a second p:sldId entry for it would make the package unopenable`
			)
		}

		const relId = presRels.add(SLIDE_REL, relativePartName(presPart.partName, newPart.partName)).id

		const root = presPart.dom.documentElement
		if (!root)
			throw new PackageReadError(
				'package/part-has-no-root',
				'presentation.xml has no document element to append a slide to'
			)
		// A template with zero slides omits p:sldIdLst entirely; create it in
		// CT_Presentation document order (after the *IdLst children, before p:sldSz).
		const sldIdLst = getOrAddChild(root, 'p:sldIdLst', PRESENTATION_SLD_ID_LST_SUCCESSORS)
		const existing = getElements(sldIdLst, 'p:sldId')
		const newSlideId = this.#nextSlideId(existing)
		const sldId = createElement(presPart.dom, 'p:sldId')
		setAttr(sldId, 'id', String(newSlideId))
		setAttr(sldId, 'r:id', relId)

		const inRange = at !== undefined && at >= 0 && at < existing.length
		const newIndex = inRange ? at : existing.length
		const before = inRange ? existing[at] : null
		if (before) sldIdLst.insertBefore(sldId, before)
		else sldIdLst.appendChild(sldId)
		presPart.markDirty()

		return new Slide(this, newPart, newSlideId, newIndex)
	}

	/** A slide id one past the highest existing, but at least ST_SlideId's minimum. */
	#nextSlideId(sldIds: ReturnType<typeof getElements>): number {
		let max = MIN_SLIDE_ID - 1
		for (const sldId of sldIds) {
			const id = intValue(attr(sldId, 'id'))
			if (id !== null && id > max) max = id
		}
		return max + 1
	}

	/** Re-emit the package; untouched parts stay byte-identical (see `OpcPackage.save`). */
	async save(): Promise<Uint8Array> {
		return this.opc.save()
	}
}
