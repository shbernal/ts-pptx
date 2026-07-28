/**
 * Read-model entry point: `Presentation` wraps an `OpcPackage` and exposes a
 * navigable, typed view of the deck (slides → shapes → text), backed by the
 * live DOM so the same nodes can later be mutated.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { relativePartName, relsPartNameFor } from '../opc/partnames.js'
import {
	OOXML_NS,
	attr,
	createElement,
	firstChild,
	getElements,
	getOrAddChild,
	insertInOrder,
	intValue,
	ownerDocumentOf,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import {
	EMBEDDED_FONT_SLOTS,
	FONT_DATA_CONTENT_TYPE,
	FONT_DATA_EXTENSION,
	FONT_REL_TYPE,
	type EmbeddedFont,
	type EmbeddedFontSlot,
} from '../../embedded-fonts.js'
import { flattenShape, flattenSlide, remapLiteralColors, restyleSlide, type FlattenContext } from '../oxml/theme.js'
import { resolveSlideThemeParts } from './theme-context.js'
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
	ImportedSlideMaster,
	LayoutHandle,
	SlideSize,
	SlideSource,
} from './presentation-types.js'
import { carriedDecorations, collectElements, cSldName, firstShapeChild, nthShapeChild } from './slide-dom.js'
import { computeRescale, rescaleSpTree, type RescaleTransform } from './rescale.js'
import { carryTableStyles, copySourceTableStyles } from './table-styles.js'
import { promoteMasters } from './master-registry.js'
import { copyPart, type ImportContext } from './part-copy.js'

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const NOTES_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
// Microsoft 2007 `media` rel: paired with the ECMA audio/video rel (same Target),
// referenced by the slide body's <p14:media r:embed>.
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'

const SLIDE_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const SLIDE_LAYOUT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const NOTES_SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const NOTES_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml'
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
 * Content types that are shared deck chrome: reachable through the
 * presentation → master → layout → theme graph, not owned by any one slide.
 * {@link Presentation.removeSlide} never prunes these as a removed slide's
 * orphan, even while momentarily unreferenced.
 */
const SHARED_CHROME_CONTENT_TYPES = new Set([
	SLIDE_MASTER_CONTENT_TYPE,
	SLIDE_LAYOUT_CONTENT_TYPE,
	THEME_CONTENT_TYPE,
	'application/vnd.openxmlformats-officedocument.themeOverride+xml',
	NOTES_MASTER_CONTENT_TYPE,
	'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
	PRESENTATION_MAIN_CONTENT_TYPE,
])

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

/**
 * `p:embeddedFontLst`'s document-order successors in `CT_Presentation` (index 7,
 * after `smartTags`): everything that may legally follow it, so a created list
 * lands in the right slot when the deck has none yet.
 */
const PRESENTATION_EMBEDDED_FONT_LST_SUCCESSORS = [
	'p:custShowLst',
	'p:photoAlbum',
	'p:custDataLst',
	'p:kinsoku',
	'p:defaultTextStyle',
	'p:modifyVerifier',
	'p:extLst',
]

/**
 * A face slot's document-order successors in `CT_EmbeddedFontListEntry`
 * (`font`, `regular`, `bold`, `italic`, `boldItalic`), so a newly-inserted face
 * keeps the schema's child order regardless of which slots already exist.
 */
const EMBEDDED_FONT_FACE_SUCCESSORS: Record<EmbeddedFontSlot, string[]> = {
	regular: ['p:bold', 'p:italic', 'p:boldItalic'],
	bold: ['p:italic', 'p:boldItalic'],
	italic: ['p:boldItalic'],
	boldItalic: [],
}

/**
 * One typeface's faces normalized for the embedded-font merge core (`#mergeEmbeddedFontEntries`):
 * the `p:font` identity attributes plus, per face slot, a thunk that creates the
 * binary font part on demand and returns its partname. The thunk runs only for a
 * face actually being added (after the typeface+slot de-dupe), so no orphan part is
 * created for a face the deck already embeds. Lets the import-side (copy a part out
 * of a source package) and append-side (write raw generator bytes) callers share one
 * merge core while differing only in how the binary part is produced.
 */
interface IncomingEmbeddedFont {
	typeface: string
	/** `p:font` identity attrs other than `typeface` (panose/pitchFamily/charset), in document order. */
	identity: Array<{ name: string; value: string }>
	faces: Array<{ slot: EmbeddedFontSlot; createPart: () => string }>
}

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
			throw new Error(`Expected exactly one officeDocument relationship, found ${officeDocument.length}`)
		}
		const partName = packageRels.resolveTarget(officeDocumentRel.id)
		const part = this.opc.part(partName)
		if (!part) throw new Error(`officeDocument relationship targets a missing part: ${partName}`)
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
			if (!part) throw new Error(`Slide relationship ${relId} targets a missing part: ${partName}`)
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
	 * bytes verbatim and shares the source's relationship targets (layout, images,
	 * …) by copying its `.rels`; a new presentation→slide relationship and a
	 * `p:sldId` entry are wired up. Marks the presentation part dirty.
	 *
	 * Note: relationships are copied as-is, so a source slide that owns a
	 * one-to-one part (e.g. a notes slide) would end up shared with the clone.
	 */
	cloneSlide(index: number, options: { at?: number } = {}): Slide {
		const source = this.slides[index]
		if (!source) throw new Error(`No slide at index ${index} to clone`)
		const opc = this.opc
		const sourcePart = source.part

		// 1. Copy the slide part bytes verbatim into a fresh slide partname.
		const newPartName = opc.reservePartNameLike(sourcePart.partName)
		const newPart = opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)

		// 2. Copy the slide's relationships (targets resolve identically — same dir).
		const sourceRels = opc.part(relsPartNameFor(sourcePart.partName))
		if (sourceRels) opc.addPart(relsPartNameFor(newPartName), sourceRels.contentType, sourceRels.bytes)

		// 3. Wire the new slide into the presentation (rel + p:sldId entry) at `at`.
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
		if (!slide) throw new Error(`No slide at index ${index} to remove`)
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
		for (const target of formerTargets) this.#pruneIfOrphan(target)

		return partName
	}

	/**
	 * Remove `partName` if it is neither shared chrome nor still referenced by any
	 * remaining part, then recurse into the parts it referenced. The pruning a
	 * removed slide triggers (notes/media/charts the slide alone used).
	 */
	#pruneIfOrphan(partName: string): void {
		const part = this.opc.part(partName)
		if (!part || SHARED_CHROME_CONTENT_TYPES.has(part.contentType)) return
		if (this.#isReferenced(partName)) return
		const rels = this.opc.relationshipsFor(partName)
		const childTargets = [...rels]
			.filter((rel) => rel.targetMode !== 'External')
			.map((rel) => rels.resolveTarget(rel.id))
		this.opc.removePart(relsPartNameFor(partName))
		this.opc.removePart(partName)
		for (const child of childTargets) this.#pruneIfOrphan(child)
	}

	/** Whether any remaining part (or the package root) resolves an internal relationship to `partName`. */
	#isReferenced(partName: string): boolean {
		for (const owner of [...this.opc.parts.keys(), '/']) {
			if (owner.endsWith('.rels')) continue
			const rels = this.opc.relationshipsFor(owner)
			for (const rel of rels) {
				if (rel.targetMode === 'External') continue
				if (rels.resolveTarget(rel.id) === partName) return true
			}
		}
		return false
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
	 * imports from the same source deck are copied once and reused.
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
		if (!sourceSlide) throw new Error(`No slide at index ${index} to import`)

		// 1. Pre-flight: slide sizes must match unless the caller opts into a rescale.
		const target = this.slideSize
		const incoming = source.slideSize
		const sizesDiffer =
			!target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu
		if (sizesDiffer && !options.rescale) {
			const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
			throw new Error(
				`importSlide requires equal slide sizes (pass { rescale: 'fit' | 'stretch' } to rescale); target is ${fmt(target)}, source is ${fmt(incoming)}`
			)
		}
		if (sizesDiffer && options.rescale && (!target || !incoming)) {
			throw new Error('importSlide rescale requires both decks to declare a slide size (p:sldSz)')
		}

		// 2. Copy the slide and its dependencies. 'preserve' flattens the theme into
		//    the slide and attaches it to this deck's master; 'restyle' attaches it
		//    to this deck's master with theme refs left symbolic (re-brand); 'copy'
		//    brings the source theme subgraph across wholesale.
		const newPartName =
			options.theme === 'preserve'
				? this.#importSlidePreserve(source, sourceSlide, options.carryMasterGraphics === true)
				: options.theme === 'restyle'
					? this.#importSlideRestyle(
							source,
							sourceSlide,
							options.carryMasterGraphics === true,
							options.remapLiterals === true
						)
					: copyPart(this.#importContext(source.opc), sourceSlide.partName)
		const newPart = this.opc.part(newPartName)
		if (!newPart) throw new Error(`Imported slide part went missing: ${newPartName}`)

		// 2b. Rescale the imported geometry to this deck's canvas when sizes differ.
		if (sizesDiffer && options.rescale && target && incoming) {
			this.#rescaleImportedGeometry(
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
		//    drops the notesSlide rel (both copyPart and #importSlideRebind do); this
		//    re-adds it wired to the new slide and merged onto a single notesMaster.
		if (options.importNotes) this.#carryNotes(source, sourceSlide.partName, newPartName)

		// 5. Optionally carry the source deck's embedded fonts (presentation-level, so
		//    a separate traversal from the slide-part copy chain above).
		if (options.embedFonts) this.#carryEmbeddedFonts(source)

		return slide
	}

	/**
	 * Carry the source slide's speaker notes onto the just-imported slide (the
	 * `importNotes` option). The slide copy itself dropped the `notesSlide` rel, so
	 * this copies the source `notesSlide` part into a fresh partname, wires a
	 * `slide → notesSlide` rel on the new slide, and rebuilds the copied notesSlide's
	 * own relationships:
	 *
	 * - its `slide` back-rel is repointed at the new slide (`newSlidePartName`) — the
	 *   source slide is *not* copied (that would be circular);
	 * - its `notesMaster` rel is resolved through {@link #ensureNotesMaster}, which
	 *   reuses this deck's notesMaster when it has one and copies the source's only
	 *   when it has none (a deck may have at most one notesMaster);
	 * - any other internal target (media, etc.) is copied via {@link copyPart}.
	 *
	 * No-op when the source slide has no notes. Content-type registration for the
	 * copied parts is handled by `addPart`/{@link copyPart}.
	 */
	#carryNotes(source: Presentation, sourceSlidePartName: string, newSlidePartName: string): void {
		const sourceSlideRels = source.opc.relationshipsFor(sourceSlidePartName)
		const notesRel = sourceSlideRels.byType(NOTES_SLIDE_REL)[0]
		if (!notesRel) return // slide has no speaker notes
		const sourceNotesPartName = sourceSlideRels.resolveTarget(notesRel.id)
		const sourceNotesPart = source.opc.part(sourceNotesPartName)
		if (!sourceNotesPart) return

		// Copy the notesSlide bytes into a fresh partname, then wire slide → notesSlide.
		const newNotesPartName = this.opc.reservePartNameLike(sourceNotesPartName)
		this.opc.addPart(newNotesPartName, sourceNotesPart.contentType, sourceNotesPart.bytes)
		this.opc
			.relationshipsFor(newSlidePartName)
			.add(NOTES_SLIDE_REL, relativePartName(newSlidePartName, newNotesPartName))

		// Rebuild the copied notesSlide's relationships. Preserve each source rel id so
		// the notesSlide body's r:id references stay valid; only the targets are rewritten.
		const ctx = this.#importContext(source.opc)
		const notesSourceRels = source.opc.relationshipsFor(sourceNotesPartName)
		const notesTargetRels = this.opc.relationshipsFor(newNotesPartName)
		for (const rel of notesSourceRels) {
			if (rel.type === SLIDE_REL) {
				// Back-reference to the annotated slide → repoint at the new slide (don't copy it).
				notesTargetRels.addWithId(rel.id, SLIDE_REL, relativePartName(newNotesPartName, newSlidePartName))
				continue
			}
			if (rel.type === NOTES_MASTER_REL) {
				const notesMaster = this.#ensureNotesMaster(ctx, notesSourceRels.resolveTarget(rel.id))
				notesTargetRels.addWithId(rel.id, NOTES_MASTER_REL, relativePartName(newNotesPartName, notesMaster))
				continue
			}
			if (rel.targetMode === 'External') {
				notesTargetRels.addWithId(rel.id, rel.type, rel.target, 'External')
				continue
			}
			const newTarget = copyPart(ctx, notesSourceRels.resolveTarget(rel.id))
			notesTargetRels.addWithId(rel.id, rel.type, relativePartName(newNotesPartName, newTarget))
		}
	}

	/**
	 * Resolve the notesMaster an imported `notesSlide` should bind to, honouring the
	 * single-notesMaster-per-presentation rule (`p:notesMasterIdLst` holds 0..1
	 * `p:notesMasterId`). If this deck already has a notesMaster it is reused and the
	 * source's is *not* copied (the destination's notes styling wins); otherwise the
	 * source notesMaster (and, via {@link copyPart}, its theme) is copied and
	 * registered in `presentation.xml`. Returns the destination notesMaster partname.
	 */
	#ensureNotesMaster(ctx: ImportContext, sourceNotesMasterPartName: string): string {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const existing = presRels.byType(NOTES_MASTER_REL)[0]
		if (existing) return presRels.resolveTarget(existing.id)

		// No notesMaster yet: copy the source's (pulls its theme) and register it.
		return this.#registerNotesMaster(copyPart(ctx, sourceNotesMasterPartName))
	}

	/**
	 * Wire an already-added notesMaster part into `presentation.xml`: a `notesMaster`
	 * relationship plus the single `p:notesMasterId` entry that `CT_NotesMasterIdList`
	 * allows. Returns the partname, so callers can use it as a rel target.
	 *
	 * Split out of {@link #ensureNotesMaster} because the two ways a notesMaster arrives
	 * — copied from another `Presentation`, or authored by a generator and injected by
	 * {@link appendSlides} — differ only in how the *part* is created, not in how it is
	 * registered.
	 */
	#registerNotesMaster(notesMasterPartName: string): string {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const relId = presRels.add(NOTES_MASTER_REL, relativePartName(presPart.partName, notesMasterPartName)).id

		const root = presPart.dom.documentElement
		if (!root) throw new Error('presentation.xml has no document element to register a notes master in')
		// `p:notesMasterIdLst` follows `p:sldMasterIdLst` in CT_Presentation order.
		const lst = getOrAddChild(root, 'p:notesMasterIdLst', [
			'p:handoutMasterIdLst',
			'p:sldIdLst',
			'p:sldSz',
			'p:notesSz',
			'p:embeddedFontLst',
			'p:custShowLst',
			'p:photoAlbum',
			'p:custDataLst',
			'p:kinsoku',
			'p:defaultTextStyle',
			'p:modifyVerifier',
			'p:extLst',
		])
		// CT_NotesMasterIdList holds a single p:notesMasterId; replace any stray entry.
		removeChildrenByQName(lst, ['p:notesMasterId'])
		const entry = createElement(presPart.dom, 'p:notesMasterId')
		setAttr(entry, 'r:id', relId)
		lst.appendChild(entry)
		presPart.markDirty()
		return notesMasterPartName
	}

	/**
	 * Resolve the notesMaster an *appended* slide's notes should bind to. Same
	 * single-notesMaster rule as {@link #ensureNotesMaster}: this deck's own wins when it
	 * has one, so the destination's notes styling is preserved and `master.xml` is
	 * discarded. Otherwise the generator's notes master is installed, together with the
	 * theme its `.rels` requires (the normal write path emits that as `theme2.xml`).
	 */
	#ensureNotesMasterFromXml(master: { xml: string; themeXml: string }): string {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const existing = presRels.byType(NOTES_MASTER_REL)[0]
		if (existing) return presRels.resolveTarget(existing.id)

		const masterPartName = this.opc.reservePartNameLike('/ppt/notesMasters/notesMaster1.xml')
		this.opc.addPart(masterPartName, NOTES_MASTER_CONTENT_TYPE, textEncoder.encode(master.xml))

		// A notesMaster's .rels must resolve a theme; reserve alongside any theme the
		// destination already owns rather than assuming theme2.xml is free.
		const themePartName = this.opc.reservePartNameLike('/ppt/theme/theme1.xml')
		this.opc.addPart(themePartName, THEME_CONTENT_TYPE, textEncoder.encode(master.themeXml))
		this.opc.relationshipsFor(masterPartName).add(THEME_REL, relativePartName(masterPartName, themePartName))

		return this.#registerNotesMaster(masterPartName)
	}

	/**
	 * Copy `source`'s embedded fonts into this deck and merge them into our
	 * `p:embeddedFontLst`. Font binaries come across via {@link copyPart} (so the
	 * per-source registry dedupes faces shared across repeated imports); entries are
	 * merged by `typeface` + face slot, so a face this deck already embeds is reused
	 * rather than duplicated. No-op when the source embeds no fonts. See
	 * {@link ImportSlideOptions.embedFonts}.
	 */
	#carryEmbeddedFonts(source: Presentation): void {
		const sourceRoot = source.presentationPart.dom.documentElement
		const sourceLst = sourceRoot && firstChild(sourceRoot, 'p:embeddedFontLst')
		const sourceEntries = sourceLst ? getElements(sourceLst, 'p:embeddedFont') : []
		if (sourceEntries.length === 0) return

		const ctx = this.#importContext(source.opc)
		const sourcePresRels = source.opc.relationshipsFor(source.presentationPart.partName)
		const incoming: IncomingEmbeddedFont[] = []
		for (const srcEntry of sourceEntries) {
			const srcFont = firstChild(srcEntry, 'p:font')
			const typeface = srcFont ? attr(srcFont, 'typeface') : null
			if (!srcFont || !typeface) continue

			// Copy the source p:font identity attributes (panose/pitchFamily/charset).
			const identity: IncomingEmbeddedFont['identity'] = []
			for (const name of ['panose', 'pitchFamily', 'charset']) {
				const value = attr(srcFont, name)
				if (value !== null) identity.push({ name, value })
			}

			const faces: IncomingEmbeddedFont['faces'] = []
			for (const slot of EMBEDDED_FONT_SLOTS) {
				const srcFace = firstChild(srcEntry, `p:${slot}`)
				const srcRid = srcFace && attr(srcFace, 'r:id')
				if (!srcFace || !srcRid) continue
				// Binary comes across via copyPart, so the per-source registry dedupes faces
				// shared across repeated imports; the thunk runs only when the face is added.
				faces.push({ slot, createPart: () => copyPart(ctx, sourcePresRels.resolveTarget(srcRid)) })
			}
			incoming.push({ typeface, identity, faces })
		}
		this.#mergeEmbeddedFontEntries(incoming)
	}

	/**
	 * Carry a generator's presentation-level embedded fonts ({@link ExtractedSlides.embeddedFonts},
	 * from `pptx.embedFont`) into this deck during {@link appendSlides}. Each face's raw bytes are
	 * written as a fresh `/ppt/fonts/fontN.fntdata` part; merge/de-dupe by typeface + slot is shared
	 * with {@link #carryEmbeddedFonts} via {@link #mergeEmbeddedFontEntries}, so appending the same
	 * generator twice (or onto a deck that already embeds the face) carries each face once.
	 */
	#carryGeneratedEmbeddedFonts(fonts: EmbeddedFont[]): void {
		const incoming: IncomingEmbeddedFont[] = []
		for (const font of fonts) {
			if (!font.typeface) continue
			const identity: IncomingEmbeddedFont['identity'] = []
			if (font.panose !== undefined) identity.push({ name: 'panose', value: font.panose })
			if (font.pitchFamily !== undefined) identity.push({ name: 'pitchFamily', value: String(font.pitchFamily) })
			if (font.charset !== undefined) identity.push({ name: 'charset', value: String(font.charset) })

			const faces: IncomingEmbeddedFont['faces'] = []
			for (const slot of EMBEDDED_FONT_SLOTS) {
				const face = font.faces.find((f) => f.slot === slot)
				if (!face?.bytes) continue
				const bytes = face.bytes
				faces.push({
					slot,
					createPart: () => {
						const partName = this.opc.reservePartNameLike('/ppt/fonts/font1.fntdata')
						this.opc.addPart(partName, FONT_DATA_CONTENT_TYPE, bytes)
						return partName
					},
				})
			}
			if (faces.length > 0) incoming.push({ typeface: font.typeface, identity, faces })
		}
		this.#mergeEmbeddedFontEntries(incoming)
	}

	/**
	 * Merge normalized {@link IncomingEmbeddedFont} entries into this deck's
	 * `p:embeddedFontLst` — the shared core of {@link #carryEmbeddedFonts} (import-side)
	 * and {@link #carryGeneratedEmbeddedFonts} (append-side). Entries merge by `typeface`,
	 * faces de-dupe by slot (a face this deck already embeds is left as is). For each newly
	 * added face the `fntdata` Default is ensured, the binary part is created via the face's
	 * `createPart` thunk, a `font` rel is added to presentation.xml, and the `p:<slot>` element
	 * is inserted in schema child order. The list is created at CT_Presentation index 7 when
	 * the deck has none yet. No-op for empty input.
	 */
	#mergeEmbeddedFontEntries(entries: IncomingEmbeddedFont[]): void {
		if (entries.length === 0) return

		const presPart = this.presentationPart
		const presRoot = presPart.dom.documentElement
		if (!presRoot) throw new Error('presentation.xml has no document element to carry embedded fonts into')
		const presRels = this.opc.relationshipsFor(presPart.partName)

		const targetLst = getOrAddChild(presRoot, 'p:embeddedFontLst', PRESENTATION_EMBEDDED_FONT_LST_SUCCESSORS)
		const targetByTypeface = new Map<string, Element>()
		for (const entry of getElements(targetLst, 'p:embeddedFont')) {
			const font = firstChild(entry, 'p:font')
			const typeface = font && attr(font, 'typeface')
			if (typeface) targetByTypeface.set(typeface, entry)
		}

		let copiedAny = false
		for (const incoming of entries) {
			// Find or create the target entry for this typeface, carrying its
			// p:font identity attributes (typeface + optional panose/pitchFamily/charset).
			let targetEntry = targetByTypeface.get(incoming.typeface)
			if (!targetEntry) {
				targetEntry = createElement(presPart.dom, 'p:embeddedFont')
				const targetFont = createElement(presPart.dom, 'p:font')
				setAttr(targetFont, 'typeface', incoming.typeface)
				for (const { name, value } of incoming.identity) setAttr(targetFont, name, value)
				targetEntry.appendChild(targetFont)
				targetLst.appendChild(targetEntry)
				targetByTypeface.set(incoming.typeface, targetEntry)
			}

			for (const face of incoming.faces) {
				if (firstChild(targetEntry, `p:${face.slot}`)) continue // de-dupe: face already present

				// Ensure the fntdata Default exists *before* creating the part, so addPart
				// resolves the content type via the Default (no per-part Override).
				this.opc.contentTypes.ensureDefault(FONT_DATA_EXTENSION, FONT_DATA_CONTENT_TYPE)
				const newFontPart = face.createPart()
				const relId = presRels.add(FONT_REL_TYPE, relativePartName(presPart.partName, newFontPart)).id

				const targetFace = createElement(presPart.dom, `p:${face.slot}`)
				setAttr(targetFace, 'r:id', relId)
				insertInOrder(targetEntry, targetFace, EMBEDDED_FONT_FACE_SUCCESSORS[face.slot])
				copiedAny = true
			}
		}

		if (copiedAny) presPart.markDirty()
	}

	/**
	 * Rescale an imported slide's geometry onto this deck's canvas (the `rescale`
	 * option of {@link importSlide}). Rewrites every top-level shape/group/
	 * graphicFrame transform and table grid on the slide; in `copy` mode also
	 * rescales the imported layout and master shape trees (resolved via the
	 * slide → layout → master rel chain) so inherited placeholder/background geometry
	 * stays aligned. `preserve`/`restyle` rebind to this deck's own master/layout —
	 * already the destination size — so only the slide is touched. Geometry only:
	 * font sizes and line widths are left as authored.
	 */
	#rescaleImportedGeometry(
		slidePartName: string,
		theme: ImportSlideOptions['theme'],
		source: SlideSize,
		target: SlideSize,
		mode: 'fit' | 'stretch'
	): void {
		const transform = computeRescale(source, target, mode)
		this.#rescalePartGeometry(slidePartName, transform)
		if (theme === undefined || theme === 'copy') {
			const layout = this.#resolveSingleRel(this.opc, slidePartName, SLIDE_LAYOUT_REL)
			const master = layout ? this.#resolveSingleRel(this.opc, layout, SLIDE_MASTER_REL) : null
			if (layout) this.#rescalePartGeometry(layout, transform)
			if (master) this.#rescalePartGeometry(master, transform)
		}
	}

	/**
	 * Rescale one part's `p:spTree` geometry in place. Idempotent per part
	 * (#rescaledParts), so a layout/master shared across repeated imports from one
	 * source is rescaled exactly once.
	 */
	#rescalePartGeometry(partName: string, transform: RescaleTransform): void {
		if (this.#rescaledParts.has(partName)) return
		this.#rescaledParts.add(partName)
		const part = this.opc.part(partName)
		const root = part?.dom.documentElement
		const cSld = root && firstChild(root, 'p:cSld')
		const spTree = cSld && firstChild(cSld, 'p:spTree')
		if (!part || !spTree) return
		rescaleSpTree(spTree, transform)
		part.markDirty()
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
				throw new Error(
					`importSlideMasters requires equal slide sizes (pass { requireEqualSize: false } to override); target is ${fmt(target)}, source is ${fmt(incoming)}`
				)
			}
		}

		const pickMaster = options.masters ?? (() => true)
		const pickLayout = options.layouts ?? (() => true)

		const ctx = this.#importContext(source.opc)
		const imported: ImportedSlideMaster[] = []
		source.#slideMasterPartNames().forEach((masterPartName, masterIndex) => {
			if (!pickMaster(cSldName(source.opc.part(masterPartName)), masterIndex)) return

			// Copy the (lean) master first: copyPart registers it in p:sldMasterIdLst
			// and clears its layout list, then each copied layout re-links itself in.
			const newMasterPartName = copyPart(ctx, masterPartName)

			const layoutPartNames: string[] = []
			source.#layoutPartNamesOf(masterPartName).forEach((layoutPartName, layoutIndex) => {
				if (!pickLayout(cSldName(source.opc.part(layoutPartName)), layoutIndex)) return
				layoutPartNames.push(copyPart(ctx, layoutPartName))
			})

			imported.push({ partName: newMasterPartName, layoutPartNames })
		})

		// Optionally carry the source deck's presentation-level styling parts. Both are
		// separate traversals from the master/layout copy chain above, and both are
		// whole-deck: neither part records which font/style belongs to which master.
		if (options.embedFonts) this.#carryEmbeddedFonts(source)
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
		return this.#slideMasterPartNames()
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
		this.#slideMasterPartNames().forEach((masterPartName, masterIndex) => {
			this.#layoutPartNamesOf(masterPartName).forEach((layoutPartName, layoutIndex) => {
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
				throw new Error(
					`appendSlides: layout name ${JSON.stringify(options.layout)} is ambiguous (${matches.length} layouts share it); pass a LayoutHandle from layouts() instead`
				)
			}
			const [only] = matches
			if (!only) {
				const names = gallery.map((l) => JSON.stringify(l.name)).join(', ')
				throw new Error(
					`appendSlides: no layout named ${JSON.stringify(options.layout)}; available: ${names || '(none)'}`
				)
			}
			target = only
		} else {
			const handle = options.layout
			if (!gallery.some((l) => l.partName === handle.partName)) {
				throw new Error(`appendSlides: layout ${handle.partName} does not belong to this presentation`)
			}
			target = handle
		}

		// 2. Author + extract; enforce equal slide size (no geometry rescale in v1).
		const extracted = await source.extractSlides({ onMediaError: options.onMediaError })
		const size = this.slideSize
		if (!size || size.widthEmu !== extracted.widthEmu || size.heightEmu !== extracted.heightEmu) {
			const fmt = (w: number, h: number): string => `${w}×${h} EMU`
			throw new Error(
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
				const notesMasterPartName = extracted.notesMaster ? this.#ensureNotesMasterFromXml(extracted.notesMaster) : null
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
					throw new Error(
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
		this.#carryGeneratedEmbeddedFonts(extracted.embeddedFonts || [])

		return added
	}

	/** Source-side helper: master partnames in `p:sldMasterIdLst` order. */
	#slideMasterPartNames(): string[] {
		const root = this.presentationPart.dom.documentElement
		const lst = root && firstChild(root, 'p:sldMasterIdLst')
		if (!lst) return []
		const rels = this.opc.relationshipsFor(this.presentationPart.partName)
		const out: string[] = []
		for (const entry of getElements(lst, 'p:sldMasterId')) {
			const relId = attr(entry, 'r:id')
			if (relId) out.push(rels.resolveTarget(relId))
		}
		return out
	}

	/** Source-side helper: a master's layout partnames in `p:sldLayoutIdLst` order. */
	#layoutPartNamesOf(masterPartName: string): string[] {
		const root = this.opc.part(masterPartName)?.dom.documentElement
		const lst = root && firstChild(root, 'p:sldLayoutIdLst')
		if (!lst) return []
		const rels = this.opc.relationshipsFor(masterPartName)
		const out: string[] = []
		for (const entry of getElements(lst, 'p:sldLayoutId')) {
			const relId = attr(entry, 'r:id')
			if (relId) out.push(rels.resolveTarget(relId))
		}
		return out
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
		if (!shape) throw new Error(`importShape: source slide has no shape at index ${shapeIndex}`)
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
		if (target.presentation !== this) throw new Error('importShape: target slide must belong to this presentation')

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
				throw new Error(
					`importShape requires equal slide sizes (or { rescale }); target is ${fmt(targetSize)}, source is ${fmt(sourceSize)}`
				)
			}
			transform = computeRescale(sourceSize, targetSize, options.rescale === 'stretch' ? 'stretch' : 'fit')
		}

		// Resolve + validate every index up front so a bad batch throws before mutating.
		const sourceShapes = source.shapes
		const sourceElements = shapeIndices.map((i) => {
			const shape = sourceShapes[i]
			if (!shape) throw new Error(`No shape at index ${i} on the source slide (it has ${sourceShapes.length})`)
			return shape.element_
		})

		const spTree = target.shapeTree()
		if (!spTree) throw new Error(`importShape: target slide ${target.partName} has no shape tree`)
		const targetDoc = spTree.ownerDocument
		if (!targetDoc) throw new Error('importShape: target slide DOM has no owner document')

		const theme = options.theme ?? 'preserve'
		const sourceOpc = source.presentation.opc
		const sourceRels = sourceOpc.relationshipsFor(source.partName)
		const targetRels = this.opc.relationshipsFor(target.partName)
		// One rel-id map across the batch so shapes sharing a source image share a rel.
		const relIdMap = new Map<string, string>()
		// preserve: build the source theme context once; copy/restyle need none.
		const ctx = theme === 'preserve' ? this.#sourceFlattenContext(sourceOpc, source.partName) : null
		const importCtx = this.#importContext(sourceOpc)

		// Anchor for z-order: the existing shape currently at `at` (insert before it,
		// preserving batch order), else append before any trailing p:extLst.
		const extLst = firstChild(spTree, 'p:extLst')
		const anchor = options.at == null ? extLst : (nthShapeChild(spTree, options.at) ?? extLst)

		const result: AnyShape[] = []
		for (const shapeEl of sourceElements) {
			const imported = targetDoc.importNode(shapeEl, true)

			// Drag media/charts/embeddings across and rewrite refs to fresh host rels.
			this.#rewriteCarriedRels(imported, importCtx, sourceRels, target.partName, targetRels, relIdMap)

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
			if (!shape) throw new Error(`importShape: unsupported shape element <${imported.localName}>`)
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
	 * Import a slide in `preserve` mode: rebind it to this deck's master/layout
	 * (see {@link #importSlideRebind}), then flatten its source theme into the slide
	 * XML (scheme colours + style-matrix fills baked to literals). Returns the new
	 * partname.
	 *
	 * The flatten context is gathered from the *source* subgraph, so it can be read
	 * before or after the rebind; the rebind injects any carried decorations before
	 * we flatten, so a single sweep resolves the theme references on the slide's own
	 * content and on the carried decorations together.
	 */
	#importSlidePreserve(source: Presentation, sourceSlide: Slide, carryGraphics: boolean): string {
		const ctx = this.#sourceFlattenContext(source.opc, sourceSlide.partName)
		const { newPartName, slideRoot, newPart } = this.#importSlideRebind(source, sourceSlide, carryGraphics)
		flattenSlide(slideRoot, ctx)
		newPart.markDirty()
		return newPartName
	}

	/**
	 * Import a slide in `restyle` mode: rebind it to this deck's master/layout (see
	 * {@link #importSlideRebind}) and then {@link restyleSlide} it — drop its colour
	 * map override but bake *nothing*, so its symbolic theme references re-resolve
	 * against the destination theme and the slide re-brands. Returns the new
	 * partname.
	 *
	 * The deliberate inverse of `preserve`: no flatten, no inherited-background
	 * bake, no placeholder colour/size/geometry bake — every one of those would pin
	 * the slide to its source look, the opposite of re-branding. Carried
	 * decorations are left symbolic too, so they re-brand along with the slide.
	 *
	 * With `remapLiterals` it additionally force-remaps the slide's source-theme
	 * literal colours back to symbolic scheme colours and copies any referenced
	 * source table style into this deck — the two things plain `restyle` cannot
	 * re-brand (see {@link ImportSlideOptions.remapLiterals}).
	 */
	#importSlideRestyle(
		source: Presentation,
		sourceSlide: Slide,
		carryGraphics: boolean,
		remapLiterals: boolean
	): string {
		const { newPartName, slideRoot, newPart } = this.#importSlideRebind(source, sourceSlide, carryGraphics)
		restyleSlide(slideRoot)
		if (remapLiterals) {
			// The source colour context (slot ↔ RGB ↔ token) the literals are matched against.
			const parts = resolveSlideThemeParts(source.opc, sourceSlide.partName)
			remapLiteralColors(slideRoot, { clrMap: parts.clrMap, clrScheme: parts.clrScheme })
			copySourceTableStyles(this, source.opc, slideRoot)
		}
		newPart.markDirty()
		return newPartName
	}

	/**
	 * The rebind shared by `preserve` and `restyle`: copy the slide bytes into a
	 * fresh part, rebuild its relationships (drop notes, repoint the `slideLayout`
	 * rel at this deck's existing layout, copy every other internal target —
	 * media/charts — and pass externals through), and optionally bake the source
	 * master/layout decorations onto the slide. Returns the new part, its name, and
	 * its live root element for the caller's mode-specific pass (flatten vs restyle).
	 *
	 * This carries *no* theme baking of its own — not even the inherited background.
	 * `preserve` adds that via {@link flattenSlide}'s context; `restyle` must not,
	 * so the background stays symbolic and re-brands.
	 */
	#importSlideRebind(
		source: Presentation,
		sourceSlide: Slide,
		carryGraphics: boolean
	): { newPartName: string; slideRoot: Element; newPart: Part } {
		const destLayout = this.#destinationLayoutPartName()

		// Copy the slide bytes into a fresh partname; we then mutate that copy's DOM
		// (a distinct document, so the source package is never touched).
		const sourcePart = source.opc.part(sourceSlide.partName)
		if (!sourcePart) throw new Error(`importSlide: source package has no part ${sourceSlide.partName}`)
		const newPartName = this.opc.reservePartNameLike(sourceSlide.partName)
		const newPart = this.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
		const slideRoot = newPart.dom.documentElement
		if (!slideRoot) throw new Error(`Imported slide ${newPartName} has no root element`)

		// Rebuild the slide's relationships: drop notes, repoint slideLayout at the
		// destination layout, and copy every other internal target (media/charts).
		const ctx = this.#importContext(source.opc)
		const sourceRels = source.opc.relationshipsFor(sourceSlide.partName)
		const targetRels = this.opc.relationshipsFor(newPartName)
		for (const rel of sourceRels) {
			if (rel.type === NOTES_SLIDE_REL) continue
			if (rel.type === SLIDE_LAYOUT_REL) {
				targetRels.addWithId(rel.id, SLIDE_LAYOUT_REL, relativePartName(newPartName, destLayout))
				continue
			}
			if (rel.targetMode === 'External') {
				targetRels.addWithId(rel.id, rel.type, rel.target, 'External')
				continue
			}
			const newTarget = copyPart(ctx, sourceRels.resolveTarget(rel.id))
			targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTarget))
		}

		// Optionally bake the source master/layout decorations (logos, accent shapes)
		// onto the slide behind its own content. Done after the slide's own rels are
		// in place (so carried media get fresh, non-colliding ids) but before the
		// caller's flatten/restyle pass acts on the carried shapes.
		if (carryGraphics) this.#carryMasterGraphics(ctx, slideRoot, newPartName, sourceSlide.partName)

		return { newPartName, slideRoot, newPart }
	}

	/**
	 * Bake the source `slideLayout`/`slideMaster` shape-tree decorations onto the
	 * imported slide (the `carryMasterGraphics` path). Every shape on those trees
	 * *except* placeholders is deep-copied into the slide's `p:spTree` ahead of its
	 * own content — master decorations first, then layout, then the slide's shapes —
	 * so document (z-)order keeps the master furthest back. Each decoration's media
	 * and other relationship targets are copied into this package and its
	 * `r:embed`/`r:id`/… references rewritten to fresh slide-local ids. The injected
	 * shapes are left for the caller's {@link flattenSlide} pass to resolve any
	 * theme references they carry.
	 */
	#carryMasterGraphics(ctx: ImportContext, slideRoot: Element, newPartName: string, slidePartName: string): void {
		const sourceOpc = ctx.source
		const layoutPartName = this.#resolveSingleRel(sourceOpc, slidePartName, SLIDE_LAYOUT_REL)
		const masterPartName = layoutPartName ? this.#resolveSingleRel(sourceOpc, layoutPartName, SLIDE_MASTER_REL) : null
		const cSld = firstChild(slideRoot, 'p:cSld')
		const spTree = cSld && firstChild(cSld, 'p:spTree')
		if (!spTree) return

		const doc = ownerDocumentOf(slideRoot)
		const slideRels = this.opc.relationshipsFor(newPartName)
		const relIdMap = new Map<string, string>()
		// Insert ahead of the slide's own first shape so decorations render behind it.
		const anchor = firstShapeChild(spTree)
		// Master behind layout behind the slide (document order == z-order).
		for (const partName of [masterPartName, layoutPartName]) {
			if (!partName) continue
			const decorations = carriedDecorations(sourceOpc.part(partName)?.dom.documentElement ?? null)
			if (decorations.length === 0) continue
			const sourceRels = sourceOpc.relationshipsFor(partName)
			for (const deco of decorations) {
				const imported = doc.importNode(deco, true)
				this.#rewriteCarriedRels(imported, ctx, sourceRels, newPartName, slideRels, relIdMap)
				spTree.insertBefore(imported, anchor)
			}
		}
	}

	/**
	 * Rewrite every relationship reference (`r:embed`, `r:id`, `r:link`, …) inside a
	 * carried decoration so it points at a fresh slide-local relationship, copying
	 * the referenced part into this package on first sight. `relIdMap` (keyed by
	 * source part + source rel id) dedupes references shared within one import call.
	 */
	#rewriteCarriedRels(
		node: Element,
		ctx: ImportContext,
		sourceRels: Relationships,
		newPartName: string,
		slideRels: Relationships,
		relIdMap: Map<string, string>
	): void {
		const elements: Element[] = []
		collectElements(node, elements)
		for (const el of elements) {
			const refs: { local: string; id: string }[] = []
			const attrs = el.attributes
			for (let i = 0; i < attrs.length; i++) {
				const a = attrs.item(i)
				if (!a || a.namespaceURI !== OOXML_NS.r || !a.value) continue
				if (!sourceRels.get(a.value)) continue // an r-namespaced attribute that isn't a relationship id
				refs.push({ local: a.localName ?? a.name, id: a.value })
			}
			for (const { local, id } of refs) {
				setAttr(el, `r:${local}`, this.#carryRel(ctx, sourceRels, id, newPartName, slideRels, relIdMap))
			}
		}
	}

	/** Resolve a carried decoration's source relationship to a fresh slide-local id, copying its internal target. */
	#carryRel(
		ctx: ImportContext,
		sourceRels: Relationships,
		id: string,
		newPartName: string,
		slideRels: Relationships,
		relIdMap: Map<string, string>
	): string {
		const key = `${sourceRels.sourcePartName}|${id}`
		const cached = relIdMap.get(key)
		if (cached) return cached
		const rel = sourceRels.get(id)
		if (!rel) throw new Error(`Relationships of ${sourceRels.sourcePartName}: no relationship with id ${id}`)
		const newId =
			rel.targetMode === 'External'
				? slideRels.add(rel.type, rel.target, 'External').id
				: slideRels.add(rel.type, relativePartName(newPartName, copyPart(ctx, sourceRels.resolveTarget(id)))).id
		relIdMap.set(key, newId)
		return newId
	}

	/**
	 * The partname of the layout this deck's slides should attach to in `preserve`
	 * mode: the first layout of the first slide master. Throws when the deck has no
	 * master/layout to attach to (a deck ts-pptx always provides).
	 */
	#destinationLayoutPartName(): string {
		const presRels = this.opc.relationshipsFor(this.presentationPart.partName)
		const masterRel = presRels.byType(SLIDE_MASTER_REL)[0]
		if (!masterRel) throw new Error('importSlide preserve mode requires a slide master in the destination deck')
		const masterPartName = presRels.resolveTarget(masterRel.id)
		const masterRels = this.opc.relationshipsFor(masterPartName)
		const layoutRel = masterRels.byType(SLIDE_LAYOUT_REL)[0]
		if (!layoutRel) throw new Error('importSlide preserve mode requires a slide layout in the destination deck')
		return masterRels.resolveTarget(layoutRel.id)
	}

	/**
	 * Gather the flatten context for a source slide: walk slide → layout → master →
	 * theme, reading the effective colour map (the slide's `clrMapOvr` override, or
	 * the master `clrMap`), the theme `clrScheme`, and the theme `fmtScheme`.
	 */
	#sourceFlattenContext(sourceOpc: OpcPackage, slidePartName: string): FlattenContext {
		// Reuse the shared slide → layout → master → theme walk (also backing the
		// read-model colour getters), then layer the flatten-only needs on top.
		const parts = resolveSlideThemeParts(sourceOpc, slidePartName)
		const themeElements = parts.themeElements
		return {
			clrMap: parts.clrMap,
			clrScheme: parts.clrScheme,
			fmtScheme: themeElements ? firstChild(themeElements, 'a:fmtScheme') : null,
			inheritedBackground: this.#effectiveBackground(
				sourceOpc,
				parts.slideRoot,
				parts.layoutPartName,
				parts.masterPartName
			),
			layoutRoot: parts.layoutRoot,
			masterRoot: parts.masterRoot,
		}
	}

	/**
	 * The background the slide effectively inherits from its source subgraph: the
	 * layout's `p:bg`, else the master's. Returns `null` when the slide carries its
	 * own `p:bg` (it stays on the slide and is flattened directly) or none exists.
	 */
	#effectiveBackground(
		sourceOpc: OpcPackage,
		slideRoot: Element | null,
		layoutPartName: string | null,
		masterPartName: string | null
	): Element | null {
		if (slideRoot && this.#backgroundOf(slideRoot)) return null
		const layoutRoot = layoutPartName ? (sourceOpc.part(layoutPartName)?.dom.documentElement ?? null) : null
		const masterRoot = masterPartName ? (sourceOpc.part(masterPartName)?.dom.documentElement ?? null) : null
		return (layoutRoot && this.#backgroundOf(layoutRoot)) ?? (masterRoot && this.#backgroundOf(masterRoot)) ?? null
	}

	/** The `p:cSld/p:bg` element of a slide/layout/master root, or `null`. */
	#backgroundOf(root: Element): Element | null {
		const cSld = firstChild(root, 'p:cSld')
		return cSld ? firstChild(cSld, 'p:bg') : null
	}

	/** Resolve the single relationship of `type` owned by `partName`, or `null`. */
	#resolveSingleRel(sourceOpc: OpcPackage, partName: string, type: string): string | null {
		const rels = sourceOpc.relationshipsFor(partName)
		const rel = rels.byType(type)[0]
		return rel ? rels.resolveTarget(rel.id) : null
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
		const relId = presRels.add(SLIDE_REL, relativePartName(presPart.partName, newPart.partName)).id

		const root = presPart.dom.documentElement
		if (!root) throw new Error('presentation.xml has no document element to append a slide to')
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
