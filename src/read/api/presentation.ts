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
 *
 * The four import entry points went the same way, and {@link appendSlides} after them: their
 * contracts stay here as the doc comments a caller reads, and their bodies live in
 * `presentation-imports.ts` and `ops/append-slides.ts`. Those bodies are the one thing in this
 * directory that is not independent of this class — they reach back for {@link importContext},
 * {@link insertSlidePart} and {@link rescaledParts}, which {@link cloneSlide} reads too and so
 * could not travel with them. Those three are `@internal` for that and nothing else.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { relativePartName, relsPartNameFor } from '../opc/partnames.js'
import { attr, createElement, firstChild, getElements, getOrAddChild, numberValue, setAttr } from '../oxml/dom.js'
import { EMBEDDED_FONT_SLOTS } from '../../embedded-fonts.js'
import { PRESENTATION_AFTER_SLD_ID_LST } from '../../ooxml/sequence.js'
import { Slide } from './slide.js'
import { SlideMaster } from './chrome.js'
import type { AnyShape } from './shapes.js'
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
	readExtendedProperties,
	type CoreProperties,
	type CustomProperty,
	type ExtendedProperties,
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
import { cSldName } from '../oxml/slide-dom.js'
import type { ImportContext } from './ops/part-copy.js'
// Deck-mutation operations. They live beside the model rather than on it: each is a whole job
// (prune a part fringe, carry notes, merge embedded fonts, rescale onto a new canvas) that reads
// and writes the package through the deck's public surface, and none of them is something a
// caller navigates *to*.
import { appendSlides as appendSlidesInto } from './ops/append-slides.js'
import { layoutPartNamesOf, slideMasterPartNames } from './ops/part-index.js'
import { duplicateOwnedTargets } from './ops/page-owned.js'
import { pruneIfOrphan } from './ops/prune.js'
import { OFFICE_DOCUMENT_REL, PRESENTATION_MAIN_CONTENT_TYPE, SLIDE_REL } from '../../ooxml/rel-types.js'
import { InternalError, InvalidOptionError, PackageReadError } from '../../errors.js'
import { presentationRels } from './ops/deck-target.js'
import {
	importShape as importShapeInto,
	importShapes as importShapesInto,
	importSlide as importSlideInto,
	importSlideMasters as importSlideMastersInto,
	importSlides as importSlidesInto,
} from './presentation-imports.js'

/** Content type of the main part in a `.potx` template package — flipped to {@link PRESENTATION_MAIN_CONTENT_TYPE} by {@link Presentation.fromTemplate}. */
const PRESENTATION_TEMPLATE_MAIN_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'

/** ST_SlideId minimum (ECMA-376): slide ids live in [256, 2147483647]. */
const MIN_SLIDE_ID = 256

/** ST_SlideId maximum (ECMA-376). One past it is not a slide id, it is a repair prompt. */
const MAX_SLIDE_ID = 2147483647

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
			slides.push(new Slide(this, part, numberValue(attr(sldId, 'id')) ?? 0, index++))
		}
		return slides
	}

	/** Slide dimensions (`p:sldSz`), or `null` if the presentation declares none. */
	get slideSize(): SlideSize | null {
		const root = this.presentationPart.dom.documentElement
		const sldSz = root && firstChild(root, 'p:sldSz')
		if (!sldSz) return null
		const widthEmu = numberValue(attr(sldSz, 'cx'))
		const heightEmu = numberValue(attr(sldSz, 'cy'))
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
	 * The deck's extended document properties (`docProps/app.xml`): the producing
	 * application and its version, the company, and the flat `TitlesOfParts` vector.
	 * `{}` when the deck carries no extended-properties part. The read counterpart of
	 * the write-side `pptx.company` setter — the only one of the four this library
	 * writes from a caller's value; the other three it states about itself.
	 *
	 * A deliberate subset of the part: see {@link ExtendedProperties} for why the
	 * statistics (`Slides`, `Words`, `Paragraphs`, …) are not reported.
	 */
	get appProperties(): ExtendedProperties {
		return readExtendedProperties(this.opc)
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
		const presRels = presentationRels(this)
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
		return importSlideInto(this, source, index, options)
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
	 * `embedFonts` and `rescale` travel per request as well, and both are whole-deck
	 * decisions wearing a per-page spelling, so the batch reconciles them before it
	 * moves anything. A source deck's embedded fonts are carried once when *any* of
	 * its requests asks for them (the list does not record which page uses which
	 * face, so there is nothing finer to carry), and the font parts are part of the
	 * up-front dry run like everything else. `rescale` must **agree** across every
	 * request naming one source: a `'copy'` import rescales the imported layout and
	 * master alongside the page, and those are shared, so a batch that rescaled one
	 * page of a source and not another would leave the second bound to a rescaled
	 * master. Disagreement is `import/rescale-conflict` rather than a silent pick.
	 *
	 * Scope: pages come across under `'copy'` theme semantics (their own layout →
	 * master → theme subgraph, shared parts deduped via the copy registry). There is
	 * still no batch spelling for `theme`, `carryMasterGraphics` or `remapLiterals`;
	 * use {@link importSlide} when you need one of those.
	 */
	importSlides(requests: readonly ImportSlidesRequest[]): Slide[] {
		return importSlidesInto(this, requests)
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
		return importSlideMastersInto(this, source, options)
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
	 * `slide:N` link is repointed at the Nth appended slide's new partname. A chartEx
	 * (Office 2016 — waterfall, funnel, treemap, ...) chart carries too, as its own
	 * `chartEx{N}.xml` part behind the MS chartEx rel, with the style and color-style
	 * sidecars PowerPoint requires alongside it.
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
		return appendSlidesInto(this, source, options)
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
		return importShapeInto(this, target, source, shapeIndex, options)
	}

	/**
	 * Batch form of {@link importShape}: copy several shapes from one source slide
	 * onto `target` in the given order. Media/chart/embedding parts shared by the
	 * lifted shapes (and by earlier imports from the same source deck) are copied
	 * once via the copy registry, and shared images resolve to a single host-slide
	 * relationship. Returns the new {@link Shape}s in `shapeIndices` order.
	 */
	importShapes(target: Slide, source: Slide, shapeIndices: number[], options: ImportShapeOptions = {}): AnyShape[] {
		return importShapesInto(this, target, source, shapeIndices, options)
	}

	/**
	 * @internal Wire a new slide part into `p:sldIdLst` at position `at`, as the private
	 * method below does. Exposed for `presentation-imports.ts`; this one stays because
	 * {@link cloneSlide} and {@link appendSlides} wire their slides in through it too.
	 */
	insertSlidePart(newPart: Part, at?: number): Slide {
		return this.#insertSlidePart(newPart, at)
	}

	/**
	 * @internal The parts whose geometry a rescale has already rewritten. Exposed for
	 * `presentation-imports.ts` as a live set the callee adds to — the memo is what makes a
	 * layout or master shared across repeated imports from one source scale exactly once,
	 * so like the copy registry it has to outlive the call that fills it.
	 */
	get rescaledParts(): Set<string> {
		return this.#rescaledParts
	}

	/**
	 * @internal Open an import out of `source`: this deck as the destination, paired with
	 * the copy registry for that package (created on first use). The registry is held on
	 * the class rather than in the context because it must outlive any one call — that is
	 * what makes a second import from the same source reuse the layout, master, theme, and
	 * media it already copied instead of duplicating them, and it is why the import bodies
	 * in `presentation-imports.ts` call back here rather than carrying it with them.
	 */
	importContext(source: OpcPackage): ImportContext {
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
		const presRels = presentationRels(this)

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
		const sldIdLst = getOrAddChild(root, 'p:sldIdLst', PRESENTATION_AFTER_SLD_ID_LST)
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

	/**
	 * A slide id one past the highest existing, but at least ST_SlideId's minimum.
	 *
	 * The module named and enforced that minimum and not the maximum, so a deck near the ceiling
	 * got an out-of-range `p:sldId/@id` written with no diagnostic — the package then needs
	 * repair, for a reason nothing reported. Past the ceiling there is no "next" id to hand out,
	 * so the allocation falls back to the lowest id in range that the deck is not already using;
	 * a deck holding every one of the two billion has no legal answer and says so.
	 */
	#nextSlideId(sldIds: ReturnType<typeof getElements>): number {
		const used = new Set<number>()
		let max = MIN_SLIDE_ID - 1
		for (const sldId of sldIds) {
			const id = numberValue(attr(sldId, 'id'))
			if (id === null) continue
			used.add(id)
			if (id > max) max = id
		}
		if (max < MAX_SLIDE_ID) return max + 1
		for (let id = MIN_SLIDE_ID; id <= MAX_SLIDE_ID; id++) {
			if (!used.has(id)) return id
		}
		throw new InvalidOptionError(
			'slide/id-space-exhausted',
			`this deck already uses every slide id ST_SlideId allows (${MIN_SLIDE_ID}-${MAX_SLIDE_ID}); no slide can be added.`
		)
	}

	/** Re-emit the package; untouched parts stay byte-identical (see `OpcPackage.save`). */
	async save(): Promise<Uint8Array> {
		return this.opc.save()
	}
}
