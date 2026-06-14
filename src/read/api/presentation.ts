/**
 * Read-model entry point: `Presentation` wraps an `OpcPackage` and exposes a
 * navigable, typed view of the deck (slides → shapes → text), backed by the
 * live DOM so the same nodes can later be mutated.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { relativePartName, relsPartNameFor } from '../opc/partnames.js'
import { attr, createElement, firstChild, getElements, getOrAddChild, intValue, removeChildrenByQName, setAttr, type Element } from '../oxml/dom.js'
import { flattenSlide, parseClrMap, parseClrScheme, type FlattenContext } from '../oxml/theme.js'
import { Slide } from './slide.js'

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'

const SLIDE_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const SLIDE_LAYOUT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'

/** ST_SlideId minimum (ECMA-376): slide ids live in [256, 2147483647]. */
const MIN_SLIDE_ID = 256

/** ST_SlideLayoutId minimum (ECMA-376): slide-layout ids start at 2147483648. */
const MIN_SLIDE_LAYOUT_ID = 2147483648

/** Slide dimensions, in both EMU (the OOXML unit) and inches. */
export interface SlideSize {
	widthEmu: number
	heightEmu: number
	widthIn: number
	heightIn: number
}

/** Options for {@link Presentation.importSlide}. */
export interface ImportSlideOptions {
	/**
	 * How the imported slide relates to themes.
	 *
	 * - `'copy'` (default): bring the slide's own `slideLayout → slideMaster →
	 *   theme` subgraph across, so the slide renders byte-for-byte as authored.
	 *   A deck stitched from N sources then carries N themes/masters.
	 * - `'preserve'`: *flatten then attach* — bake the source theme's colours and
	 *   style-matrix fills into the slide XML (so its pixels do not change), then
	 *   bind it to this deck's existing master/layout instead of importing the
	 *   source theme. The result is a single-theme file whose imported slides are
	 *   theme-independent: it fixes renderers that mis-resolve per-slide scheme
	 *   colours against the wrong (first) theme, and tidies the deck for handoff.
	 *
	 *   To stay faithful across the rebind, `preserve` also carries the slide's
	 *   effective background and each placeholder run's *inherited* colour (from
	 *   the source layout/master text styles) explicitly onto the slide. It does
	 *   **not** carry decorative graphics that live on the source master/layout
	 *   shape tree (logos, accent shapes): those belong to the master `preserve`
	 *   deliberately drops, so re-add such branding as explicit slide elements.
	 */
	theme?: 'copy' | 'preserve'
}

export class Presentation {
	#presentationPart: Part | undefined
	/**
	 * Per-source copy registry for {@link importSlide}: source `OpcPackage` →
	 * (source partname → partname allocated in this package). Lets parts shared
	 * across imports from the same source deck (layout, master, theme, media) be
	 * copied once and reused on later calls.
	 */
	#importRegistry = new Map<OpcPackage, Map<string, string>>()

	private constructor(readonly opc: OpcPackage) {}

	/** Open a `.pptx` from bytes and wrap it as a navigable `Presentation`. */
	static async load(input: OpcInput): Promise<Presentation> {
		return new Presentation(await OpcPackage.load(input))
	}

	/** Wrap an already-loaded OPC package (e.g. from the lower-level API). */
	static fromPackage(opc: OpcPackage): Presentation {
		return new Presentation(opc)
	}

	/** The main presentation part (`/ppt/presentation.xml`), resolved via the package `officeDocument` relationship. */
	get presentationPart(): Part {
		if (this.#presentationPart) return this.#presentationPart
		const packageRels = this.opc.relationshipsFor('/')
		const officeDocument = packageRels.byType(OFFICE_DOCUMENT_REL)
		if (officeDocument.length !== 1) {
			throw new Error(`Expected exactly one officeDocument relationship, found ${officeDocument.length}`)
		}
		const partName = packageRels.resolveTarget(officeDocument[0].id)
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
	 * Duplicate the slide at `index` and append the copy at the end of the deck,
	 * returning it. The new slide part copies the source bytes verbatim and
	 * shares the source's relationship targets (layout, images, …) by copying its
	 * `.rels`; a new presentation→slide relationship and a `p:sldId` entry are
	 * wired up. Marks the presentation part dirty.
	 *
	 * Note: relationships are copied as-is, so a source slide that owns a
	 * one-to-one part (e.g. a notes slide) would end up shared with the clone.
	 */
	cloneSlide(index: number): Slide {
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

		// 3. Wire the new slide into the presentation (rel + p:sldId entry).
		return this.#appendSlidePart(newPart)
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
	 * master/layout — see {@link ImportSlideOptions}.
	 *
	 * v1 limitations: the source slide size must equal this presentation's (no
	 * geometry rescaling); source notes are dropped; fonts embedded via
	 * `presentation.xml` are not copied.
	 */
	importSlide(source: Presentation, index: number, options: ImportSlideOptions = {}): Slide {
		const sourceSlide = source.slides[index]
		if (!sourceSlide) throw new Error(`No slide at index ${index} to import`)

		// 1. Pre-flight: v1 does not rescale geometry, so slide sizes must match.
		const target = this.slideSize
		const incoming = source.slideSize
		if (!target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu) {
			const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
			throw new Error(`importSlide requires equal slide sizes; target is ${fmt(target)}, source is ${fmt(incoming)}`)
		}

		// 2. Copy the slide and its dependencies. 'preserve' flattens the theme into
		//    the slide and attaches it to this deck's master; 'copy' brings the
		//    source theme subgraph across wholesale.
		const newPartName = options.theme === 'preserve' ? this.#importSlidePreserve(source, sourceSlide) : this.#copyPart(source.opc, sourceSlide.partName)
		const newPart = this.opc.part(newPartName)
		if (!newPart) throw new Error(`Imported slide part went missing: ${newPartName}`)

		// 3. Wire the new slide into the presentation (rel + p:sldId entry).
		return this.#appendSlidePart(newPart)
	}

	/**
	 * Import a slide in `preserve` mode: copy the slide part, flatten its source
	 * theme into the slide XML (scheme colours + style-matrix fills baked to
	 * literals), copy its non-theme dependencies (media, charts, …), and point its
	 * `slideLayout` rel at this deck's existing layout. Returns the new partname.
	 */
	#importSlidePreserve(source: Presentation, sourceSlide: Slide): string {
		const destLayout = this.#destinationLayoutPartName()
		const ctx = this.#sourceFlattenContext(source.opc, sourceSlide.partName)

		// Copy the slide bytes into a fresh partname, then flatten its own DOM (a
		// distinct copy, so the source package is never mutated).
		const sourcePart = source.opc.part(sourceSlide.partName)
		if (!sourcePart) throw new Error(`importSlide: source package has no part ${sourceSlide.partName}`)
		const newPartName = this.opc.reservePartNameLike(sourceSlide.partName)
		const newPart = this.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
		const slideRoot = newPart.dom.documentElement
		if (!slideRoot) throw new Error(`Imported slide ${newPartName} has no root element`)
		flattenSlide(slideRoot, ctx)
		newPart.markDirty()

		// Rebuild the slide's relationships: drop notes, repoint slideLayout at the
		// destination layout, and copy every other internal target (media/charts).
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
			const newTarget = this.#copyPart(source.opc, sourceRels.resolveTarget(rel.id))
			targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTarget))
		}
		return newPartName
	}

	/**
	 * The partname of the layout this deck's slides should attach to in `preserve`
	 * mode: the first layout of the first slide master. Throws when the deck has no
	 * master/layout to attach to (a deck pptxgenjs always provides).
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
		const layoutPartName = this.#resolveSingleRel(sourceOpc, slidePartName, SLIDE_LAYOUT_REL)
		const masterPartName = layoutPartName && this.#resolveSingleRel(sourceOpc, layoutPartName, SLIDE_MASTER_REL)
		const themePartName = masterPartName && this.#resolveSingleRel(sourceOpc, masterPartName, THEME_REL)

		const masterRoot = masterPartName ? (sourceOpc.part(masterPartName)?.dom.documentElement ?? null) : null
		const masterClrMap = masterRoot ? firstChild(masterRoot, 'p:clrMap') : null
		const layoutRoot = layoutPartName ? (sourceOpc.part(layoutPartName)?.dom.documentElement ?? null) : null

		// A slide's clrMapOvr/overrideClrMapping (if present) wins over the master map.
		const slideRoot = sourceOpc.part(slidePartName)?.dom.documentElement
		const clrMapOvr = slideRoot ? firstChild(slideRoot, 'p:clrMapOvr') : null
		const override = clrMapOvr ? firstChild(clrMapOvr, 'a:overrideClrMapping') : null

		const themeRoot = themePartName ? sourceOpc.part(themePartName)?.dom.documentElement : null
		const themeElements = themeRoot ? firstChild(themeRoot, 'a:themeElements') : null

		return {
			clrMap: parseClrMap(override ?? masterClrMap),
			clrScheme: parseClrScheme(themeElements ? firstChild(themeElements, 'a:clrScheme') : null),
			fmtScheme: themeElements ? firstChild(themeElements, 'a:fmtScheme') : null,
			inheritedBackground: this.#effectiveBackground(sourceOpc, slideRoot ?? null, layoutPartName, masterPartName),
			layoutRoot,
			masterRoot,
		}
	}

	/**
	 * The background the slide effectively inherits from its source subgraph: the
	 * layout's `p:bg`, else the master's. Returns `null` when the slide carries its
	 * own `p:bg` (it stays on the slide and is flattened directly) or none exists.
	 */
	#effectiveBackground(sourceOpc: OpcPackage, slideRoot: Element | null, layoutPartName: string | null, masterPartName: string | null): Element | null {
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

	/** The copy registry for one source package (created on first use). */
	#registryFor(sourceOpc: OpcPackage): Map<string, string> {
		let registry = this.#importRegistry.get(sourceOpc)
		if (!registry) {
			registry = new Map()
			this.#importRegistry.set(sourceOpc, registry)
		}
		return registry
	}

	/**
	 * Copy `sourcePartName` (and, recursively, every internal part it references)
	 * from `sourceOpc` into this package, returning the new partname. Idempotent
	 * per source package via the copy registry. Relationship ids are preserved so
	 * the copied part body's `r:id`/`r:embed` references stay valid; targets are
	 * rewritten to the freshly-allocated partnames. Notes relationships are
	 * dropped. A copied `slideMaster` does not drag in all its sibling layouts —
	 * each imported `slideLayout` wires itself into the master instead (see
	 * {@link #linkLayoutIntoMaster}).
	 */
	#copyPart(sourceOpc: OpcPackage, sourcePartName: string): string {
		const registry = this.#registryFor(sourceOpc)
		const existing = registry.get(sourcePartName)
		if (existing) return existing

		const sourcePart = sourceOpc.part(sourcePartName)
		if (!sourcePart) throw new Error(`importSlide: source package has no part ${sourcePartName}`)

		const newPartName = this.opc.reservePartNameLike(sourcePartName)
		this.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
		// Record before recursing so the master↔layout cycle terminates.
		registry.set(sourcePartName, newPartName)

		const isMaster = sourcePart.contentType === SLIDE_MASTER_CONTENT_TYPE
		const sourceRels = sourceOpc.relationshipsFor(sourcePartName)
		const targetRels = this.opc.relationshipsFor(newPartName)
		for (const rel of sourceRels) {
			// Notes pull in a notesMaster + its own theme; an imported slide does not need them.
			if (rel.type === NOTES_SLIDE_REL) continue
			// Lean master: skip its layout rels; copied layouts re-link themselves.
			if (isMaster && rel.type === SLIDE_LAYOUT_REL) continue
			if (rel.targetMode === 'External') {
				targetRels.addWithId(rel.id, rel.type, rel.target, 'External')
				continue
			}
			const newTargetPartName = this.#copyPart(sourceOpc, sourceRels.resolveTarget(rel.id))
			targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTargetPartName))
		}

		if (isMaster) this.#clearLayoutIdList(newPartName)
		if (sourcePart.contentType === SLIDE_LAYOUT_CONTENT_TYPE) {
			this.#linkLayoutIntoMaster(sourceOpc, sourceRels, newPartName)
		}

		return newPartName
	}

	/** Empty a freshly-copied master's `p:sldLayoutIdLst`; copied layouts repopulate it. */
	#clearLayoutIdList(masterPartName: string): void {
		const masterPart = this.opc.part(masterPartName)
		const root = masterPart?.dom.documentElement
		const lst = root && firstChild(root, 'p:sldLayoutIdLst')
		if (!masterPart || !lst) return
		removeChildrenByQName(lst, ['p:sldLayoutId'])
		masterPart.markDirty()
	}

	/**
	 * Wire a just-copied layout into its (already-copied) master: add a
	 * master→layout relationship and append a `p:sldLayoutId` entry. Called once
	 * per copied layout, so the master accumulates exactly the imported layouts.
	 */
	#linkLayoutIntoMaster(sourceOpc: OpcPackage, layoutSourceRels: ReturnType<OpcPackage['relationshipsFor']>, layoutPartName: string): void {
		const masterRel = layoutSourceRels.byType(SLIDE_MASTER_REL)[0]
		if (!masterRel) return
		const masterPartName = this.#registryFor(sourceOpc).get(layoutSourceRels.resolveTarget(masterRel.id))
		if (!masterPartName) return
		const masterPart = this.opc.part(masterPartName)
		const root = masterPart?.dom.documentElement
		if (!masterPart || !root) return

		const masterRels = this.opc.relationshipsFor(masterPartName)
		const relId = masterRels.add(SLIDE_LAYOUT_REL, relativePartName(masterPartName, layoutPartName)).id
		const lst = getOrAddChild(root, 'p:sldLayoutIdLst', ['p:transition', 'p:timing', 'p:hf', 'p:txStyles', 'p:extLst'])
		const entry = createElement(masterPart.dom, 'p:sldLayoutId')
		setAttr(entry, 'id', String(this.#nextSlideLayoutId(lst)))
		setAttr(entry, 'r:id', relId)
		lst.appendChild(entry)
		masterPart.markDirty()
	}

	/** Wire a new slide part into `p:sldIdLst` (rel + `p:sldId`), append it, and return it. */
	#appendSlidePart(newPart: Part): Slide {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const relId = presRels.add(SLIDE_REL, relativePartName(presPart.partName, newPart.partName)).id

		const root = presPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (!sldIdLst) throw new Error('presentation.xml has no p:sldIdLst to append a slide to')
		const existing = getElements(sldIdLst, 'p:sldId')
		const newIndex = existing.length
		const newSlideId = this.#nextSlideId(existing)
		const sldId = createElement(presPart.dom, 'p:sldId')
		setAttr(sldId, 'id', String(newSlideId))
		setAttr(sldId, 'r:id', relId)
		sldIdLst.appendChild(sldId)
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

	/** A slide-layout id one past the highest in `sldLayoutIdLst`, floored at ST_SlideLayoutId's minimum. */
	#nextSlideLayoutId(sldLayoutIdLst: Element): number {
		let max = MIN_SLIDE_LAYOUT_ID - 1
		for (const entry of getElements(sldLayoutIdLst, 'p:sldLayoutId')) {
			const id = intValue(attr(entry, 'id'))
			if (id !== null && id > max) max = id
		}
		return max + 1
	}

	/** Re-emit the package; untouched parts stay byte-identical (see `OpcPackage.save`). */
	async save(): Promise<Uint8Array> {
		return this.opc.save()
	}
}
