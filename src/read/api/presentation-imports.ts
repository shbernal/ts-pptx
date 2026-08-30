/**
 * ts-pptx: Presentation import entry points
 *
 * The four ways a `Presentation` takes content out of another open package: one page
 * ({@link importSlide}), a batch of pages ({@link importSlides}), a master with its layout
 * family ({@link importSlideMasters}), and shapes ({@link importShape} / {@link importShapes}).
 * Each is the body of the same-named `Presentation` method, which keeps the public contract in
 * its doc comment and delegates here -- so the class reads as its own table of contents and one
 * feature's internals stop outnumbering the read model's surface.
 *
 * These are free functions taking the destination deck, the shape every module in `ops/`
 * already has. The difference, and the cost of this split, is that they are not independent of
 * `Presentation` the way an `ops/` module is: three things they need stay on the class and are
 * reached back through its `@internal` accessors. `insertSlidePart` stays because
 * `cloneSlide` and `appendSlides` wire slides in through it too. `importContext` and
 * `rescaledParts` stay for a different reason -- both are memos that must outlive any one
 * call, which is what makes a second import from the same source reuse what it already copied
 * and stops a shared layout being scaled twice. Neither reason is removable by moving more
 * code, so three `#private` members were widened into an internal API. Paid deliberately: the
 * coupling was already there, and naming it here is what keeps it visible.
 */

import type { OpcPackage } from '../opc/package.js'
import { OOXML_NS, attr, createElement, firstChild, intValue, setAttr } from '../oxml/dom.js'
import { cSldName, nthShapeChild } from '../oxml/slide-dom.js'
import { InternalError, InvalidOptionError, PackageReadError, UnsupportedFeatureError } from '../../errors.js'
import { NOTES_MASTER_REL } from '../../ooxml/rel-types.js'
import { carryShapeAnimations } from './animation.js'
import { flattenShape } from './ops/flatten.js'
import { wrapShapeElement, type AnyShape } from './shapes.js'
import type { Presentation } from './presentation.js'
import type { Slide } from './slide.js'
import type {
	ImportShapeOptions,
	ImportSlideMastersOptions,
	ImportSlideOptions,
	ImportSlidesRequest,
	ImportedSlideMaster,
} from './presentation-types.js'
import { rewriteCarriedRels } from './ops/carried-rels.js'
import { carryEmbeddedFonts, checkEmbeddedFontsCopyable } from './ops/embedded-fonts.js'
import { sourceFlattenContext } from './ops/flatten-context.js'
import { importSlidePreserve, importSlideRestyle } from './ops/import-slide.js'
import { promoteMasters } from './ops/master-registry.js'
import { carryNotes } from './ops/notes-master.js'
import { layoutPartNamesOf, slideMasterPartNames } from './ops/part-index.js'
import { checkSelectionCopyable, copyPart, copySlidePart, newOwnedScope, type ImportContext } from './ops/part-copy.js'
import { computeRescale, rescaleSpTree, type RescaleTransform } from './ops/rescale.js'
import { rescaleImportedGeometry } from './ops/rescale-import.js'
import { requireEqualSlideSize, requireKnownSlideSizes, slideSizesMatch } from './ops/slide-size.js'
import { carryTableStyles } from './ops/table-styles.js'

/** Body of {@link Presentation.importSlide}, whose doc comment carries the contract. */
export function importSlide(
	deck: Presentation,
	source: Presentation,
	index: number,
	options: ImportSlideOptions = {}
): Slide {
	const sourceSlide = source.slides[index]
	if (!sourceSlide) throw new InvalidOptionError('slide/index-out-of-range', `No slide at index ${index} to import`)

	// 1. Pre-flight: slide sizes must match unless the caller opts into a rescale. Either way
	//    both decks have to declare one — a rescale cannot be computed from a size that is
	//    not there, and neither can a comparison.
	const target = deck.slideSize
	const incoming = source.slideSize
	if (options.rescale) requireKnownSlideSizes(target, incoming, 'importSlide rescale')
	else requireEqualSlideSize(target, incoming, 'importSlide', "pass { rescale: 'fit' | 'stretch' } to rescale")
	const sizesDiffer = !slideSizesMatch(target, incoming)

	// 2. Copy the slide and its dependencies. 'preserve' flattens the theme into
	//    the slide and attaches it to this deck's master; 'restyle' attaches it
	//    to this deck's master with theme refs left symbolic (re-brand); 'copy'
	//    brings the source theme subgraph across wholesale.
	const importCtx = deck.importContext(source.opc)
	const newPartName =
		options.theme === 'preserve'
			? importSlidePreserve(deck, importCtx, source, sourceSlide, options.carryMasterGraphics === true)
			: options.theme === 'restyle'
				? importSlideRestyle(
						deck,
						importCtx,
						source,
						sourceSlide,
						options.carryMasterGraphics === true,
						options.remapLiterals === true
					)
				: copySlidePart(importCtx, sourceSlide.partName)
	const newPart = deck.opc.part(newPartName)
	if (!newPart) throw new InternalError('import/part-went-missing', `Imported slide part went missing: ${newPartName}`)

	// 2b. Rescale the imported geometry to this deck's canvas when sizes differ.
	if (sizesDiffer && options.rescale && target && incoming) {
		rescaleImportedGeometry(
			deck,
			deck.rescaledParts,
			newPartName,
			options.theme,
			incoming,
			target,
			options.rescale === true ? 'fit' : options.rescale
		)
	}

	// 3. Wire the new slide into the presentation (rel + p:sldId entry) at `at`.
	const slide = deck.insertSlidePart(newPart, options.at)

	// 4. Optionally carry the source slide's speaker notes. The slide copy above
	//    drops the notesSlide rel (both copyPart and importSlideRebind do); this
	//    re-adds it wired to the new slide and merged onto a single notesMaster.
	if (options.importNotes) carryNotes(deck, source, importCtx, sourceSlide.partName, newPartName)

	// 5. Optionally carry the source deck's embedded fonts (presentation-level, so
	//    a separate traversal from the slide-part copy chain above).
	if (options.embedFonts) carryEmbeddedFonts(deck, source, importCtx)

	return slide
}

/** Body of {@link Presentation.importSlides}, whose doc comment carries the contract. */
export function importSlides(deck: Presentation, requests: readonly ImportSlidesRequest[]): Slide[] {
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
	if (!deck.presentationPart.dom.documentElement) {
		throw new PackageReadError(
			'package/part-has-no-root',
			'presentation.xml has no document element to insert slides into'
		)
	}

	const finalSlideCount = deck.slides.length + resolved.length
	// Which pages each source is being asked for. A page may appear in several
	// requests: the set is what the dry run walks, and the per-request output
	// parts are allocated in step 2. `notesPages` is the subset whose notes are
	// coming too, which the dry run has to walk past the dropped notes rel.
	// `fontSources` and `rescaleBySource` are the two deck-level options wearing a
	// per-request spelling, reconciled here so the copy below reads one answer per
	// source rather than re-deciding per page.
	const selectedPages = new Map<OpcPackage, Set<string>>()
	const notesPages = new Map<OpcPackage, Set<string>>()
	const fontSources = new Map<OpcPackage, Presentation>()
	const rescaleBySource = new Map<OpcPackage, 'fit' | 'stretch' | false>()
	for (const request of resolved) {
		let pages = selectedPages.get(request.source.opc)
		if (!pages) {
			pages = new Set()
			selectedPages.set(request.source.opc, pages)
		}
		pages.add(request.sourceSlide.partName)
		if (request.embedFonts) fontSources.set(request.source.opc, request.source)
		// A rescale rewrites the shared imported layout and master as well as the
		// page, so the requests naming one source have to agree on it: rescaling one
		// of a source's pages and not another would leave the second aligned against
		// a master that moved under it. Normalized first, so `true` and `'fit'` are
		// the same answer rather than a conflict.
		const rescale = request.rescale === true ? 'fit' : (request.rescale ?? false)
		const agreed = rescaleBySource.get(request.source.opc)
		if (agreed !== undefined && agreed !== rescale) {
			throw new InvalidOptionError(
				'import/rescale-conflict',
				`importSlides: requests from one source must agree on \`rescale\`; got ${JSON.stringify(agreed)} and ${JSON.stringify(rescale)}`
			)
		}
		rescaleBySource.set(request.source.opc, rescale)
		if (!request.importNotes) continue
		let withNotes = notesPages.get(request.source.opc)
		if (!withNotes) {
			withNotes = new Set()
			notesPages.set(request.source.opc, withNotes)
		}
		withNotes.add(request.sourceSlide.partName)
	}

	const target = deck.slideSize
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
		// A rescale needs both sizes known but not equal; without one they must match,
		// and the hint names the spelling that answers the mismatch.
		if (rescaleBySource.get(request.source.opc))
			requireKnownSlideSizes(target, request.source.slideSize, 'importSlides rescale')
		else
			requireEqualSlideSize(
				target,
				request.source.slideSize,
				'importSlides',
				"pass { rescale: 'fit' | 'stretch' } on the request to rescale"
			)
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
	const copyMaster = deck.opc.relationshipsFor(deck.presentationPart.partName).byType(NOTES_MASTER_REL).length === 0
	for (const [sourceOpc, pages] of selectedPages) {
		checkSelectionCopyable(sourceOpc, deck.importContext(sourceOpc).registry, pages, {
			pages: notesPages.get(sourceOpc) ?? new Set(),
			copyMaster,
		})
	}
	// The font carry runs after the copy and reaches parts the page graph never
	// touches, so it gets its own dry run for the same reason the notes do.
	for (const source of fontSources.values()) checkEmbeddedFontsCopyable(source)

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
		const newPartName = deck.opc.reservePartNameLike(request.sourceSlide.partName)
		const destPart = deck.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
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
		const base = deck.importContext(sourceOpc)
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
			deck,
			request.source,
			deck.importContext(request.source.opc),
			request.sourceSlide.partName,
			request.destPart.partName
		)
	}

	// 3c. Rescale each rescaling source's pages onto this deck's canvas. Nothing
	//     to do where the sizes already match, and the memo makes the shared layout
	//     and master — which `'copy'` semantics rescale alongside the page — move
	//     exactly once however many of that source's pages the batch brought over.
	for (const request of planned) {
		const mode = rescaleBySource.get(request.source.opc)
		if (!mode) continue
		const incoming = request.source.slideSize
		if (!target || !incoming || slideSizesMatch(target, incoming)) continue
		rescaleImportedGeometry(deck, deck.rescaledParts, request.destPart.partName, 'copy', incoming, target, mode)
	}

	// 3d. Carry the embedded fonts of every source that asked, once per source. A
	//     face this deck already embeds is reused rather than duplicated, and the
	//     binaries come across through the same per-source registry the pages used,
	//     so a repeated import copies each face once.
	for (const [sourceOpc, source] of fontSources) {
		carryEmbeddedFonts(deck, source, deck.importContext(sourceOpc))
	}

	// 4. Wire into p:sldIdLst at each requested final position. Ascending order
	//    makes the raw final index the correct insertion point at every step:
	//    earlier inserts all sit before it, so they shift it by exactly the
	//    number of entries the final position already counts. The result is
	//    written back by request index, so `result[i]` is `requests[i]`'s page
	//    whatever order the positions were given in.
	const added: Slide[] = []
	for (const request of [...planned].sort((left, right) => left.outputIndex - right.outputIndex)) {
		added[request.requestIndex] = deck.insertSlidePart(request.destPart, request.outputIndex)
	}
	return added
}

/** Body of {@link Presentation.importSlideMasters}, whose doc comment carries the contract. */
export function importSlideMasters(
	deck: Presentation,
	source: Presentation,
	options: ImportSlideMastersOptions = {}
): ImportedSlideMaster[] {
	if (options.requireEqualSize !== false) {
		requireEqualSlideSize(
			deck.slideSize,
			source.slideSize,
			'importSlideMasters',
			'pass { requireEqualSize: false } to override'
		)
	}

	const pickMaster = options.masters ?? (() => true)
	const pickLayout = options.layouts ?? (() => true)

	const ctx = deck.importContext(source.opc)
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
	if (options.embedFonts) carryEmbeddedFonts(deck, source, deck.importContext(source.opc))
	if (options.tableStyles) carryTableStyles(deck, source.opc)
	if (options.primary)
		promoteMasters(
			deck,
			imported.map((m) => m.partName)
		)

	return imported
}

/** Body of {@link Presentation.importShape}, whose doc comment carries the contract. */
export function importShape(
	deck: Presentation,
	target: Slide,
	source: Slide,
	shapeIndex: number,
	options: ImportShapeOptions = {}
): AnyShape {
	const [shape] = importShapes(deck, target, source, [shapeIndex], options)
	if (!shape)
		throw new InvalidOptionError(
			'shape/index-out-of-range',
			`importShape: source slide has no shape at index ${shapeIndex}`
		)
	return shape
}

/** Body of {@link Presentation.importShapes}, whose doc comment carries the contract. */
export function importShapes(
	deck: Presentation,
	target: Slide,
	source: Slide,
	shapeIndices: number[],
	options: ImportShapeOptions = {}
): AnyShape[] {
	if (target.presentation !== deck)
		throw new InvalidOptionError('slide/foreign-target', 'importShape: target slide must belong to deck presentation')

	// Pre-flight: slide sizes must match unless { rescale } opts into scaling the
	// lifted geometry onto this canvas (computed once, applied per shape below).
	const targetSize = deck.slideSize
	const sourceSize = source.presentation.slideSize
	let transform: RescaleTransform | null = null
	if (!slideSizesMatch(targetSize, sourceSize)) {
		if (!options.rescale) requireEqualSlideSize(targetSize, sourceSize, 'importShape', 'or { rescale }')
		const [known, incoming] = requireKnownSlideSizes(targetSize, sourceSize, 'importShape rescale')
		transform = computeRescale(incoming, known, options.rescale === 'stretch' ? 'stretch' : 'fit')
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
		throw new PackageReadError('slide/no-shape-tree', `importShape: target slide ${target.partName} has no shape tree`)
	const targetDoc = spTree.ownerDocument
	if (!targetDoc)
		throw new InternalError('oxml/node-has-no-document', 'importShape: target slide DOM has no owner document')

	const theme = options.theme ?? 'preserve'
	const sourceOpc = source.presentation.opc
	const sourceRels = sourceOpc.relationshipsFor(source.partName)
	const targetRels = deck.opc.relationshipsFor(target.partName)
	// One rel-id map across the batch so shapes sharing a source image share a rel.
	const relIdMap = new Map<string, string>()
	// preserve: build the source theme context once; copy/restyle need none.
	const ctx = theme === 'preserve' ? sourceFlattenContext(sourceOpc, source.partName) : null
	const importCtx = deck.importContext(sourceOpc)

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
