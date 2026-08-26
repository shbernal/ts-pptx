/**
 * The part-copy traversal: pulling a part -- and, recursively, every internal part
 * it references -- out of a source package into a destination deck, rewriting
 * relationship targets to the freshly-allocated partnames as it goes.
 *
 * This is the source-side half of the import machinery; the destination-side
 * bookkeeping it hands off to lives in `master-registry.ts`. It takes an
 * {@link ImportContext} rather than a `Presentation`, so it stays independent of
 * the class that calls it.
 */

import type { OpcPackage } from '../../opc/package.js'
import { relativePartName } from '../../opc/partnames.js'
import type { DeckTarget } from './deck-target.js'
import { addLayoutToMaster, clearLayoutIdList, registerMaster } from './master-registry.js'
import {
	NOTES_MASTER_REL,
	NOTES_SLIDE_CONTENT_TYPE,
	NOTES_SLIDE_REL,
	SLIDE_LAYOUT_CONTENT_TYPE,
	SLIDE_LAYOUT_REL,
	SLIDE_MASTER_CONTENT_TYPE,
	SLIDE_MASTER_REL,
	SLIDE_REL,
} from '../../../ooxml/rel-types.js'
import { isSharedByPageCopies } from './page-owned.js'
import { InvalidOptionError, PackageReadError } from '../../../errors.js'

/**
 * One import in progress: where parts are going, where they are coming from, and
 * what has already come across.
 *
 * `source` and `registry` are one fact, not two — a registry maps partnames *of
 * that package*, so pairing it with any other package silently returns partnames
 * for parts that were never copied. Holding them in a single value makes the
 * pairing a type invariant rather than a call-site convention; mint one with
 * `Presentation`'s factory, which reads through to the per-source registry that
 * outlives any one call (see {@link copyPart}'s idempotence).
 */
export interface ImportContext {
	/** The deck being copied into. */
	readonly dest: DeckTarget
	/** The package being copied out of. */
	readonly source: OpcPackage
	/** Source partname → the partname allocated for it in `dest`. */
	readonly registry: Map<string, string>
	/**
	 * The pages this call is materializing, and where each is headed: source slide
	 * partname → the destination partname pre-allocated for it. A page named here
	 * gets its own part even when the registry already holds a copy of it, which is
	 * what lets one source page be imported more than once.
	 *
	 * Both import entry points set it — {@link Presentation.importSlides} for the
	 * whole batch, {@link copySlidePart} for the single page of one
	 * {@link Presentation.importSlide}. For the batch it carries a second meaning:
	 * `slide → slide` relationships resolve only within the set, so an imported page
	 * may link to another *selected* page (rewritten to its fresh partname) but must
	 * not drag an unselected source page across as a dependency.
	 *
	 * One destination per source page is deliberate: a page is a link target as well
	 * as an output, and a map to *many* destinations would make every jump link
	 * ambiguous. A batch asked for the same page N times therefore runs the
	 * traversal N times, each round naming that page's Nth reserved partname and the
	 * other pages' first — see `importSlides` step 3.
	 */
	readonly selection?: SelectionPlan
}

/**
 * The parts one page copy has taken for itself: source partname → the copy made
 * for *this* page, never entered in the copy registry. A page owns a chart, a
 * SmartArt diagram, an OLE embedding and their subtrees; sharing one of those
 * with a second copy of the page writes a deck PowerPoint refuses to open, which
 * is what {@link isSharedByPageCopies} draws the line for.
 *
 * One scope per page copy, opened by {@link copyPart} when it reaches a page of
 * the selection plan and by {@link importSlideRebind} for the page it rebinds.
 */
export type OwnedScope = Map<string, string>

/** Open an ownership scope for one page copy. */
export function newOwnedScope(): OwnedScope {
	return new Map()
}

/** Which pages of one source package an import is materializing, and where each is headed. */
export interface SelectionPlan {
	/** Source slide partname → the destination partname reserved for it. */
	readonly destinations: ReadonlyMap<string, string>
}

/**
 * Copy one slide page across as a part of its own, deduping everything under it
 * but never the page itself. This is the `theme: 'copy'` arm of
 * {@link Presentation.importSlide}.
 *
 * The page is the one part an import is *not* allowed to share. `copyPart`'s
 * registry idempotence is right for a theme, master, layout or image, where a
 * second copy is waste; applied to the slide it made a repeated import of the
 * same source page return the first copy's partname, and the caller then wired a
 * second `p:sldId` to a part that already had one — a package PowerPoint refuses
 * to open (0x80070570), with nothing in the read model to show for it. So the
 * page goes into the selection plan `copyPart` already honours for a batch, which
 * re-materializes exactly the named pages and leaves the dedup of their
 * dependencies alone.
 *
 * @param ctx             the open import out of the source package
 * @param sourcePartName  partname of the source slide to bring across
 * @return                partname of the new slide part in `ctx.dest`
 */
export function copySlidePart(ctx: ImportContext, sourcePartName: string): string {
	const destinations = new Map([[sourcePartName, ctx.dest.opc.reservePartNameLike(sourcePartName)]])
	return copyPart({ ...ctx, selection: { destinations } }, sourcePartName)
}

/**
 * Copy `sourcePartName` (and, recursively, every internal part it references)
 * from `ctx.source` into `ctx.dest`, returning the new partname. Idempotent
 * per source package via the copy registry. Relationship ids are preserved so
 * the copied part body's `r:id`/`r:embed` references stay valid; targets are
 * rewritten to the freshly-allocated partnames. Notes relationships are
 * dropped. A copied `slideMaster` does not drag in all its sibling layouts —
 * each imported `slideLayout` wires itself into the master instead (see
 * {@link linkLayoutIntoMaster}).
 *
 * With `ctx.selection`, partnames for the selected slides were already reserved
 * by the caller ({@link Presentation.importSlides}); this traversal wires their
 * relationships to each other instead of re-copying them. The batch's rule that
 * a `slide → slide` link may not leave the selection is enforced ahead of the
 * copy by {@link checkSelectionCopyable}, which also proves every part this
 * traversal will reach exists — so once copying starts there is nothing left to
 * throw, and a rejected batch never leaves a half-copied deck behind.
 *
 * Idempotence stops at the page's own parts. Reaching a page opens an
 * {@link OwnedScope}, and everything under it that {@link isSharedByPageCopies}
 * does not clear for sharing — a chart, a diagram, an OLE embedding, each with
 * its own subtree — is copied into that scope instead of the registry, so the
 * next copy of the page gets parts of its own. Pass `owned` to open the scope
 * from outside, as the rebinding import paths do for the page they build
 * themselves.
 */
export function copyPart(ctx: ImportContext, sourcePartName: string, owned?: OwnedScope): string {
	// Inside a page's ownership scope the registry is not consulted at all: the
	// point of the scope is that this page copy gets parts of its own, and the
	// scope's own map is what keeps a part two of its relationships reach from
	// being copied twice within the one copy.
	if (owned) {
		const alreadyOwned = owned.get(sourcePartName)
		if (alreadyOwned !== undefined) return alreadyOwned
	} else {
		const existing = ctx.registry.get(sourcePartName)
		const selectedDest = ctx.selection?.destinations.get(sourcePartName)
		// Registry hit without a selection plan: plain idempotence. With a plan, a
		// hit on the *selected* destination means this batch already walked the page
		// (register-before-recurse makes mutually-linked selected pages terminate);
		// a hit on any OTHER partname is either a shared non-slide dependency from
		// an earlier traversal or a page a previous import brought across — both are
		// reused rather than duplicated. Only a selected page whose registered
		// partname differs (imported by a previous call) is re-materialized fresh:
		// each batch request owns exactly one output page.
		if (existing && (selectedDest === undefined || existing === selectedDest)) return existing
	}

	const sourcePart = ctx.source.part(sourcePartName)
	if (!sourcePart)
		throw new PackageReadError('package/part-missing', `importSlide: source package has no part ${sourcePartName}`)

	const newPartName = owned
		? ctx.dest.opc.reservePartNameLike(sourcePartName)
		: (ctx.selection?.destinations.get(sourcePartName) ?? ctx.dest.opc.reservePartNameLike(sourcePartName))
	// A selected page's part was already materialized by the batch allocator.
	if (!ctx.dest.opc.part(newPartName)) ctx.dest.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
	// Record before recursing so the master↔layout cycle terminates.
	if (owned) owned.set(sourcePartName, newPartName)
	else ctx.registry.set(sourcePartName, newPartName)

	// A page opens an ownership scope for everything under it; parts inside one
	// stay inside it. See `page-owned.ts` for what that scope covers and why.
	const scope = owned ?? (ctx.selection?.destinations.has(sourcePartName) ? newOwnedScope() : undefined)

	const isMaster = sourcePart.contentType === SLIDE_MASTER_CONTENT_TYPE
	const sourceRels = ctx.source.relationshipsFor(sourcePartName)
	const targetRels = ctx.dest.opc.relationshipsFor(newPartName)
	for (const rel of sourceRels) {
		// Notes pull in a notesMaster + its own theme; an imported slide does not need them.
		if (rel.type === NOTES_SLIDE_REL) continue
		// Lean master: skip its layout rels; copied layouts re-link themselves.
		if (isMaster && rel.type === SLIDE_LAYOUT_REL) continue
		if (rel.targetMode === 'External') {
			targetRels.addWithId(rel.id, rel.type, rel.target, 'External')
			continue
		}
		// A batch import's slide→slide rule is enforced by checkSelectionCopyable
		// before this traversal starts, so the recursion below cannot reach an
		// unselected page: by here every target is selected, already copied, or
		// not a slide at all.
		const newTargetPartName = copyPart(
			ctx,
			sourceRels.resolveTarget(rel.id),
			isSharedByPageCopies(rel.type) ? undefined : scope
		)
		targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTargetPartName))
	}

	if (isMaster) {
		clearLayoutIdList(ctx.dest, newPartName)
		// Register the copied master in presentation.xml. Without a
		// `p:sldMasterId` entry (and a presentation→master relationship) the
		// master is inert: PowerPoint/LibreOffice ignore its background and shape
		// tree, so a `copy`-imported slide whose look lives on its master (a
		// cover/closer) renders blank. Idempotent, so masters shared across
		// repeated imports are registered exactly once.
		registerMaster(ctx.dest, newPartName)
	}
	if (sourcePart.contentType === SLIDE_LAYOUT_CONTENT_TYPE) {
		linkLayoutIntoMaster(ctx, sourceRels, newPartName)
	}

	return newPartName
}

/**
 * Which of a batch's selected pages are also carrying their speaker notes across,
 * and whether the notes subgraph will pull a `notesMaster` with it. Handed to
 * {@link checkSelectionCopyable} so the dry run walks what `carryNotes` will walk.
 */
export interface NotesSelection {
	/** Source partnames of the selected pages whose request asked for `importNotes`. */
	readonly pages: ReadonlySet<string>
	/**
	 * Whether a source `notesMaster` would be copied. False once the destination
	 * has one of its own, since `p:notesMasterIdLst` is `0..1` and `carryNotes`
	 * then binds to the destination's master instead of copying the source's.
	 */
	readonly copyMaster: boolean
}

const NO_NOTES: NotesSelection = { pages: new Set(), copyMaster: false }

/**
 * Dry-run {@link copyPart} over one source's batch selection, reading only the
 * *source* package: walk the graph the copy will walk, under the same
 * relationship-skipping rules, and throw what the copy would throw — a source
 * part that is missing, XML that will not parse, or a `slide → slide` link that
 * escapes the selection.
 *
 * This is what makes a rejected batch leave the destination byte-identical.
 * `importSlides` reserves partnames and mutates the destination only after this
 * passes, at which point the copy has no reachable failure left; a check that
 * lived inside the traversal instead would fire with parts already added, a
 * master already registered in `presentation.xml`, and no way back.
 *
 * Keep this in step with `copyPart`: the two must skip the same relationships
 * and stop at the same registry hits, or the guarantee is only as good as the
 * drift between them. With `notes` the same obligation extends to `carryNotes`,
 * which runs after the copy and is the batch's other way to reach a source part.
 *
 * @param source     the package the pages are coming from
 * @param registry   the copy registry for that source (parts already in `dest`)
 * @param selected   source partnames of the pages this batch selected
 * @param notes      the pages of `selected` whose notes travel too; none by default
 */
export function checkSelectionCopyable(
	source: OpcPackage,
	registry: ReadonlyMap<string, string>,
	selected: ReadonlySet<string>,
	notes: NotesSelection = NO_NOTES
): void {
	const visited = new Set<string>()

	const walk = (partName: string): void => {
		if (visited.has(partName)) return
		// `copyPart` reuses an already-copied part rather than recursing into it,
		// and its subgraph was validated when that copy happened. A *selected*
		// page is the exception the batch re-materializes, so it is always walked.
		if (registry.has(partName) && !selected.has(partName)) return
		visited.add(partName)

		const part = source.part(partName)
		if (!part)
			throw new PackageReadError('package/part-missing', `importSlides: source package has no part ${partName}`)

		const isMaster = part.contentType === SLIDE_MASTER_CONTENT_TYPE
		const isNotesSlide = part.contentType === NOTES_SLIDE_CONTENT_TYPE
		// The copy re-parses a master/layout to rebuild its layout id list; force
		// the same parse here so unparseable XML fails before anything moves.
		if (isMaster || part.contentType === SLIDE_LAYOUT_CONTENT_TYPE) void part.dom

		const rels = source.relationshipsFor(partName)
		for (const rel of rels) {
			// `copyPart` always drops the notes rel; `carryNotes` picks it up
			// afterwards for the pages that asked, so those subgraphs are walked here
			// and no other.
			if (rel.type === NOTES_SLIDE_REL) {
				if (notes.pages.has(partName)) walk(rels.resolveTarget(rel.id))
				continue
			}
			if (isNotesSlide) {
				// The notes' back-rel to the page it annotates is repointed at the new
				// slide, not copied — walking it would re-enter the page graph.
				if (rel.type === SLIDE_REL) continue
				// A deck already holding a notesMaster keeps it; the source's is then
				// never read, so a dry run that walked it would reject a batch the copy
				// would have accepted.
				if (rel.type === NOTES_MASTER_REL && !notes.copyMaster) continue
			}
			if (isMaster && rel.type === SLIDE_LAYOUT_REL) continue
			if (rel.targetMode === 'External') continue
			const targetPartName = rels.resolveTarget(rel.id)
			// A jump link must land on another selected page, or on one an earlier
			// import from this source already brought across. Anything else would
			// drag a page nobody asked for into the deck, or strand the link.
			if (rel.type === SLIDE_REL && !selected.has(targetPartName) && !registry.has(targetPartName)) {
				throw new InvalidOptionError(
					'import/unresolved-slide-link',
					`importSlides: source slide ${partName} links to ${targetPartName}, which is not among the selected imported pages`
				)
			}
			walk(targetPartName)
		}
	}

	for (const partName of selected) walk(partName)
}

/**
 * Wire a just-copied layout into its (already-copied) master. Resolves which
 * destination master that is by running the layout's *source* master rel through
 * the copy registry, then hands the destination-side wiring to
 * {@link addLayoutToMaster}. Called once per copied layout, so the master
 * accumulates exactly the imported layouts.
 */
function linkLayoutIntoMaster(
	ctx: ImportContext,
	layoutSourceRels: ReturnType<OpcPackage['relationshipsFor']>,
	layoutPartName: string
): void {
	const masterRel = layoutSourceRels.byType(SLIDE_MASTER_REL)[0]
	if (!masterRel) return
	const masterPartName = ctx.registry.get(layoutSourceRels.resolveTarget(masterRel.id))
	if (!masterPartName) return
	addLayoutToMaster(ctx.dest, masterPartName, layoutPartName)
}
