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
import { NOTES_SLIDE_REL, SLIDE_LAYOUT_REL, SLIDE_MASTER_REL, SLIDE_REL } from '../../../ooxml/rel-types.js'
import { InvalidOptionError, PackageReadError } from '../../../errors.js'

const SLIDE_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const SLIDE_LAYOUT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'

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
	 * The batch-selection plan when copying for {@link Presentation.importSlides}:
	 * source slide partname → the destination partname pre-allocated for it. When
	 * present, `slide → slide` relationships resolve only within this set — an
	 * imported page may link to another *selected* page (rewritten to its fresh
	 * partname) but must not drag an unselected source page across as a dependency.
	 */
	readonly selection?: SelectionPlan
}

/** Which pages of one source package a batch import selected, and where each is headed. */
export interface SelectionPlan {
	/** Source slide partname → the destination partname reserved for it. */
	readonly destinations: ReadonlyMap<string, string>
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
 */
export function copyPart(ctx: ImportContext, sourcePartName: string): string {
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

	const sourcePart = ctx.source.part(sourcePartName)
	if (!sourcePart)
		throw new PackageReadError('package/part-missing', `importSlide: source package has no part ${sourcePartName}`)

	const newPartName =
		ctx.selection?.destinations.get(sourcePartName) ?? ctx.dest.opc.reservePartNameLike(sourcePartName)
	// A selected page's part was already materialized by the batch allocator.
	if (!ctx.dest.opc.part(newPartName)) ctx.dest.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
	// Record before recursing so the master↔layout cycle terminates.
	ctx.registry.set(sourcePartName, newPartName)

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
		const newTargetPartName = copyPart(ctx, sourceRels.resolveTarget(rel.id))
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
 * drift between them.
 *
 * @param source     the package the pages are coming from
 * @param registry   the copy registry for that source (parts already in `dest`)
 * @param selected   source partnames of the pages this batch selected
 */
export function checkSelectionCopyable(
	source: OpcPackage,
	registry: ReadonlyMap<string, string>,
	selected: ReadonlySet<string>
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
		// The copy re-parses a master/layout to rebuild its layout id list; force
		// the same parse here so unparseable XML fails before anything moves.
		if (isMaster || part.contentType === SLIDE_LAYOUT_CONTENT_TYPE) void part.dom

		const rels = source.relationshipsFor(partName)
		for (const rel of rels) {
			if (rel.type === NOTES_SLIDE_REL) continue
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
