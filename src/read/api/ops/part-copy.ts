/**
 * The part-copy traversal: pulling a part -- and, recursively, every internal part
 * it references -- out of a source package into a destination deck, rewriting
 * relationship targets to the freshly-allocated partnames as it goes.
 *
 * This is the source-side half of the import machinery; the destination-side
 * bookkeeping it hands off to lives in `master-registry.ts`. It takes an
 * {@link ImportContext} rather than a `Presentation`, so it stays independent of
 * the class that calls it.
 *
 * The same traversal is every import's dry run. With a {@link CopyPlan} in its
 * context it writes nothing and throws what the copy would throw, so an import
 * runs it that way before anything in the deck moves.
 */

import type { OpcPackage } from '../../opc/package.js'
import type { Part } from '../../opc/part.js'
import type { Relationship, Relationships } from '../../opc/relationships.js'
import { relativePartName } from '../../opc/partnames.js'
import type { DeckTarget } from './deck-target.js'
import { addLayoutToMaster, clearLayoutIdList, registerMaster } from './master-registry.js'
import {
	SLIDE_LAYOUT_CONTENT_TYPE,
	SLIDE_CONTENT_TYPE,
	SLIDE_MASTER_CONTENT_TYPE,
	SLIDE_MASTER_REL,
	SLIDE_REL,
} from '../../../ooxml/rel-types.js'
import { copyTraversalStep } from './copy-traversal.js'
import { isSharedByPageCopies } from './page-owned.js'
import { destinationAlreadyHolds } from './part-reuse.js'
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
	/**
	 * Set when the traversal is planning rather than copying: every write goes to the
	 * plan instead of `dest`, and `registry` is the plan's own copy of the real one.
	 */
	readonly plan?: CopyPlan
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
interface SelectionPlan {
	/** Source slide partname → the destination partname reserved for it. */
	readonly destinations: ReadonlyMap<string, string>
}

/**
 * What a copy writes through: the destination package, or a {@link CopyPlan} standing in for it.
 * `OpcPackage` satisfies it.
 */
export interface CopyTarget {
	reservePartNameLike(templatePartName: string): string
	part(partName: string): unknown
	addPart(partName: string, contentType: string, bytes: Uint8Array): unknown
	relationshipsFor(partName: string): Pick<Relationships, 'add' | 'addWithId'>
}

/** Where a traversal writes: its plan when it is planning, its destination package otherwise. */
export function copyTarget(ctx: ImportContext): CopyTarget {
	return ctx.plan ?? ctx.dest.opc
}

/** One part a {@link CopyPlan} would add. */
export interface PlannedPart {
	readonly partName: string
	readonly contentType: string
}

/** One relationship a {@link CopyPlan} would add; `id` is absent where the copy allocates one. */
export interface PlannedRel {
	readonly from: string
	readonly id?: string
	readonly type: string
	readonly target: string
	readonly targetMode?: 'Internal' | 'External'
}

/**
 * An import's copy, run without writing: the parts and relationships it would add, and every error
 * it would throw, found before anything in the destination moves.
 *
 * This is what lets a refused import leave the deck byte-identical. The copy changes the deck as it
 * goes — parts added, a master registered in `presentation.xml`, a layout linked into its master —
 * so a failure it met halfway would leave all of that behind. An import therefore runs the same
 * traversal twice: first with a plan in its context, which throws what the copy would and writes
 * nothing, then for real. The dry run used to be a separate walk of the source, transcribed from the
 * copy under a comment asking readers to keep the two in step, and it had drifted: it walked a
 * rebinding import's source layout chain, which the rebind never reads, and every source's notes
 * master in a batch, where the copy reads only the first one it installs.
 *
 * A plan stands in for the destination package as a {@link CopyTarget}. It hands out the partnames
 * the copy will be given, keeps its own copy of each source's registry so a part it planned is found
 * again as the copy would find it, and remembers the notes master it would install in a deck that has
 * none. Destination bookkeeping that adds no part — registering a master, linking a layout into it —
 * is not recorded.
 */
export class CopyPlan implements CopyTarget {
	/** The parts the copy would add, in the order it would add them. */
	readonly parts: PlannedPart[] = []
	/** The relationships the copy would add to the parts it copies. */
	readonly rels: PlannedRel[] = []
	/** The notes master the copy would install in a deck that has none; later notes bind to it. */
	notesMaster: string | undefined
	readonly #dest: DeckTarget
	/** Every partname handed out, added or not yet added. */
	readonly #names = new Set<string>()
	readonly #added = new Set<string>()
	readonly #registries = new Map<OpcPackage, Map<string, string>>()

	/**
	 * @param dest      the deck the import copies into
	 * @param api       the public method planning, which opens every error the plan throws
	 * @param linkable  the pages a jump link may land on when the traversal has no selection of its
	 *                  own: the page a rebinding import builds itself
	 */
	constructor(
		dest: DeckTarget,
		readonly api: string,
		readonly linkable?: ReadonlySet<string>
	) {
		this.#dest = dest
	}

	/** This plan's context for an import out of `ctx.source`, over the plan's copy of its registry. */
	contextFor(ctx: ImportContext): ImportContext {
		let registry = this.#registries.get(ctx.source)
		if (!registry) {
			registry = new Map(ctx.registry)
			this.#registries.set(ctx.source, registry)
		}
		return { dest: ctx.dest, source: ctx.source, registry, plan: this }
	}

	reservePartNameLike(templatePartName: string): string {
		const partName = this.#dest.opc.reservePartNameLike(templatePartName, this.#names)
		this.#names.add(partName)
		return partName
	}

	part(partName: string): unknown {
		return this.#added.has(partName) || this.#dest.opc.part(partName)
	}

	addPart(partName: string, contentType: string): void {
		this.#names.add(partName)
		this.#added.add(partName)
		this.parts.push({ partName, contentType })
	}

	relationshipsFor(from: string): Pick<Relationships, 'add' | 'addWithId'> {
		const record = (rel: PlannedRel): Relationship => {
			this.rels.push(rel)
			const { id = `planned${this.rels.length}`, type, target, targetMode } = rel
			return { id, type, target, ...(targetMode ? { targetMode } : {}) }
		}
		return {
			add: (type, target, targetMode) => record({ from, type, target, ...(targetMode ? { targetMode } : {}) }),
			addWithId: (id, type, target, targetMode) =>
				record({ from, id, type, target, ...(targetMode ? { targetMode } : {}) }),
		}
	}
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
	const destinations = new Map([[sourcePartName, copyTarget(ctx).reservePartNameLike(sourcePartName)]])
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
 * a `slide → slide` link may not leave the selection is enforced as each link is
 * followed (see {@link rebuildRels}). Imports run this traversal as a
 * {@link CopyPlan} first, which throws that and every other failure the copy
 * could meet — so once copying starts there is nothing left to throw, and a
 * rejected batch never leaves a half-copied deck behind.
 *
 * Idempotence stops at the page's own parts. Reaching a page opens an
 * {@link OwnedScope}, and everything under it that {@link isSharedByPageCopies}
 * does not clear for sharing — a chart, a diagram, an OLE embedding, each with
 * its own subtree — is copied into that scope instead of the registry, so the
 * next copy of the page gets parts of its own. Pass `owned` to open the scope
 * from outside, as the rebinding import paths do for the page they build
 * themselves.
 */
export function copyPart(
	ctx: ImportContext,
	sourcePartName: string,
	owned?: OwnedScope,
	allowReuse: boolean = false
): string {
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

	// The destination may already hold this exact part: a deck opened with
	// `fromTemplate` from the *same file* carries the source's own layouts, master and
	// theme, under their own partnames and byte-identical. Binding to what is there
	// beats copying it in again under a fresh name, which is what grew a duplicate
	// layout-gallery entry per imported slide. Never inside a page's ownership scope,
	// and never for a page the caller is materializing: those must be parts of their own.
	if (allowReuse && !owned && !ctx.selection?.destinations.has(sourcePartName)) {
		if (destinationAlreadyHolds(ctx.dest, ctx.source, sourcePartName)) {
			ctx.registry.set(sourcePartName, sourcePartName)
			return sourcePartName
		}
	}

	const sourcePart = ctx.source.part(sourcePartName)
	if (!sourcePart)
		throw new PackageReadError(
			'package/part-missing',
			`${ctx.plan?.api ?? 'importSlide'}: source package has no part ${sourcePartName}`
		)

	const target = copyTarget(ctx)
	const newPartName = owned
		? target.reservePartNameLike(sourcePartName)
		: (ctx.selection?.destinations.get(sourcePartName) ?? target.reservePartNameLike(sourcePartName))
	// A selected page's part was already materialized by the batch allocator.
	if (!target.part(newPartName)) target.addPart(newPartName, sourcePart.contentType, sourcePart.serialize())
	// Record before recursing so the master↔layout cycle terminates.
	if (owned) owned.set(sourcePartName, newPartName)
	else ctx.registry.set(sourcePartName, newPartName)

	// A page opens an ownership scope for everything under it; parts inside one
	// stay inside it. See `page-owned.ts` for what that scope covers and why.
	const scope = owned ?? (ctx.selection?.destinations.has(sourcePartName) ? newOwnedScope() : undefined)

	rebuildRels(ctx, {
		source: sourcePart,
		newPartName,
		owned: scope,
		// Reuse decides at the page boundary and never below it. A part reached from
		// the page can be answered with the destination's own identical copy; a part
		// reached from something this import *copied* is copied too, so a copied
		// subgraph stays self-contained rather than half-linking into deck chrome
		// that merely happens to match.
		allowReuse: sourcePart.contentType === SLIDE_CONTENT_TYPE,
	})

	if (ctx.plan) {
		// A plan registers and links nothing. The dry run has always parsed a master or
		// layout it would copy, as the copy re-parses a master to empty its layout id
		// list, so XML that will not parse is refused before anything moves.
		if (sourcePart.contentType === SLIDE_MASTER_CONTENT_TYPE || sourcePart.contentType === SLIDE_LAYOUT_CONTENT_TYPE)
			void sourcePart.dom
		return newPartName
	}

	if (sourcePart.contentType === SLIDE_MASTER_CONTENT_TYPE) {
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
		linkLayoutIntoMaster(ctx, ctx.source.relationshipsFor(sourcePartName), newPartName)
	}

	return newPartName
}

/**
 * What a caller decides about one relationship before the shared rule is asked: leave it out
 * (`skip`), carry it exactly as the source wrote it (`keep`), or point it at a destination part of
 * the caller's choosing. `undefined` hands it to the rule.
 */
export type RelOverride = 'skip' | 'keep' | { target: string } | undefined

/** What {@link rebuildRels} rebuilds, and the caller's special cases. */
export interface RebuildRelsOptions {
	/** The source part whose relationships are copied. */
	source: Part
	/** The destination part they are rebuilt on. */
	newPartName: string
	/** The ownership scope of the page copy this part belongs to, if it belongs to one. */
	owned?: OwnedScope | undefined
	/** Whether a followed target the destination already holds may be bound to rather than copied. */
	allowReuse: boolean
	/** Sees every relationship first; see {@link RelOverride}. */
	override?: (rel: Relationship) => RelOverride
}

/**
 * Rebuild a copied part's relationships from its source's: the one loop every import that copies a
 * part runs.
 *
 * Each relationship keeps its source id, so the copied body's `r:id` and `r:embed` references stay
 * valid without the body being rewritten; only targets change. A caller's `override` sees each
 * relationship first and may skip it or name its target: the rebinding import points the layout at
 * this deck's, and a notes copy points its slide and its notes master at theirs. Everything else
 * follows {@link copyTraversalStep}: skipped, carried as an external link, or followed into the
 * part {@link copyPart} makes for it, inside `owned` unless {@link isSharedByPageCopies} says the
 * page may share it.
 *
 * It is also where a jump link is held to the pages the import brings. A `slide` relationship the
 * rule would follow must land on a page of `ctx.selection` (or, for a traversal with no selection of
 * its own, a page its plan names), or on one an earlier import from this source already brought
 * across. Anything else would drag a page nobody asked for into the deck, or strand the link, so it
 * is refused; imports plan first, so the refusal comes before anything moves.
 */
export function rebuildRels(ctx: ImportContext, options: RebuildRelsOptions): void {
	const { source, newPartName, owned, allowReuse, override } = options
	const sourceRels = ctx.source.relationshipsFor(source.partName)
	const targetRels = copyTarget(ctx).relationshipsFor(newPartName)
	const linkable = ctx.selection?.destinations ?? ctx.plan?.linkable
	for (const rel of sourceRels) {
		const decided = override?.(rel)
		if (decided === 'skip') continue
		if (decided === 'keep') {
			targetRels.addWithId(rel.id, rel.type, rel.target, rel.targetMode)
			continue
		}
		if (decided) {
			targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, decided.target))
			continue
		}
		const step = copyTraversalStep(source, rel)
		if (step === 'skip') continue
		if (step === 'external') {
			targetRels.addWithId(rel.id, rel.type, rel.target, 'External')
			continue
		}
		const targetPartName = sourceRels.resolveTarget(rel.id)
		if (rel.type === SLIDE_REL && linkable && !linkable.has(targetPartName) && !ctx.registry.has(targetPartName)) {
			throw new InvalidOptionError(
				'import/unresolved-slide-link',
				`${ctx.plan?.api ?? 'importSlides'}: source slide ${source.partName} links to ${targetPartName}, which is not among the imported pages`
			)
		}
		const target = copyPart(ctx, targetPartName, isSharedByPageCopies(rel.type) ? undefined : owned, allowReuse)
		targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, target))
	}
}

/**
 * Give a page cloned within its own package the source page's relationships, with its own copy of
 * every part the source page owned.
 *
 * The page's shared targets are carried as the source wrote them. Each owned one is repointed at a
 * fresh copy, and the subtree under it is copied the same way, so a chart's workbook and
 * user-shapes drawing come along while the image inside that drawing stays shared. A relationship
 * *back* to the source page — a notes slide names the slide it annotates — is repointed at the
 * clone, which is what makes the copied notes belong to it.
 *
 * Unlike an import, a clone keeps its page's notes, and nothing here consults the copy registry or
 * the reuse check: every relationship is decided by an override, so {@link copyTraversalStep} is
 * never asked. A dangling relationship is left dangling rather than made to throw: cloning a
 * damaged deck is not this function's problem to discover.
 *
 * @param dest            the deck both pages live in
 * @param sourcePart      the page that was cloned
 * @param clonePartName   partname of the clone, which has no relationships yet
 */
export function rebuildClonedPageRels(dest: DeckTarget, sourcePart: Part, clonePartName: string): void {
	// Within one package the source is the destination.
	const ctx: ImportContext = { dest, source: dest.opc, registry: new Map() }
	// Seeded with the page itself, so a back-reference to it lands on the clone.
	const copies = new Map<string, string>([[sourcePart.partName, clonePartName]])
	const sourceRels = dest.opc.relationshipsFor(sourcePart.partName)
	rebuildRels(ctx, {
		source: sourcePart,
		newPartName: clonePartName,
		allowReuse: false,
		override: (rel) => {
			if (rel.targetMode === 'External' || isSharedByPageCopies(rel.type)) return 'keep'
			const target = sourceRels.resolveTarget(rel.id)
			const fresh = copyOwnedSubtree(ctx, target, copies)
			return fresh === target ? 'keep' : { target: fresh }
		},
	})
}

/**
 * Copy `partName` and, recursively, every part it owns, into fresh partnames in the same package.
 * `copies` dedupes within the one page copy, so a part two of the page's relationships reach is
 * copied once here even though the next page copy gets its own. Returns the copy's partname, or
 * `partName` unchanged when there is no such part to copy.
 */
function copyOwnedSubtree(ctx: ImportContext, partName: string, copies: Map<string, string>): string {
	const already = copies.get(partName)
	if (already !== undefined) return already
	const opc = ctx.dest.opc
	const part = opc.part(partName)
	if (!part) return partName

	const fresh = opc.reservePartNameLike(partName)
	opc.addPart(fresh, part.contentType, part.serialize())
	// Record before recursing, so a cycle (a notes slide naming its slide) terminates.
	copies.set(partName, fresh)

	const sourceRels = opc.relationshipsFor(partName)
	rebuildRels(ctx, {
		source: part,
		newPartName: fresh,
		allowReuse: false,
		override: (rel) => {
			if (rel.targetMode === 'External') return 'keep'
			const target = sourceRels.resolveTarget(rel.id)
			// A shared target is still routed through `copies`: that is how the notes
			// slide's `slide` relationship finds the clone instead of the original.
			if (isSharedByPageCopies(rel.type)) return { target: copies.get(target) ?? target }
			return { target: copyOwnedSubtree(ctx, target, copies) }
		},
	})
	return fresh
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
