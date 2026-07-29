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

import type { OpcPackage } from '../opc/package.js'
import { relativePartName } from '../opc/partnames.js'
import type { DeckTarget } from './deck-target.js'
import { addLayoutToMaster, clearLayoutIdList, registerMaster } from './master-registry.js'
import { PackageReadError } from '../../errors.js'

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'

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
 */
export function copyPart(ctx: ImportContext, sourcePartName: string): string {
	const existing = ctx.registry.get(sourcePartName)
	if (existing) return existing

	const sourcePart = ctx.source.part(sourcePartName)
	if (!sourcePart)
		throw new PackageReadError('package/part-missing', `importSlide: source package has no part ${sourcePartName}`)

	const newPartName = ctx.dest.opc.reservePartNameLike(sourcePartName)
	ctx.dest.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
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
