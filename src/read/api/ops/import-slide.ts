/**
 * The three ways a slide can be brought across from another deck.
 *
 * {@link Presentation.importSlide} picks one by `options.theme`, and the difference between
 * them is entirely about what happens to the slide's *look*:
 *
 * - **`copy`** — not here. It copies the source theme subgraph wholesale via `copyPart`, so
 *   the slide keeps its look by keeping its own theme; nothing needs rewriting.
 * - **`preserve`** ({@link importSlidePreserve}) — rebind to this deck's master, then bake
 *   the source theme into the slide XML. The slide looks the same as it did; it just no
 *   longer depends on the theme it came from.
 * - **`restyle`** ({@link importSlideRestyle}) — rebind and bake *nothing*, so the slide's
 *   symbolic theme references re-resolve against the destination theme and it re-brands.
 *
 * The two share {@link importSlideRebind} and are exact inverses on top of it: preserve
 * resolves every reference it can, restyle resolves none. Any baking added to the shared
 * rebind would break restyle by pinning the slide to its source look, which is why the
 * rebind carries no theme handling of its own — not even the inherited background.
 *
 * These were private methods on `Presentation`, where they made the read model's own surface
 * (open a deck, walk slides, read shapes) hard to see: the import machinery was roughly two
 * thirds of the class. They live here as free functions taking the destination deck, matching
 * the other `ops/` modules.
 */

import { ownerDocumentOf, type Element } from '../../oxml/dom.js'
import type { Part } from '../../opc/part.js'
import { relativePartName } from '../../opc/partnames.js'
import { PackageReadError, InvalidOptionError } from '../../../errors.js'
import { flattenSlide, remapLiteralColors, restyleSlide } from './flatten.js'
import { resolveSlideThemeParts } from '../theme-context.js'
import { carriedDecorations, firstShapeChild, spTreeOf } from '../../oxml/slide-dom.js'
import { copySourceTableStyles } from './table-styles.js'
import { copyPart, newOwnedScope, type ImportContext, type OwnedScope } from './part-copy.js'
import { isSharedByPageCopies } from './page-owned.js'
import type { Presentation } from '../presentation.js'
import type { Slide } from '../slide.js'
import { NOTES_SLIDE_REL, SLIDE_LAYOUT_REL, SLIDE_MASTER_REL } from '../../../ooxml/rel-types.js'
import { resolveSingleRel } from './part-index.js'
import { rewriteCarriedRels } from './carried-rels.js'
import { sourceFlattenContext } from './flatten-context.js'

/** What {@link importSlideRebind} hands back for the caller's mode-specific pass. */
interface RebindResult {
	newPartName: string
	slideRoot: Element
	newPart: Part
}

/**
 * Import a slide in `preserve` mode: rebind it to this deck's master/layout, then flatten its
 * source theme into the slide XML (scheme colours + style-matrix fills baked to literals).
 *
 * The flatten context is gathered from the *source* subgraph, so it can be read before or
 * after the rebind; the rebind injects any carried decorations before we flatten, so a single
 * sweep resolves the theme references on the slide's own content and on the carried
 * decorations together.
 * @param {Presentation} dest - the destination deck
 * @param {ImportContext} ctx - the open import out of the source package
 * @param {Presentation} source - the source deck
 * @param {Slide} sourceSlide - the slide being imported
 * @param {boolean} carryGraphics - bake the source master/layout decorations onto the slide
 * @return {string} partname of the new slide part in the destination package
 */
export function importSlidePreserve(
	dest: Presentation,
	ctx: ImportContext,
	source: Presentation,
	sourceSlide: Slide,
	carryGraphics: boolean
): string {
	const flattenCtx = sourceFlattenContext(source.opc, sourceSlide.partName)
	const { newPartName, slideRoot, newPart } = importSlideRebind(dest, ctx, source, sourceSlide, carryGraphics)
	flattenSlide(slideRoot, flattenCtx)
	newPart.markDirty()
	return newPartName
}

/**
 * Import a slide in `restyle` mode: rebind it to this deck's master/layout and then
 * {@link restyleSlide} it — drop its colour map override but bake *nothing*, so its symbolic
 * theme references re-resolve against the destination theme and the slide re-brands.
 *
 * The deliberate inverse of `preserve`: no flatten, no inherited-background bake, no
 * placeholder colour/size/geometry bake — every one of those would pin the slide to its
 * source look, the opposite of re-branding. Carried decorations are left symbolic too, so
 * they re-brand along with the slide.
 *
 * With `remapLiterals` it additionally force-remaps the slide's source-theme literal colours
 * back to symbolic scheme colours and copies any referenced source table style into this deck
 * — the two things plain `restyle` cannot re-brand.
 * @param {Presentation} dest - the destination deck
 * @param {ImportContext} ctx - the open import out of the source package
 * @param {Presentation} source - the source deck
 * @param {Slide} sourceSlide - the slide being imported
 * @param {boolean} carryGraphics - bake the source master/layout decorations onto the slide
 * @param {boolean} remapLiterals - additionally remap source-theme literal colours to scheme tokens
 * @return {string} partname of the new slide part in the destination package
 */
export function importSlideRestyle(
	dest: Presentation,
	ctx: ImportContext,
	source: Presentation,
	sourceSlide: Slide,
	carryGraphics: boolean,
	remapLiterals: boolean
): string {
	const { newPartName, slideRoot, newPart } = importSlideRebind(dest, ctx, source, sourceSlide, carryGraphics)
	restyleSlide(slideRoot)
	if (remapLiterals) {
		// The source colour context (slot ↔ RGB ↔ token) the literals are matched against.
		const parts = resolveSlideThemeParts(source.opc, sourceSlide.partName)
		remapLiteralColors(slideRoot, { clrMap: parts.clrMap, clrScheme: parts.clrScheme })
		copySourceTableStyles(dest, source.opc, slideRoot)
	}
	newPart.markDirty()
	return newPartName
}

/**
 * The rebind shared by `preserve` and `restyle`: copy the slide bytes into a fresh part,
 * rebuild its relationships (drop notes, repoint the `slideLayout` rel at this deck's
 * existing layout, copy every other internal target — media/charts — and pass externals
 * through), and optionally bake the source master/layout decorations onto the slide. Returns
 * the new part, its name, and its live root element for the caller's mode-specific pass.
 *
 * This carries *no* theme baking of its own — not even the inherited background.
 * `preserve` adds that via {@link flattenSlide}'s context; `restyle` must not, so the
 * background stays symbolic and re-brands.
 * @param {Presentation} dest - the destination deck
 * @param {ImportContext} ctx - the open import out of the source package
 * @param {Presentation} source - the source deck
 * @param {Slide} sourceSlide - the slide being imported
 * @param {boolean} carryGraphics - bake the source master/layout decorations onto the slide
 * @return {RebindResult} the new part, its partname, and its live root element
 */
export function importSlideRebind(
	dest: Presentation,
	ctx: ImportContext,
	source: Presentation,
	sourceSlide: Slide,
	carryGraphics: boolean
): RebindResult {
	const destLayout = destinationLayoutPartName(dest)

	// Copy the slide bytes into a fresh partname; we then mutate that copy's DOM
	// (a distinct document, so the source package is never touched).
	const sourcePart = source.opc.part(sourceSlide.partName)
	if (!sourcePart)
		throw new PackageReadError(
			'package/part-missing',
			`importSlide: source package has no part ${sourceSlide.partName}`
		)
	const newPartName = dest.opc.reservePartNameLike(sourceSlide.partName)
	const newPart = dest.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
	const slideRoot = newPart.dom.documentElement
	if (!slideRoot)
		throw new PackageReadError('package/part-has-no-root', `Imported slide ${newPartName} has no root element`)

	// Rebuild the slide's relationships: drop notes, repoint slideLayout at the
	// destination layout, and copy every other internal target (media/charts).
	const sourceRels = source.opc.relationshipsFor(sourceSlide.partName)
	const targetRels = dest.opc.relationshipsFor(newPartName)
	// This page is built here rather than by `copyPart`, so its ownership scope is
	// opened here too: the chart or diagram under a page rebound twice must not be
	// the same part twice (see `page-owned.ts`).
	const owned = newOwnedScope()
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
		const newTarget = copyPart(
			ctx,
			sourceRels.resolveTarget(rel.id),
			isSharedByPageCopies(rel.type) ? undefined : owned
		)
		targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTarget))
	}

	// Optionally bake the source master/layout decorations (logos, accent shapes) onto the
	// slide behind its own content. Done after the slide's own rels are in place (so carried
	// media get fresh, non-colliding ids) but before the caller's flatten/restyle pass acts
	// on the carried shapes.
	if (carryGraphics) carryMasterGraphics(dest, ctx, slideRoot, newPartName, sourceSlide.partName, owned)

	return { newPartName, slideRoot, newPart }
}

/**
 * Bake the source `slideLayout`/`slideMaster` shape-tree decorations onto the imported slide
 * (the `carryMasterGraphics` path). Every shape on those trees *except* placeholders is
 * deep-copied into the slide's `p:spTree` ahead of its own content — master decorations
 * first, then layout, then the slide's shapes — so document (z-)order keeps the master
 * furthest back. Each decoration's media and other relationship targets are copied into this
 * package and its `r:embed`/`r:id`/… references rewritten to fresh slide-local ids. The
 * injected shapes are left for the caller's {@link flattenSlide} pass to resolve any theme
 * references they carry.
 * @param {Presentation} dest - the destination deck
 * @param {ImportContext} ctx - the open import out of the source package
 * @param {Element} slideRoot - root element of the new slide part (mutated in place)
 * @param {string} newPartName - partname of the new slide part
 * @param {string} slidePartName - partname of the source slide, for resolving its layout/master
 * @param {OwnedScope} owned - the page's ownership scope, so a decoration's own parts land in it
 */
function carryMasterGraphics(
	dest: Presentation,
	ctx: ImportContext,
	slideRoot: Element,
	newPartName: string,
	slidePartName: string,
	owned: OwnedScope
): void {
	const sourceOpc = ctx.source
	const layoutPartName = resolveSingleRel(sourceOpc, slidePartName, SLIDE_LAYOUT_REL)
	const masterPartName = layoutPartName ? resolveSingleRel(sourceOpc, layoutPartName, SLIDE_MASTER_REL) : null
	const spTree = spTreeOf(slideRoot)
	if (!spTree) return

	const doc = ownerDocumentOf(slideRoot)
	const slideRels = dest.opc.relationshipsFor(newPartName)
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
			rewriteCarriedRels(imported, ctx, sourceRels, newPartName, slideRels, relIdMap, owned)
			spTree.insertBefore(imported, anchor)
		}
	}
}

/**
 * The partname of the layout this deck's slides should attach to in `preserve` mode: the
 * first layout of the first slide master. Throws when the deck has no master/layout to attach
 * to (a deck ts-pptx always provides).
 * @param {Presentation} dest - the destination deck
 * @return {string} partname of the destination layout
 */
function destinationLayoutPartName(dest: Presentation): string {
	const presRels = dest.opc.relationshipsFor(dest.presentationPart.partName)
	const masterRel = presRels.byType(SLIDE_MASTER_REL)[0]
	if (!masterRel)
		throw new InvalidOptionError(
			'import/destination-missing-master',
			'importSlide preserve mode requires a slide master in the destination deck'
		)
	const masterPartName = presRels.resolveTarget(masterRel.id)
	const masterRels = dest.opc.relationshipsFor(masterPartName)
	const layoutRel = masterRels.byType(SLIDE_LAYOUT_REL)[0]
	if (!layoutRel)
		throw new InvalidOptionError(
			'import/destination-missing-layout',
			'importSlide preserve mode requires a slide layout in the destination deck'
		)
	return masterRels.resolveTarget(layoutRel.id)
}
