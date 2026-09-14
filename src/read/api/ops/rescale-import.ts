/**
 * Rescaling an imported slide's geometry onto the destination canvas.
 *
 * The memo of already-rescaled parts is the load-bearing part: a layout or master shared by
 * several imports from one source deck must be rescaled exactly once, or the second import
 * scales it again on top of the first. It records the mode each part was rescaled in, because a
 * shared part can hold only one scaling: a later import asking for the other mode is refused
 * before it copies anything. The memo is owned by the calling `Presentation` and threaded through
 * here, so it lives as long as the deck does rather than as long as one call.
 */

import { computeRescale, rescaleSpTree, type RescaleMode, type RescaleTransform } from './rescale.js'
import { resolveSingleRel } from './part-index.js'
import type { ImportContext } from './part-copy.js'
import type { ImportSlideOptions, SlideSize } from '../presentation-types.js'
import type { Presentation } from '../presentation.js'
import { SLIDE_LAYOUT_REL, SLIDE_MASTER_REL } from '../../../ooxml/rel-types.js'
import { spTreeOf } from '../../oxml/slide-dom.js'
import { InternalError, InvalidOptionError } from '../../../errors.js'

/**
 * Rescale an imported slide's geometry onto this deck's canvas (the `rescale`
 * option of {@link Presentation.importSlide}). Rewrites every top-level shape/group/
 * graphicFrame transform and table grid on the slide; in `copy` mode also
 * rescales the imported layout and master shape trees (resolved via the
 * slide → layout → master rel chain) so inherited placeholder/background geometry
 * stays aligned. `preserve`/`restyle` rebind to this deck's own master/layout —
 * already the destination size — so only the slide is touched. Geometry only:
 * font sizes and line widths are left as authored.
 */
export function rescaleImportedGeometry(
	dest: Presentation,
	rescaledParts: Map<string, RescaleMode>,
	slidePartName: string,
	theme: ImportSlideOptions['theme'],
	source: SlideSize,
	target: SlideSize,
	mode: RescaleMode
): void {
	const transform = computeRescale(source, target, mode)
	rescalePartGeometry(dest, rescaledParts, slidePartName, transform, mode)
	if (theme === undefined || theme === 'copy') {
		const layout = resolveSingleRel(dest.opc, slidePartName, SLIDE_LAYOUT_REL)
		const master = layout ? resolveSingleRel(dest.opc, layout, SLIDE_MASTER_REL) : null
		if (layout) rescalePartGeometry(dest, rescaledParts, layout, transform, mode)
		if (master) rescalePartGeometry(dest, rescaledParts, master, transform, mode)
	}
}

/**
 * Refuse a `copy`-themed import whose rescale disagrees with the one an earlier import from the
 * same source already applied to the layout or master this import would share.
 *
 * A second import from a source reuses the layout and master the first one copied, and those were
 * rewritten in the first import's mode. Scaling them again in the other mode would move them under
 * the first import's page. Checked before anything is copied, through the copy registry that
 * decides the reuse, so a refusal leaves the deck as it was.
 * @param rescaledParts - the deck's rescale memo
 * @param ctx - the import's context, whose registry names what an earlier import copied
 * @param sourceSlidePartName - the source page being imported
 * @param mode - the mode this import rescales in
 * @param api - the public method, opening the error message
 */
export function requireRescaleAgrees(
	rescaledParts: ReadonlyMap<string, RescaleMode>,
	ctx: ImportContext,
	sourceSlidePartName: string,
	mode: RescaleMode,
	api: string
): void {
	const layout = resolveSingleRel(ctx.source, sourceSlidePartName, SLIDE_LAYOUT_REL)
	const master = layout ? resolveSingleRel(ctx.source, layout, SLIDE_MASTER_REL) : null
	for (const sourcePartName of [layout, master]) {
		const destPartName = sourcePartName ? ctx.registry.get(sourcePartName) : undefined
		const applied = destPartName === undefined ? undefined : rescaledParts.get(destPartName)
		if (applied !== undefined && applied !== mode) {
			throw new InvalidOptionError(
				'import/rescale-conflict',
				`${api}: ${destPartName} was rescaled with ${JSON.stringify(applied)} by an earlier import from this source, and this import asks for ${JSON.stringify(mode)}`
			)
		}
	}
}

/**
 * Rescale one part's `p:spTree` geometry in place. Idempotent per part
 * (`rescaledParts`), so a layout/master shared across repeated imports from one
 * source is rescaled exactly once.
 */
function rescalePartGeometry(
	dest: Presentation,
	rescaledParts: Map<string, RescaleMode>,
	partName: string,
	transform: RescaleTransform,
	mode: RescaleMode
): void {
	const applied = rescaledParts.get(partName)
	if (applied === mode) return
	if (applied !== undefined) {
		throw new InternalError(
			'import/rescale-mode-changed',
			`${partName} was rescaled with ${JSON.stringify(applied)} and is now asked for ${JSON.stringify(mode)}; the import should have been refused before it copied anything`
		)
	}
	rescaledParts.set(partName, mode)
	const part = dest.opc.part(partName)
	const spTree = spTreeOf(part?.dom.documentElement)
	if (!part || !spTree) return
	rescaleSpTree(spTree, transform)
	part.markDirty()
}
