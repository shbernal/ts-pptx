/**
 * Rescaling an imported slide's geometry onto the destination canvas.
 *
 * The memo of already-rescaled parts is the load-bearing part: a layout or master shared by
 * several imports from one source deck must be rescaled exactly once, or the second import
 * scales it again on top of the first. It is owned by the calling `Presentation` and threaded
 * through here, so it lives as long as the deck does rather than as long as one call.
 */

import { firstChild } from '../../oxml/dom.js'
import { computeRescale, rescaleSpTree, type RescaleTransform } from '../rescale.js'
import { resolveSingleRel } from './part-index.js'
import type { ImportSlideOptions, SlideSize } from '../presentation-types.js'
import type { Presentation } from '../presentation.js'
import { SLIDE_LAYOUT_REL, SLIDE_MASTER_REL } from '../rel-types.js'

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
	rescaledParts: Set<string>,
	slidePartName: string,
	theme: ImportSlideOptions['theme'],
	source: SlideSize,
	target: SlideSize,
	mode: 'fit' | 'stretch'
): void {
	const transform = computeRescale(source, target, mode)
	rescalePartGeometry(dest, rescaledParts, slidePartName, transform)
	if (theme === undefined || theme === 'copy') {
		const layout = resolveSingleRel(dest.opc, slidePartName, SLIDE_LAYOUT_REL)
		const master = layout ? resolveSingleRel(dest.opc, layout, SLIDE_MASTER_REL) : null
		if (layout) rescalePartGeometry(dest, rescaledParts, layout, transform)
		if (master) rescalePartGeometry(dest, rescaledParts, master, transform)
	}
}

/**
 * Rescale one part's `p:spTree` geometry in place. Idempotent per part
 * (`rescaledParts`), so a layout/master shared across repeated imports from one
 * source is rescaled exactly once.
 */
function rescalePartGeometry(
	dest: Presentation,
	rescaledParts: Set<string>,
	partName: string,
	transform: RescaleTransform
): void {
	if (rescaledParts.has(partName)) return
	rescaledParts.add(partName)
	const part = dest.opc.part(partName)
	const root = part?.dom.documentElement
	const cSld = root && firstChild(root, 'p:cSld')
	const spTree = cSld && firstChild(cSld, 'p:spTree')
	if (!part || !spTree) return
	rescaleSpTree(spTree, transform)
	part.markDirty()
}
