/**
 * The theme context a source slide is flattened against.
 *
 * `preserve`-mode import bakes a slide's *source* look into its XML — scheme colours and
 * style-matrix fills resolved to literals — so it keeps that look after being attached to a
 * different deck's master. Doing so needs the source subgraph's colour map, colour scheme,
 * format scheme, and the background the slide inherits but does not itself carry. This
 * gathers exactly that.
 *
 * Pure with respect to the destination deck: it reads only the source package. It was a
 * private method on `Presentation` that used no instance state.
 */

import type { Element } from '../../oxml/dom.js'
import type { OpcPackage } from '../../opc/package.js'
import type { FlattenContext } from './flatten.js'
import { resolveSlideColorContext, resolveSlideThemeParts } from '../theme-context.js'
import { backgroundElementOf } from '../slide-background.js'
import { TABLE_STYLES_CONTENT_TYPE } from '../../../ooxml/rel-types.js'

/**
 * The background the slide effectively inherits from its source subgraph: the layout's
 * `p:bg`, else the master's. Returns `null` when the slide carries its own `p:bg` (it stays
 * on the slide and is flattened directly) or when none exists.
 */
function effectiveBackground(
	sourceOpc: OpcPackage,
	slideRoot: Element | null,
	layoutPartName: string | null,
	masterPartName: string | null
): Element | null {
	if (slideRoot && backgroundElementOf(slideRoot)) return null
	const layoutRoot = layoutPartName ? (sourceOpc.part(layoutPartName)?.dom.documentElement ?? null) : null
	const masterRoot = masterPartName ? (sourceOpc.part(masterPartName)?.dom.documentElement ?? null) : null
	return (layoutRoot && backgroundElementOf(layoutRoot)) ?? (masterRoot && backgroundElementOf(masterRoot)) ?? null
}

/**
 * Gather the flatten context for a source slide: the colour context the slide's read-model
 * getters resolve against, built from one slide → layout → master → theme walk, plus the
 * background the slide inherits.
 *
 * The presentation's default text style rides along, because a run in a shape that is not a
 * placeholder inherits its size, weight and colour from it, and the rebind would hand that run the
 * destination's instead. The flatten pass bakes those values onto the run. The source table styles
 * ride along for the same bake in table cells, which must leave what a table style states to it.
 * @param {OpcPackage} sourceOpc - the source package to read
 * @param {string} slidePartName - partname of the source slide
 * @return {FlattenContext} the context {@link flattenSlide} / {@link flattenShape} resolve against
 */
export function sourceFlattenContext(sourceOpc: OpcPackage, slidePartName: string): FlattenContext {
	const parts = resolveSlideThemeParts(sourceOpc, slidePartName)
	const ctx = resolveSlideColorContext(sourceOpc, parts)
	return {
		...ctx,
		inheritedBackground: effectiveBackground(sourceOpc, parts.slideRoot, parts.layoutPartName, parts.masterPartName),
		tableStyles: sourceOpc.partsByContentType(TABLE_STYLES_CONTENT_TYPE)[0]?.dom.documentElement ?? null,
	}
}
