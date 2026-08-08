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

import { firstChild, type Element } from '../../oxml/dom.js'
import type { OpcPackage } from '../../opc/package.js'
import type { FlattenContext } from './flatten.js'
import { resolveSlideThemeParts } from '../theme-context.js'

/** The `p:cSld/p:bg` element of a slide/layout/master root, or `null`. */
function backgroundOf(root: Element): Element | null {
	const cSld = firstChild(root, 'p:cSld')
	return cSld ? firstChild(cSld, 'p:bg') : null
}

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
	if (slideRoot && backgroundOf(slideRoot)) return null
	const layoutRoot = layoutPartName ? (sourceOpc.part(layoutPartName)?.dom.documentElement ?? null) : null
	const masterRoot = masterPartName ? (sourceOpc.part(masterPartName)?.dom.documentElement ?? null) : null
	return (layoutRoot && backgroundOf(layoutRoot)) ?? (masterRoot && backgroundOf(masterRoot)) ?? null
}

/**
 * Gather the flatten context for a source slide: walk slide → layout → master → theme,
 * reading the effective colour map (the slide's `clrMapOvr` override, or the master
 * `clrMap`), the theme `clrScheme`, and the theme `fmtScheme`.
 * @param {OpcPackage} sourceOpc - the source package to read
 * @param {string} slidePartName - partname of the source slide
 * @return {FlattenContext} the context {@link flattenSlide} / {@link flattenShape} resolve against
 */
export function sourceFlattenContext(sourceOpc: OpcPackage, slidePartName: string): FlattenContext {
	// Reuse the shared slide → layout → master → theme walk (also backing the read-model
	// colour getters), then layer the flatten-only needs on top.
	const parts = resolveSlideThemeParts(sourceOpc, slidePartName)
	const themeElements = parts.themeElements
	return {
		clrMap: parts.clrMap,
		clrScheme: parts.clrScheme,
		fmtScheme: themeElements ? firstChild(themeElements, 'a:fmtScheme') : null,
		inheritedBackground: effectiveBackground(sourceOpc, parts.slideRoot, parts.layoutPartName, parts.masterPartName),
		layoutRoot: parts.layoutRoot,
		masterRoot: parts.masterRoot,
	}
}
