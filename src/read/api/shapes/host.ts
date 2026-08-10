/**
 * The contract a shape proxy needs from the part that owns its shape tree.
 *
 * Three parts carry a `p:cSld/p:spTree`, and each one's shapes are the same five
 * elements resolved against the same tiers: a {@link import('../slide.js').Slide},
 * a {@link import('../chrome.js').SlideLayout}, and a
 * {@link import('../chrome.js').SlideMaster}. A shape proxy therefore holds a
 * back-reference to its *host* rather than to a slide specifically — it needs the
 * owning part (to mark dirty and to build text frames against), that part's
 * relationships (image embeds, hyperlinks, chart parts), the deck's package (to
 * reach a referenced part), the host's resolved theme context (so a `schemeClr`
 * token becomes a literal hex), and a deep shape lookup (so a connector binding
 * resolves to the shape it names).
 *
 * Every member here is one a slide, a layout, and a master can all answer, which
 * is what lets `Slide.shapes`, `SlideLayout.shapes` and `SlideMaster.shapes` return
 * the same `AnyShape` union: a consumer's shape-walking code applies unchanged to a
 * template's bands, rules and logos, not only to a slide's own content.
 */
import type { OpcPackage } from '../../opc/package.js'
import type { Part } from '../../opc/part.js'
import type { Relationships } from '../../opc/relationships.js'
import type { ThemeContext } from '../../oxml/theme.js'
import type { AnyShape } from '../shapes.js'

/** The part-level surface a {@link import('./base.js').Shape} resolves against. */
export interface ShapeHost {
	/** The OPC part holding the shape tree (`p:sld` / `p:sldLayout` / `p:sldMaster`). */
	readonly part: Part
	/** Partname of {@link part}. */
	readonly partName: string
	/** The package the host belongs to, for reaching a referenced part (image, chart, …). */
	readonly opc: OpcPackage
	/** {@link part}'s relationships — image embeds, hyperlinks, chart references. */
	readonly relationships: Relationships
	/** The host's resolved colour/font context, backing every `resolved*` getter. */
	themeContext(): ThemeContext
	/** The shape anywhere in this host's tree with the given drawing id, or `undefined`. */
	shapeByIdDeep(id: number): AnyShape | undefined
}
