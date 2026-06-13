/**
 * Read-model proxy for one slide (`p:sld`), backed by its live part DOM.
 */
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { attr, firstChild, type Element } from '../oxml/dom.js'
import type { Presentation } from './presentation.js'
import { buildShapes, type Shape } from './shapes.js'

export class Slide {
	constructor(
		readonly presentation: Presentation,
		/** The slide's OPC part (`/ppt/slides/slideN.xml`). */
		readonly part: Part,
		/** The slide id from `p:sldIdLst` (`p:sldId/@id`). */
		readonly slideId: number,
		/** Zero-based position in presentation order. */
		readonly index: number
	) {}

	/** Partname of this slide's part. */
	get partName(): string {
		return this.part.partName
	}

	/** This slide part's relationships (image embeds, layout, hyperlinks, …). */
	get relationships(): Relationships {
		return this.presentation.opc.relationshipsFor(this.partName)
	}

	/** Authoring name of the slide (`p:cSld/@name`), or `null` if unnamed. */
	get name(): string | null {
		const cSld = this.#cSld()
		return cSld ? attr(cSld, 'name') : null
	}

	/** Top-level shapes in the slide's shape tree, in document order. */
	get shapes(): Shape[] {
		const spTree = this.#spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	#cSld(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'p:cSld') : null
	}

	#spTree(): Element | null {
		const cSld = this.#cSld()
		return cSld ? firstChild(cSld, 'p:spTree') : null
	}
}
