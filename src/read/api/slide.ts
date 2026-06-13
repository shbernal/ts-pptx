/**
 * Read-model proxy for one slide (`p:sld`), backed by its live part DOM.
 */
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { OOXML_NS, attr, createElement, firstChild, intValue, setAttr, type Document, type Element } from '../oxml/dom.js'
import type { Presentation } from './presentation.js'
import { AutoShape, buildShapes, type Shape } from './shapes.js'

/** Options for {@link Slide.addTextBox}. Geometry is in EMU. */
export interface AddTextBoxOptions {
	/** Left edge in EMU (`a:off/@x`). */
	left: number
	/** Top edge in EMU (`a:off/@y`). */
	top: number
	/** Width in EMU (`a:ext/@cx`); must be positive. */
	width: number
	/** Height in EMU (`a:ext/@cy`); must be positive. */
	height: number
	/** Initial text; omitted/empty yields an empty paragraph. */
	text?: string
	/** Shape name (`p:cNvPr/@name`); defaults to `TextBox <id>`. */
	name?: string
}

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

	/**
	 * Append a text box (`p:sp` with `txBox="1"`) to the slide's shape tree and
	 * return it. Geometry is required (EMU); width and height must be positive.
	 * Allocates a drawing id unique within the slide. Marks the slide part dirty.
	 */
	addTextBox(options: AddTextBoxOptions): AutoShape {
		const { left, top, width, height } = options
		requireFinite(left, 'left')
		requireFinite(top, 'top')
		requirePositive(width, 'width')
		requirePositive(height, 'height')

		const spTree = this.#spTree()
		if (!spTree) throw new Error(`Slide ${this.partName} has no spTree to add a shape to`)
		const doc = spTree.ownerDocument
		if (!doc) throw new Error('Slide DOM has no owner document')

		const id = this.#nextShapeId()
		const sp = buildTextBox(doc, {
			id,
			name: options.name ?? `TextBox ${id}`,
			text: options.text ?? '',
			left: Math.round(left),
			top: Math.round(top),
			width: Math.round(width),
			height: Math.round(height),
		})
		// A shape goes after grpSpPr and before any trailing p:extLst on the tree.
		spTree.insertBefore(sp, firstChild(spTree, 'p:extLst'))
		this.part.markDirty()
		return new AutoShape(sp, this)
	}

	/** The smallest drawing id (`p:cNvPr/@id`) not already used on the slide. */
	#nextShapeId(): number {
		const root = this.part.dom.documentElement
		let max = 1
		if (root) {
			const cNvPrs = root.getElementsByTagNameNS(OOXML_NS.p, 'cNvPr')
			for (let i = 0; i < cNvPrs.length; i++) {
				const id = intValue(attr(cNvPrs[i], 'id'))
				if (id !== null && id > max) max = id
			}
		}
		return max + 1
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

function requireFinite(value: number, name: string): void {
	if (!Number.isFinite(value)) throw new Error(`${name} must be a finite number of EMU, got ${value}`)
}

function requirePositive(value: number, name: string): void {
	requireFinite(value, name)
	if (value <= 0) throw new Error(`${name} must be positive, got ${value}`)
}

interface TextBoxSpec {
	id: number
	name: string
	text: string
	left: number
	top: number
	width: number
	height: number
}

/** Build a minimal, schema-valid text-box `p:sp` element (not yet attached). */
function buildTextBox(doc: Document, spec: TextBoxSpec): Element {
	const make = (qname: string): Element => createElement(doc, qname)
	const append = (parent: Element, qname: string): Element => {
		const child = make(qname)
		parent.appendChild(child)
		return child
	}

	const sp = make('p:sp')

	const nvSpPr = append(sp, 'p:nvSpPr')
	const cNvPr = append(nvSpPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvSpPr = append(nvSpPr, 'p:cNvSpPr')
	setAttr(cNvSpPr, 'txBox', '1')
	append(nvSpPr, 'p:nvPr')

	const spPr = append(sp, 'p:spPr')
	const xfrm = append(spPr, 'a:xfrm')
	const off = append(xfrm, 'a:off')
	setAttr(off, 'x', String(spec.left))
	setAttr(off, 'y', String(spec.top))
	const ext = append(xfrm, 'a:ext')
	setAttr(ext, 'cx', String(spec.width))
	setAttr(ext, 'cy', String(spec.height))
	const prstGeom = append(spPr, 'a:prstGeom')
	setAttr(prstGeom, 'prst', 'rect')
	append(prstGeom, 'a:avLst')

	const txBody = append(sp, 'p:txBody')
	append(txBody, 'a:bodyPr')
	append(txBody, 'a:lstStyle')
	const p = append(txBody, 'a:p')
	if (spec.text !== '') {
		const r = append(p, 'a:r')
		const t = append(r, 'a:t')
		t.textContent = spec.text
		if (spec.text !== spec.text.trim()) setAttr(t, 'xml:space', 'preserve')
	}

	return sp
}
