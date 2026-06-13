/**
 * Read-model entry point: `Presentation` wraps an `OpcPackage` and exposes a
 * navigable, typed view of the deck (slides → shapes → text), backed by the
 * live DOM so the same nodes can later be mutated.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { relativePartName, relsPartNameFor } from '../opc/partnames.js'
import { attr, createElement, firstChild, getElements, intValue, setAttr } from '../oxml/dom.js'
import { Slide } from './slide.js'

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

/** ST_SlideId minimum (ECMA-376): slide ids live in [256, 2147483647]. */
const MIN_SLIDE_ID = 256

/** Slide dimensions, in both EMU (the OOXML unit) and inches. */
export interface SlideSize {
	widthEmu: number
	heightEmu: number
	widthIn: number
	heightIn: number
}

export class Presentation {
	#presentationPart: Part | undefined

	private constructor(readonly opc: OpcPackage) {}

	/** Open a `.pptx` from bytes and wrap it as a navigable `Presentation`. */
	static async load(input: OpcInput): Promise<Presentation> {
		return new Presentation(await OpcPackage.load(input))
	}

	/** Wrap an already-loaded OPC package (e.g. from the lower-level API). */
	static fromPackage(opc: OpcPackage): Presentation {
		return new Presentation(opc)
	}

	/** The main presentation part (`/ppt/presentation.xml`), resolved via the package `officeDocument` relationship. */
	get presentationPart(): Part {
		if (this.#presentationPart) return this.#presentationPart
		const packageRels = this.opc.relationshipsFor('/')
		const officeDocument = packageRels.byType(OFFICE_DOCUMENT_REL)
		if (officeDocument.length !== 1) {
			throw new Error(`Expected exactly one officeDocument relationship, found ${officeDocument.length}`)
		}
		const partName = packageRels.resolveTarget(officeDocument[0].id)
		const part = this.opc.part(partName)
		if (!part) throw new Error(`officeDocument relationship targets a missing part: ${partName}`)
		this.#presentationPart = part
		return part
	}

	/** The slides in presentation order (resolved from `p:sldIdLst` + the presentation's relationships). */
	get slides(): Slide[] {
		const root = this.presentationPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (!sldIdLst) return []
		const rels = this.opc.relationshipsFor(this.presentationPart.partName)
		const slides: Slide[] = []
		let index = 0
		for (const sldId of getElements(sldIdLst, 'p:sldId')) {
			const relId = attr(sldId, 'r:id')
			if (!relId) continue
			const partName = rels.resolveTarget(relId)
			const part = this.opc.part(partName)
			if (!part) throw new Error(`Slide relationship ${relId} targets a missing part: ${partName}`)
			slides.push(new Slide(this, part, intValue(attr(sldId, 'id')) ?? 0, index++))
		}
		return slides
	}

	/** Slide dimensions (`p:sldSz`), or `null` if the presentation declares none. */
	get slideSize(): SlideSize | null {
		const root = this.presentationPart.dom.documentElement
		const sldSz = root && firstChild(root, 'p:sldSz')
		if (!sldSz) return null
		const widthEmu = intValue(attr(sldSz, 'cx'))
		const heightEmu = intValue(attr(sldSz, 'cy'))
		if (widthEmu === null || heightEmu === null) return null
		return { widthEmu, heightEmu, widthIn: emuToInches(widthEmu), heightIn: emuToInches(heightEmu) }
	}

	/**
	 * Duplicate the slide at `index` and append the copy at the end of the deck,
	 * returning it. The new slide part copies the source bytes verbatim and
	 * shares the source's relationship targets (layout, images, …) by copying its
	 * `.rels`; a new presentation→slide relationship and a `p:sldId` entry are
	 * wired up. Marks the presentation part dirty.
	 *
	 * Note: relationships are copied as-is, so a source slide that owns a
	 * one-to-one part (e.g. a notes slide) would end up shared with the clone.
	 */
	cloneSlide(index: number): Slide {
		const source = this.slides[index]
		if (!source) throw new Error(`No slide at index ${index} to clone`)
		const opc = this.opc
		const sourcePart = source.part

		// 1. Copy the slide part bytes verbatim into a fresh slide partname.
		const newPartName = this.#reserveSlidePartName()
		const newPart = opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)

		// 2. Copy the slide's relationships (targets resolve identically — same dir).
		const sourceRels = opc.part(relsPartNameFor(sourcePart.partName))
		if (sourceRels) opc.addPart(relsPartNameFor(newPartName), sourceRels.contentType, sourceRels.bytes)

		// 3. Wire a presentation→slide relationship.
		const presPart = this.presentationPart
		const presRels = opc.relationshipsFor(presPart.partName)
		const relId = presRels.add(SLIDE_REL, relativePartName(presPart.partName, newPartName)).id

		// 4. Append a p:sldId entry referencing the new relationship.
		const root = presPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (!sldIdLst) throw new Error('presentation.xml has no p:sldIdLst to append a slide to')
		const existing = getElements(sldIdLst, 'p:sldId')
		const newIndex = existing.length
		const newSlideId = this.#nextSlideId(existing)
		const sldId = createElement(presPart.dom, 'p:sldId')
		setAttr(sldId, 'id', String(newSlideId))
		setAttr(sldId, 'r:id', relId)
		sldIdLst.appendChild(sldId)
		presPart.markDirty()

		return new Slide(this, newPart, newSlideId, newIndex)
	}

	/** Reserve an unused `/ppt/slides/slide<n>.xml` partname (n one past the highest). */
	#reserveSlidePartName(): string {
		const re = /^\/ppt\/slides\/slide(\d+)\.xml$/
		let max = 0
		for (const partName of this.opc.parts.keys()) {
			const match = re.exec(partName)
			if (match) max = Math.max(max, Number(match[1]))
		}
		return `/ppt/slides/slide${max + 1}.xml`
	}

	/** A slide id one past the highest existing, but at least ST_SlideId's minimum. */
	#nextSlideId(sldIds: ReturnType<typeof getElements>): number {
		let max = MIN_SLIDE_ID - 1
		for (const sldId of sldIds) {
			const id = intValue(attr(sldId, 'id'))
			if (id !== null && id > max) max = id
		}
		return max + 1
	}

	/** Re-emit the package; untouched parts stay byte-identical (see `OpcPackage.save`). */
	async save(): Promise<Uint8Array> {
		return this.opc.save()
	}
}
