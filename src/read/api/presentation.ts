/**
 * Read-model entry point: `Presentation` wraps an `OpcPackage` and exposes a
 * navigable, typed view of the deck (slides → shapes → text), backed by the
 * live DOM so the same nodes can later be mutated.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { attr, firstChild, getElements, intValue } from '../oxml/dom.js'
import { Slide } from './slide.js'

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

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

	/** Re-emit the package; untouched parts stay byte-identical (see `OpcPackage.save`). */
	async save(): Promise<Uint8Array> {
		return this.opc.save()
	}
}
