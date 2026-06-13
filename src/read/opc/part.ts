import { parseXml, serializeXml, type Document } from '../oxml/dom.js'

const textDecoder = new TextDecoder('utf-8')
const textEncoder = new TextEncoder()

const XML_DECLARATION_RE = /^\uFEFF?<\?xml[^>]*\?>\s*/

/**
 * One OPC part. Holds the original bytes from the zip for its whole life;
 * parses to a DOM only on demand; serializes back to the original bytes
 * unless something mutated the DOM and called `markDirty()`.
 *
 * Untouched parts therefore round-trip byte-identically.
 */
export class Part {
	readonly partName: string
	readonly contentType: string
	#bytes: Uint8Array
	#dom: Document | null = null
	#dirty = false

	constructor(partName: string, contentType: string, bytes: Uint8Array) {
		this.partName = partName
		this.contentType = contentType
		this.#bytes = bytes
	}

	/** Whether this part's body is XML (by content type). */
	get isXmlPart(): boolean {
		return this.contentType.endsWith('+xml') || this.contentType === 'application/xml' || this.contentType === 'text/xml'
	}

	/** True once the body has been materialized as a DOM. */
	get isParsed(): boolean {
		return this.#dom !== null
	}

	/** The original bytes from the package. Do not mutate. */
	get bytes(): Uint8Array {
		return this.#bytes
	}

	/**
	 * Lazily parsed DOM for XML parts. Parsing alone does not mark the part
	 * dirty; call `markDirty()` after mutating the tree.
	 */
	get dom(): Document {
		if (!this.isXmlPart) throw new Error(`Part ${this.partName} (${this.contentType}) is not an XML part and has no DOM`)
		if (!this.#dom) this.#dom = parseXml(textDecoder.decode(this.#bytes))
		return this.#dom
	}

	/** Call after mutating the DOM so `serialize()` reserializes this part. */
	markDirty(): void {
		if (!this.isXmlPart) throw new Error(`Part ${this.partName} (${this.contentType}) is not an XML part and cannot be marked dirty`)
		this.#dirty = true
	}

	get isDirty(): boolean {
		return this.#dirty
	}

	/**
	 * Bytes to write on save: the original bytes when clean (byte-identical),
	 * the serialized DOM when dirty (semantically equivalent, schema-valid).
	 */
	serialize(): Uint8Array {
		if (!this.#dirty) return this.#bytes
		let xml = serializeXml(this.dom)
		if (!xml.startsWith('<?xml')) {
			const declaration = XML_DECLARATION_RE.exec(textDecoder.decode(this.#bytes))?.[0]
			xml = (declaration ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n') + xml
		}
		return textEncoder.encode(xml)
	}
}
