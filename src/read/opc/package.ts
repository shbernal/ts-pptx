import JSZip from 'jszip'
import { ContentTypes } from './content-types.js'
import { Part } from './part.js'
import { Relationships } from './relationships.js'
import { partNameToZipPath, relsPartNameFor, zipPathToPartName } from './partnames.js'

export type OpcInput = string | number[] | Uint8Array | ArrayBuffer | Blob

const CONTENT_TYPES_ZIP_PATH = '[Content_Types].xml'

const textDecoder = new TextDecoder('utf-8')

/**
 * An OPC package loaded from `.pptx` bytes.
 *
 * Fidelity contract: parts that are never mutated (`Part.markDirty()`) are
 * written back byte-identically by `save()`. Whole-zip byte-identity is not
 * promised — zip metadata and compression may differ — but part bodies, the
 * part-name set, and part order are preserved.
 */
export class OpcPackage {
	/** All parts, keyed by partname (e.g. `/ppt/slides/slide1.xml`), in original zip order. */
	readonly parts: ReadonlyMap<string, Part>
	/** Content-type resolution overlay (read-only in Phase 1). */
	readonly contentTypes: ContentTypes
	#contentTypesBytes: Uint8Array
	#relationshipsCache = new Map<string, Relationships>()

	private constructor(parts: Map<string, Part>, contentTypes: ContentTypes, contentTypesBytes: Uint8Array) {
		this.parts = parts
		this.contentTypes = contentTypes
		this.#contentTypesBytes = contentTypesBytes
	}

	static async load(input: OpcInput): Promise<OpcPackage> {
		const zip = await JSZip.loadAsync(input)
		const contentTypesEntry = zip.file(CONTENT_TYPES_ZIP_PATH)
		if (!contentTypesEntry) throw new Error('Not an OPC package: missing [Content_Types].xml')
		const contentTypesBytes = await contentTypesEntry.async('uint8array')
		const contentTypes = ContentTypes.parse(textDecoder.decode(contentTypesBytes))

		const parts = new Map<string, Part>()
		for (const entry of Object.values(zip.files)) {
			if (entry.dir || entry.name === CONTENT_TYPES_ZIP_PATH) continue
			const partName = zipPathToPartName(entry.name)
			const contentType = contentTypes.contentTypeFor(partName)
			if (!contentType) {
				throw new Error(`No content type for part ${partName}: [Content_Types].xml has no matching Override or Default`)
			}
			parts.set(partName, new Part(partName, contentType, await entry.async('uint8array')))
		}
		return new OpcPackage(parts, contentTypes, contentTypesBytes)
	}

	part(partName: string): Part | undefined {
		return this.parts.get(partName)
	}

	partsByContentType(contentType: string): Part[] {
		return [...this.parts.values()].filter((part) => part.contentType === contentType)
	}

	/**
	 * Relationships owned by a part ('/' for the package-level
	 * `/_rels/.rels`). Parts without a `.rels` part get an empty set.
	 */
	relationshipsFor(sourcePartName = '/'): Relationships {
		const cached = this.#relationshipsCache.get(sourcePartName)
		if (cached) return cached
		const relsPart = this.parts.get(relsPartNameFor(sourcePartName))
		const relationships = relsPart
			? Relationships.parse(textDecoder.decode(relsPart.bytes), sourcePartName)
			: Relationships.empty(sourcePartName)
		this.#relationshipsCache.set(sourcePartName, relationships)
		return relationships
	}

	/**
	 * Re-emit the package. Clean parts are written from their original bytes;
	 * dirty parts from their DOM. Part order is preserved from load.
	 */
	async save(): Promise<Uint8Array> {
		const zip = new JSZip()
		zip.file(CONTENT_TYPES_ZIP_PATH, this.#contentTypesBytes)
		for (const part of this.parts.values()) {
			zip.file(partNameToZipPath(part.partName), part.serialize())
		}
		return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
	}
}
