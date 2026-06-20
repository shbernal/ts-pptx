import { ZipWriter, readZip } from '../../zip.js'
import { ContentTypes } from './content-types.js'
import { Part } from './part.js'
import { Relationships } from './relationships.js'
import { partNameToZipPath, relsPartNameFor, zipPathToPartName } from './partnames.js'

export type OpcInput = string | number[] | Uint8Array | ArrayBuffer | Blob

const CONTENT_TYPES_ZIP_PATH = '[Content_Types].xml'
const RELATIONSHIPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml'

const textDecoder = new TextDecoder('utf-8')
const textEncoder = new TextEncoder()

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
	#parts: Map<string, Part>
	/** Content-type resolution + registration overlay. */
	readonly contentTypes: ContentTypes
	#contentTypesBytes: Uint8Array
	#relationshipsCache = new Map<string, Relationships>()

	private constructor(parts: Map<string, Part>, contentTypes: ContentTypes, contentTypesBytes: Uint8Array) {
		this.#parts = parts
		this.contentTypes = contentTypes
		this.#contentTypesBytes = contentTypesBytes
	}

	/** All parts, keyed by partname, in zip/add order. */
	get parts(): ReadonlyMap<string, Part> {
		return this.#parts
	}

	static async load(input: OpcInput): Promise<OpcPackage> {
		const entries = await readZip(input)
		const contentTypesBytes = entries.get(CONTENT_TYPES_ZIP_PATH)
		if (!contentTypesBytes) throw new Error('Not an OPC package: missing [Content_Types].xml')
		const contentTypes = ContentTypes.parse(textDecoder.decode(contentTypesBytes))

		const parts = new Map<string, Part>()
		for (const [zipPath, bytes] of entries) {
			if (zipPath === CONTENT_TYPES_ZIP_PATH) continue
			const partName = zipPathToPartName(zipPath)
			// PowerPoint leaves deleted parts in a `[trash]` folder that is not
			// registered in [Content_Types].xml and is referenced by nothing. These
			// are inert artifacts (common in real-world authored decks); drop them on
			// load rather than failing, so the package model holds only live parts.
			if (partName.startsWith('/[trash]/')) continue
			const contentType = contentTypes.contentTypeFor(partName)
			if (!contentType) {
				throw new Error(`No content type for part ${partName}: [Content_Types].xml has no matching Override or Default`)
			}
			parts.set(partName, new Part(partName, contentType, bytes))
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
	 * Add a new part and register its content type (`ensureRegistered`), so the
	 * package stays consistent. Throws if the partname is already taken.
	 */
	addPart(partName: string, contentType: string, bytes: Uint8Array): Part {
		if (this.#parts.has(partName)) throw new Error(`Cannot add part ${partName}: a part with that name already exists`)
		this.contentTypes.ensureRegistered(partName, contentType)
		const part = new Part(partName, contentType, bytes)
		this.#parts.set(partName, part)
		return part
	}

	/**
	 * Remove a part and return whether it existed. Drops its `Override` content-type
	 * registration and any cached relationship set it owned, so a later `save()`
	 * neither re-emits it nor flushes a stale `.rels` for it. Low-level: it does not
	 * touch references *to* this part (dangling rels are the caller's concern) nor
	 * cascade to parts this one referenced — see {@link Presentation.removeSlide}
	 * for the slide-aware variant that unwires the presentation and prunes orphans.
	 */
	removePart(partName: string): boolean {
		if (!this.#parts.delete(partName)) return false
		this.contentTypes.removeOverride(partName)
		// Cache is keyed by owning part; clearing it stops a stale set being flushed.
		this.#relationshipsCache.delete(partName)
		return true
	}

	/**
	 * Reserve an unused media partname `/ppt/media/<base><n>.<ext>` (n one past
	 * the highest existing index for that base). Does not create the part.
	 */
	reserveMediaPartName(extension: string, base = 'image'): string {
		const ext = extension.toLowerCase().replace(/^\./, '')
		const re = new RegExp(`^/ppt/media/${base}(\\d+)\\.${ext}$`, 'i')
		let max = 0
		for (const partName of this.#parts.keys()) {
			const match = re.exec(partName)
			if (match) max = Math.max(max, Number(match[1]))
		}
		return `/ppt/media/${base}${max + 1}.${ext}`
	}

	/**
	 * Reserve an unused partname in the same directory and naming family as
	 * `templatePartName`, with a fresh index one past the highest existing one.
	 * The template's filename is split into a base (its stem minus any trailing
	 * digits) and an extension, e.g. `/ppt/slideLayouts/slideLayout1.xml` →
	 * `/ppt/slideLayouts/slideLayout<n>.xml`, `/ppt/media/image1.png` →
	 * `/ppt/media/image<n>.png`. Used when copying a part in from another package.
	 * Does not create the part.
	 */
	reservePartNameLike(templatePartName: string): string {
		const slash = templatePartName.lastIndexOf('/')
		const dir = templatePartName.slice(0, slash)
		const fileName = templatePartName.slice(slash + 1)
		const dot = fileName.lastIndexOf('.')
		const ext = dot > 0 ? fileName.slice(dot) : '' // includes the leading '.'
		const stem = dot > 0 ? fileName.slice(0, dot) : fileName
		const base = stem.replace(/\d+$/, '')
		const escape = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
		const re = new RegExp(`^${escape(`${dir}/${base}`)}(\\d+)${escape(ext)}$`)
		let max = 0
		for (const partName of this.#parts.keys()) {
			const match = re.exec(partName)
			if (match) max = Math.max(max, Number(match[1]))
		}
		return `${dir}/${base}${max + 1}${ext}`
	}

	/**
	 * Re-emit the package. Clean parts are written from their original bytes;
	 * dirty parts from their DOM. Dirty relationship sets and content types are
	 * flushed first. Part order is preserved from load; added parts append.
	 */
	async save(): Promise<Uint8Array> {
		this.#flushRelationships()
		const zip = new ZipWriter()
		zip.add(CONTENT_TYPES_ZIP_PATH, this.contentTypes.isDirty ? this.contentTypes.serialize() : this.#contentTypesBytes)
		for (const part of this.#parts.values()) {
			zip.add(partNameToZipPath(part.partName), part.serialize())
		}
		return zip.toBytes()
	}

	/** Write each dirty `Relationships` set back into its `.rels` part (creating it if new). */
	#flushRelationships(): void {
		for (const relationships of this.#relationshipsCache.values()) {
			if (!relationships.isDirty) continue
			const relsPartName = relsPartNameFor(relationships.sourcePartName)
			const bytes = textEncoder.encode(relationships.serialize())
			// Overwriting an existing key preserves its position in the Map.
			this.#parts.set(relsPartName, new Part(relsPartName, RELATIONSHIPS_CONTENT_TYPE, bytes))
		}
	}
}
