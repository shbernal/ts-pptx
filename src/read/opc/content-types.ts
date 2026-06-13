import { escapeXmlAttribute, getElements, parseXml } from '../oxml/dom.js'
import { partNameExtension } from './partnames.js'

/**
 * Overlay over `[Content_Types].xml`, used to resolve part content types at
 * load and to register new ones when parts are added. While clean, the original
 * bytes pass through verbatim on save (byte-identical); once a mutation marks it
 * dirty, `OpcPackage.save()` writes `serialize()` instead.
 */
export class ContentTypes {
	#defaults = new Map<string, string>()
	#overrides = new Map<string, string>()
	#dirty = false

	static parse(xml: string): ContentTypes {
		const contentTypes = new ContentTypes()
		const types = parseXml(xml).documentElement
		if (!types || types.localName !== 'Types') throw new Error('[Content_Types].xml: expected <Types> root element')
		for (const element of getElements(types, 'ct:Default')) {
			const extension = element.getAttribute('Extension')
			const contentType = element.getAttribute('ContentType')
			if (!extension || !contentType) throw new Error('[Content_Types].xml: <Default> missing Extension or ContentType')
			contentTypes.#defaults.set(extension.toLowerCase(), contentType)
		}
		for (const element of getElements(types, 'ct:Override')) {
			const partName = element.getAttribute('PartName')
			const contentType = element.getAttribute('ContentType')
			if (!partName || !contentType) throw new Error('[Content_Types].xml: <Override> missing PartName or ContentType')
			contentTypes.#overrides.set(partName, contentType)
		}
		return contentTypes
	}

	/** Exact `Override` match first, else `Default` by lowercased extension. */
	contentTypeFor(partName: string): string | undefined {
		return this.#overrides.get(partName) ?? this.#defaults.get(partNameExtension(partName))
	}

	/** True once a registration changed the overlay; `serialize()` is then authoritative. */
	get isDirty(): boolean {
		return this.#dirty
	}

	/**
	 * Ensure `partName` resolves to `contentType`, adding the minimum needed.
	 * A no-op when an existing `Default`/`Override` already resolves it to the
	 * same type; otherwise an `Override` is added (always valid, never conflicts
	 * with a differing `Default`).
	 */
	ensureRegistered(partName: string, contentType: string): void {
		if (this.contentTypeFor(partName) === contentType) return
		this.#overrides.set(partName, contentType)
		this.#dirty = true
	}

	/** Register a `Default` content type for a (lowercased) extension if absent. */
	ensureDefault(extension: string, contentType: string): void {
		const ext = extension.toLowerCase()
		if (this.#defaults.get(ext) === contentType) return
		this.#defaults.set(ext, contentType)
		this.#dirty = true
	}

	serialize(): string {
		const lines: string[] = [
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
		]
		for (const [extension, contentType] of this.#defaults) {
			lines.push(`<Default Extension="${escapeXmlAttribute(extension)}" ContentType="${escapeXmlAttribute(contentType)}"/>`)
		}
		for (const [partName, contentType] of this.#overrides) {
			lines.push(`<Override PartName="${escapeXmlAttribute(partName)}" ContentType="${escapeXmlAttribute(contentType)}"/>`)
		}
		lines.push('</Types>')
		return lines.join('')
	}
}
