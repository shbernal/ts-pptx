import { escapeXmlAttribute, getElements, parseXml } from '../oxml/dom.js'
import { partNameExtension } from './partnames.js'

/**
 * Read-only overlay over `[Content_Types].xml`, used to resolve part content
 * types at load. The original bytes pass through verbatim on save; this
 * object is never the source of truth until a mutation API lands (Phase 3).
 */
export class ContentTypes {
	#defaults = new Map<string, string>()
	#overrides = new Map<string, string>()

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
