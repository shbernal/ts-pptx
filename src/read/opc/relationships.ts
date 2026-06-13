import { escapeXmlAttribute, getElements, parseXml } from '../oxml/dom.js'
import { resolveRelativePartName } from './partnames.js'

export interface Relationship {
	id: string
	type: string
	target: string
	targetMode?: 'Internal' | 'External'
}

/**
 * Read-only overlay over one `.rels` part, keyed by relationship id. Like
 * `ContentTypes`, this is a query view: the `.rels` bytes pass through
 * verbatim on save until a mutation API lands (Phase 3).
 */
export class Relationships {
	/** Partname of the part owning these relationships ('/' for package-level). */
	readonly sourcePartName: string
	#byId = new Map<string, Relationship>()

	private constructor(sourcePartName: string) {
		this.sourcePartName = sourcePartName
	}

	/**
	 * @param xml body of the `.rels` part
	 * @param sourcePartName partname of the owning part ('/' for `/_rels/.rels`)
	 */
	static parse(xml: string, sourcePartName: string): Relationships {
		const relationships = new Relationships(sourcePartName)
		const root = parseXml(xml).documentElement
		if (!root || root.localName !== 'Relationships') throw new Error(`Relationships of ${sourcePartName}: expected <Relationships> root element`)
		for (const element of getElements(root, 'pr:Relationship')) {
			const id = element.getAttribute('Id')
			const type = element.getAttribute('Type')
			const target = element.getAttribute('Target')
			if (!id || !type || !target) throw new Error(`Relationships of ${sourcePartName}: <Relationship> missing Id, Type, or Target`)
			const targetModeAttribute = element.getAttribute('TargetMode')
			const targetMode = targetModeAttribute === 'Internal' || targetModeAttribute === 'External' ? targetModeAttribute : undefined
			if (targetModeAttribute !== null && targetMode === undefined) {
				throw new Error(`Relationships of ${sourcePartName}: invalid TargetMode "${targetModeAttribute}" on ${id}`)
			}
			const relationship: Relationship = { id, type, target }
			if (targetMode) relationship.targetMode = targetMode
			if (relationships.#byId.has(id)) throw new Error(`Relationships of ${sourcePartName}: duplicate relationship id ${id}`)
			relationships.#byId.set(id, relationship)
		}
		return relationships
	}

	/** Empty relationship set (for parts that have no `.rels`). */
	static empty(sourcePartName: string): Relationships {
		return new Relationships(sourcePartName)
	}

	get size(): number {
		return this.#byId.size
	}

	get(id: string): Relationship | undefined {
		return this.#byId.get(id)
	}

	byType(type: string): Relationship[] {
		return [...this.#byId.values()].filter((relationship) => relationship.type === type)
	}

	[Symbol.iterator](): IterableIterator<Relationship> {
		return this.#byId.values()
	}

	/** Resolve an internal relationship's target to an absolute partname. */
	resolveTarget(id: string): string {
		const relationship = this.#byId.get(id)
		if (!relationship) throw new Error(`Relationships of ${this.sourcePartName}: no relationship with id ${id}`)
		if (relationship.targetMode === 'External') {
			throw new Error(`Relationships of ${this.sourcePartName}: ${id} is External (${relationship.target}) and has no partname`)
		}
		return resolveRelativePartName(this.sourcePartName, relationship.target)
	}

	serialize(): string {
		const lines: string[] = [
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
		]
		for (const { id, type, target, targetMode } of this.#byId.values()) {
			const mode = targetMode ? ` TargetMode="${targetMode}"` : ''
			lines.push(`<Relationship Id="${escapeXmlAttribute(id)}" Type="${escapeXmlAttribute(type)}" Target="${escapeXmlAttribute(target)}"${mode}/>`)
		}
		lines.push('</Relationships>')
		return lines.join('')
	}
}
