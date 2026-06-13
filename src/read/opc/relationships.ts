import { escapeXmlAttribute, getElements, parseXml } from '../oxml/dom.js'
import { resolveRelativePartName } from './partnames.js'

export interface Relationship {
	id: string
	type: string
	target: string
	targetMode?: 'Internal' | 'External'
}

/**
 * Overlay over one `.rels` part, keyed by relationship id. While clean, the
 * `.rels` bytes pass through verbatim on save (byte-identical); once `add()`
 * (or another mutation) marks it dirty, `OpcPackage.save()` writes
 * `serialize()` into the owning `.rels` part instead.
 */
export class Relationships {
	/** Partname of the part owning these relationships ('/' for package-level). */
	readonly sourcePartName: string
	#byId = new Map<string, Relationship>()
	#dirty = false

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

	/** True once a relationship was added/changed; `serialize()` is then authoritative. */
	get isDirty(): boolean {
		return this.#dirty
	}

	/**
	 * Add a relationship and return it. Allocates an id `rId<n>` with `n` one
	 * past the highest existing numeric id. `target` is relative to the source
	 * part's directory for internal targets (e.g. `../media/image1.png`), or an
	 * absolute URI for external ones. Marks this set dirty.
	 */
	add(type: string, target: string, targetMode?: 'Internal' | 'External'): Relationship {
		let max = 0
		for (const id of this.#byId.keys()) {
			const match = /^rId(\d+)$/.exec(id)
			if (match) max = Math.max(max, Number(match[1]))
		}
		const id = `rId${max + 1}`
		const relationship: Relationship = { id, type, target }
		if (targetMode) relationship.targetMode = targetMode
		this.#byId.set(id, relationship)
		this.#dirty = true
		return relationship
	}

	/**
	 * Add a relationship under a caller-chosen id and return it. Used when
	 * rebuilding a copied part's relationships across a package boundary: keeping
	 * each source id lets the copied part body's `r:id`/`r:embed` references stay
	 * valid without rewriting the body. Throws if the id is already present.
	 * Marks this set dirty.
	 */
	addWithId(id: string, type: string, target: string, targetMode?: 'Internal' | 'External'): Relationship {
		if (this.#byId.has(id)) throw new Error(`Relationships of ${this.sourcePartName}: duplicate relationship id ${id}`)
		const relationship: Relationship = { id, type, target }
		if (targetMode) relationship.targetMode = targetMode
		this.#byId.set(id, relationship)
		this.#dirty = true
		return relationship
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
