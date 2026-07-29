/**
 * Read a deck's document properties — the `docProps/core.xml` core properties
 * (Dublin Core title/subject/creator + OPC keywords/revision/timestamps) and the
 * `docProps/custom.xml` user-defined name/value pairs. These are the read
 * counterparts of the write-side `pptx.title`/`subject`/`author`/`revision`
 * setters (`makeXmlCore`) and `pptx.setCustomProperty(...)` (`makeXmlCustomProperties`),
 * so a reader is verified by a genuine write→read round-trip.
 *
 * Both parts are declared at the package root (`/_rels/.rels`): core-properties
 * via the OPC `metadata/core-properties` rel, custom-properties via the
 * officeDocument `custom-properties` rel.
 */
import type { CustomPropertyValue } from '../../types/index.js'
import { OpcPackage } from '../opc/package.js'
import { attr, childElements, firstChild, firstChildElement, intValue, type Element } from '../oxml/dom.js'

/** Package-root rel to `/docProps/core.xml`. */
const CORE_PROPS_REL_TYPE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
/** Content type of the core-properties part (fallback lookup when the rel is absent). */
const CORE_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'
/** Package-root rel to `/docProps/custom.xml`. */
const CUSTOM_PROPS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
/** Content type of the custom-properties part (fallback lookup when the rel is absent). */
const CUSTOM_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'

/**
 * Decoded `docProps/core.xml`. Every field is optional (present only when the
 * element is). `created`/`modified`/`lastPrinted` are kept as the **raw W3CDTF
 * string** rather than parsed to `Date`, avoiding timezone round-trip loss and
 * matching the byte-preserve ethos of the read model.
 */
export interface CoreProperties {
	/** `dc:title`. */
	title?: string
	/** `dc:subject`. */
	subject?: string
	/** `dc:creator` (the write-side `pptx.author`). */
	creator?: string
	/** `cp:keywords`. */
	keywords?: string
	/** `dc:description`. */
	description?: string
	/** `cp:lastModifiedBy`. */
	lastModifiedBy?: string
	/** `cp:revision`. */
	revision?: string
	/** `cp:category`. */
	category?: string
	/** `cp:contentStatus`. */
	contentStatus?: string
	/** `dcterms:created`, raw W3CDTF string. */
	created?: string
	/** `dcterms:modified`, raw W3CDTF string. */
	modified?: string
	/** `cp:lastPrinted`, raw W3CDTF string. */
	lastPrinted?: string
}

/** One user-defined custom document property (`docProps/custom.xml`). */
export interface CustomProperty {
	/** The property's `@name`. */
	name: string
	/** The decoded value, typed from its `vt:` child element. */
	value: CustomPropertyValue
}

/** `dc:title`-style core-property fields, keyed by output field → element qname. */
const CORE_FIELDS: ReadonlyArray<[keyof CoreProperties, string]> = [
	['title', 'dc:title'],
	['subject', 'dc:subject'],
	['creator', 'dc:creator'],
	['keywords', 'cp:keywords'],
	['description', 'dc:description'],
	['lastModifiedBy', 'cp:lastModifiedBy'],
	['revision', 'cp:revision'],
	['category', 'cp:category'],
	['contentStatus', 'cp:contentStatus'],
	['created', 'dcterms:created'],
	['modified', 'dcterms:modified'],
	['lastPrinted', 'cp:lastPrinted'],
]

/**
 * Read the deck's core document properties from `/docProps/core.xml`, resolved
 * via the package-root `core-properties` relationship (fallback: the part's
 * content type). Missing part → `{}` (all fields undefined). Present-but-empty
 * elements decode to the empty string.
 */
export function readCoreProperties(opc: OpcPackage): CoreProperties {
	const rels = opc.relationshipsFor('/')
	const rel = rels.byType(CORE_PROPS_REL_TYPE)[0]
	const part = rel ? opc.part(rels.resolveTarget(rel.id)) : opc.partsByContentType(CORE_PROPS_CONTENT_TYPE)[0]
	const root = part?.dom.documentElement
	if (!root) return {}
	const out: CoreProperties = {}
	for (const [field, qname] of CORE_FIELDS) {
		const el = firstChild(root, qname)
		if (el) out[field] = el.textContent ?? ''
	}
	return out
}

/**
 * Decode one `<property>`'s single `vt:` child into a typed value. String types
 * (`lpwstr`/`lpstr`/`bstr`) → string; integer types (`i1`..`i8`/`int`/`ui*`) →
 * number; real types (`r4`/`r8`/`decimal`) → number; `bool` → boolean;
 * `filetime`/`date` → the **raw string** (see {@link CoreProperties} on
 * timestamps). Any unrecognized `vt:` element falls back to its text content as a
 * string — lossy but never throws.
 */
function decodeValue(vt: Element): CustomPropertyValue {
	const text = vt.textContent ?? ''
	switch (vt.localName) {
		case 'lpwstr':
		case 'lpstr':
		case 'bstr':
			return text
		case 'i1':
		case 'i2':
		case 'i4':
		case 'i8':
		case 'int':
		case 'ui1':
		case 'ui2':
		case 'ui4':
		case 'ui8':
		case 'uint':
			return intValue(text) ?? text
		case 'r4':
		case 'r8':
		case 'decimal': {
			const n = Number.parseFloat(text)
			return Number.isFinite(n) ? n : text
		}
		case 'bool':
			return text.trim().toLowerCase() === 'true' || text.trim() === '1'
		case 'filetime':
		case 'date':
			return text
		default:
			// Unknown vt type (vt:array, vt:cy, …): keep the raw text, lossy but safe.
			return text
	}
}

/**
 * Read the deck's user-defined custom properties from `/docProps/custom.xml`,
 * resolved via the package-root `custom-properties` relationship (fallback: the
 * part's content type). Each `<property name=...>` maps to `{ name, value }`,
 * with `value` typed from its single `vt:` child. Missing part → `[]`. A
 * `<property>` without a `@name` or without a `vt:` child is skipped.
 *
 * The `vt:` child and `<property>` element sit in namespaces outside the read
 * model's qname registry, so children are matched by local name rather than the
 * `firstChild`/`getElements` helpers.
 */
export function readCustomProperties(opc: OpcPackage): CustomProperty[] {
	const rels = opc.relationshipsFor('/')
	const rel = rels.byType(CUSTOM_PROPS_REL_TYPE)[0]
	const part = rel ? opc.part(rels.resolveTarget(rel.id)) : opc.partsByContentType(CUSTOM_PROPS_CONTENT_TYPE)[0]
	const root = part?.dom.documentElement
	if (!root) return []
	const out: CustomProperty[] = []
	for (const property of childElements(root)) {
		if (property.localName !== 'property') continue
		const name = attr(property, 'name')
		const vt = firstChildElement(property)
		if (name === null || !vt) continue
		out.push({ name, value: decodeValue(vt) })
	}
	return out
}
