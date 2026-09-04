/**
 * Read a deck's document properties — the `docProps/core.xml` core properties
 * (Dublin Core title/subject/creator + OPC keywords/revision/timestamps), the
 * `docProps/app.xml` extended properties, and the `docProps/custom.xml`
 * user-defined name/value pairs. These are the read counterparts of the write-side
 * `pptx.title`/`subject`/`author`/`revision` setters (`makeXmlCore`), `pptx.company`
 * (`makeXmlApp`) and `pptx.setCustomProperty(...)` (`makeXmlCustomProperties`),
 * so a reader is verified by a genuine write→read round-trip.
 *
 * All three parts are declared at the package root (`/_rels/.rels`): core-properties
 * via the OPC `metadata/core-properties` rel, extended-properties and
 * custom-properties via the officeDocument `extended-properties` /
 * `custom-properties` rels.
 */
import type { CustomPropertyValue } from '../../types/index.js'
import { OpcPackage } from '../opc/package.js'
import { singleRelPart } from '../opc/partnames.js'
import { attr, childElements, firstChild, firstChildElement, numberValue, type Element } from '../oxml/dom.js'
import { CORE_PROPS_REL, CUSTOM_PROPS_REL, EXTENDED_PROPS_REL } from '../../ooxml/rel-types.js'
import { boolValue } from '../../ooxml/xsd-boolean.js'

// The two content types below are the fallback lookup for when the rel is absent, and this is
// the only module that matches on them — the write side spells its own out next to the part it
// declares (`gen/opc/content-types.ts`). The rel types themselves are shared with that side.
/** Content type of the core-properties part. */
const CORE_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'
/** Content type of the extended-properties part. */
const EXTENDED_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
/** Content type of the custom-properties part. */
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

/**
 * Decoded `docProps/app.xml` — the *extended* properties, which are the producer's own
 * account of the deck rather than the author's metadata. Every field is optional (present
 * only when the element is).
 *
 * Deliberately a subset. `<Slides>`, `<Words>`, `<Paragraphs>` and the rest are statistics
 * the producing application computed for the file it wrote; reporting them from a read model
 * that can hand back an edited deck would be reporting a number about a document that no
 * longer exists. The four here identify the producer and name the deck's parts, which stay
 * true, and `company` is the one this library's write API can set.
 */
export interface ExtendedProperties {
	/** `<Application>` — the producing application, e.g. `Microsoft Office PowerPoint`. */
	application?: string
	/** `<AppVersion>` — that application's version, in its own `major.minor` spelling (`16.0000`). */
	appVersion?: string
	/** `<Company>` (the write-side `pptx.company`). */
	company?: string
	/**
	 * `<TitlesOfParts>` — the flat `vt:lpstr` vector naming the deck's fonts, themes and slide
	 * titles in one list. It is the vector as written, NOT split by section: `<HeadingPairs>`
	 * holds the counts that partition it, and this library does not read them, so the caller
	 * that wants the slide titles alone has to pair the two itself.
	 */
	titlesOfParts?: string[]
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
	const part = singleRelPart(opc, '/', CORE_PROPS_REL) ?? opc.partsByContentType(CORE_PROPS_CONTENT_TYPE)[0]
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
 * Read the deck's extended document properties from `/docProps/app.xml`, resolved
 * via the package-root `extended-properties` relationship (fallback: the part's
 * content type). Missing part → `{}` (all fields undefined). Present-but-empty
 * elements decode to the empty string.
 *
 * Children are matched by **local name**: the extended-properties namespace is not in the
 * read model's qname registry, and its elements are unprefixed there, so this reads them the
 * way {@link readCustomProperties} reads a `<property>`'s `vt:` child.
 */
export function readExtendedProperties(opc: OpcPackage): ExtendedProperties {
	const part = singleRelPart(opc, '/', EXTENDED_PROPS_REL) ?? opc.partsByContentType(EXTENDED_PROPS_CONTENT_TYPE)[0]
	const root = part?.dom.documentElement
	if (!root) return {}
	const out: ExtendedProperties = {}
	for (const child of childElements(root)) {
		switch (child.localName) {
			case 'Application':
				out.application = child.textContent ?? ''
				break
			case 'AppVersion':
				out.appVersion = child.textContent ?? ''
				break
			case 'Company':
				out.company = child.textContent ?? ''
				break
			case 'TitlesOfParts': {
				// `<TitlesOfParts>` wraps exactly one `<vt:vector>`; the entries are its `vt:lpstr`
				// children. An empty vector is still a stated one, so `[]` is reported rather than
				// the key being left off.
				const vector = firstChildElement(child)
				if (vector) out.titlesOfParts = childElements(vector).map((item) => item.textContent ?? '')
				break
			}
			default:
				break
		}
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
			return numberValue(text) ?? text
		case 'r4':
		case 'r8':
		case 'decimal': {
			const n = Number.parseFloat(text)
			return Number.isFinite(n) ? n : text
		}
		case 'bool':
			// `vt:bool` is `xsd:boolean`, so `boolValue` is the one reading of its four lexical
			// forms. The hand-rolled test accepted `'TRUE'`, which `xsd:boolean` does not.
			return boolValue(text.trim()) ?? false
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
	const part = singleRelPart(opc, '/', CUSTOM_PROPS_REL) ?? opc.partsByContentType(CUSTOM_PROPS_CONTENT_TYPE)[0]
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
