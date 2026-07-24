/**
 * ts-pptx: XML element builder (write-side substrate)
 *
 * The write path historically concatenated template strings and called
 * `encodeXmlEntities` by hand at every interpolation — the source of escaping,
 * attribute-order and child-sequence bugs. This builder centralizes escaping
 * while reproducing today's exact byte layout, so a migration onto it can be
 * gated on strict byte-identity (`pnpm run byte-identity:check`) rather than on
 * looser schema equivalence.
 *
 * Mirrors `src/read/oxml/dom.ts` on the write side.
 *
 * Two rules make byte-identity achievable:
 *
 * 1. Self-closing is decided by ARITY, never by child value. `el()` always emits
 *    a paired tag; `voidEl()` always self-closes. This matters because
 *    `encodeXmlEntities(undefined)` returns `''`, so a value-based rule would
 *    turn today's `<dc:title></dc:title>` into `<dc:title/>` — a silent byte
 *    regression on every optional-text element in the tree.
 *
 * 2. Whitespace is explicit, per element, via `fmt`. Most emitted parts are flat
 *    (no `fmt` needed at all); the pretty-printed ones carry indentation that is
 *    not always depth-regular, so it is described rather than derived. Quirks
 *    like a misindented closing tag stay visible in the call site instead of
 *    being hidden inside a template literal.
 */

import { encodeXmlEntities } from '../../gen-utils.js'

/** Pre-serialized XML, interpolated verbatim (not escaped). */
export interface RawXml {
	raw: string
}

/** A text child (escaped) or pre-serialized markup (verbatim). */
export type XmlChild = string | number | RawXml | null | undefined

/** Attribute values; `undefined`/`null` omit the attribute entirely. */
export type XmlAttrs = Record<string, string | number | null | undefined>

/**
 * Byte-layout control. Omit for flat output.
 * - `openPrefix` — emitted before `<name`
 * - `childPrefix` — emitted before each child
 * - `closePrefix` — emitted before the closing delimiter (`</name>`, or `/>` for `voidEl`)
 */
export interface XmlFmt {
	openPrefix?: string
	childPrefix?: string
	closePrefix?: string
}

/** Wrap already-serialized XML so it is interpolated verbatim. */
export function raw(xml: string): RawXml {
	return { raw: xml }
}

function isRaw(child: XmlChild): child is RawXml {
	return typeof child === 'object' && child !== null && typeof child.raw === 'string'
}

function openTag(name: string, attrs: XmlAttrs | null | undefined, fmt: XmlFmt | undefined): string {
	let out = (fmt?.openPrefix ?? '') + '<' + name
	for (const [key, value] of Object.entries(attrs ?? {})) {
		if (value === undefined || value === null) continue
		out += ' ' + key + '="' + encodeXmlEntities(value) + '"'
	}
	return out
}

/**
 * Paired element: `<name …>children</name>` — emitted even when `children` is
 * empty or `undefined`. Text children are escaped; `raw()` children are not.
 */
export function el(name: string, attrs?: XmlAttrs | null, children?: XmlChild | XmlChild[], fmt?: XmlFmt): string {
	let out = openTag(name, attrs, fmt) + '>'
	const childPrefix = fmt?.childPrefix ?? ''
	for (const child of Array.isArray(children) ? children : [children]) {
		if (child === undefined || child === null) continue
		out += childPrefix + (isRaw(child) ? child.raw : encodeXmlEntities(child))
	}
	return out + (fmt?.closePrefix ?? '') + '</' + name + '>'
}

/**
 * Self-closing element: `<name …/>`. `fmt.closePrefix` goes before the `/>` — a
 * handful of emitters write `<a:avLst />` with a space there, and that space is
 * byte-significant like any other.
 */
export function voidEl(name: string, attrs?: XmlAttrs | null, fmt?: XmlFmt): string {
	return openTag(name, attrs, fmt) + (fmt?.closePrefix ?? '') + '/>'
}
