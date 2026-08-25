/**
 * OOXML namespace URIs, by their canonical prefix.
 *
 * A namespace URI is a fact about the format, not about reading or writing one, so it belongs
 * here rather than on either side — the same reasoning `rel-types.ts` states at length, and for
 * the same failure mode: these are long, near-identical strings, and a typo does not throw. It
 * silently matches nothing on the read side and produces a part PowerPoint ignores on the write
 * side.
 *
 * The table lived in `read/oxml/dom.ts` and could not be reached from `gen/`, because that
 * module imports `@xmldom/xmldom` at module scope and pulling an XML DOM into the write-only
 * bundle is a real cost. So the write side hand-wrote the URIs instead: three separate `A_NS`
 * constants and seven copies of the same `xmlns:a`/`xmlns:r`/`xmlns:p` triple. This module has
 * **no runtime imports** — like `rel-types.ts`, `st-enums.ts` and `sequence.ts` beside it — so
 * both sides can reach it and neither pays for the other's dependencies. `read/oxml/dom.ts`
 * re-exports it, so no read-side import path changed.
 */

/** Canonical OOXML prefix → namespace URI registry. */
export const OOXML_NS = Object.freeze({
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
	cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
	ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
	cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
	dc: 'http://purl.org/dc/elements/1.1/',
	dcterms: 'http://purl.org/dc/terms/',
	dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
	ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
	p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
	p188: 'http://schemas.microsoft.com/office/powerpoint/2018/8/main',
	pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
})

/**
 * The `xmlns` declarations every PresentationML part root carries, in the order the emitters
 * have always written them.
 *
 * Spread this **first** in a root element's attribute map: `el()` serializes attributes in
 * object-key order, so the three keys have to keep the positions the literals occupied or the
 * emitted bytes move. A root with extra attributes (`show` on `p:sld`, `def` on `a:tblStyleLst`)
 * lists them after the spread, which is where they already were.
 */
export const PML_ROOT_NS = Object.freeze({
	'xmlns:a': OOXML_NS.a,
	'xmlns:r': OOXML_NS.r,
	'xmlns:p': OOXML_NS.p,
})
