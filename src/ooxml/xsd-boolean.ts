/**
 * Parsing an `xsd:boolean` attribute value.
 *
 * **Why a function at all.** `xsd:boolean` has four lexical forms — `1`, `0`, `true`, `false` —
 * and OOXML attributes typed with it accept every one of them. PowerPoint writes only `1` and
 * `0`, so a hand-rolled `=== '1'` test passes against every deck this library produces and
 * against every deck PowerPoint produces, and silently misreads the other producers. That is the
 * worst shape a bug can have: invisible to the fixtures, visible only on a customer's file.
 *
 * **Why this module is neither `gen/` nor `read/`.** The lexical space of a schema type is a fact
 * about the schema, like `sequence.ts`'s document order and `body-insets.ts`'s defaults. The
 * parser lived in `read/oxml/dom.ts` because reading is where most of the demand is, and three
 * private copies had collapsed into it — but `read/oxml/dom.ts` imports `@xmldom/xmldom`, so
 * anything outside `read/` that wants this five-line function has to take a DOM dependency to get
 * it, or write a fourth copy. `script/from-read/detect.ts` hit exactly that and wrote the copy.
 *
 * The DOM-typed convenience wrapper stays behind in `read/oxml/dom.ts` as `boolAttr`, because
 * *that* one really is about reading an element.
 */

/**
 * Parse an `xsd:boolean` OOXML attribute (`1`/`0`/`true`/`false`); else `null`.
 *
 * `null` for an absent *or* unparseable value is deliberate and load-bearing at the call sites:
 * several attributes have a schema default of `true`, so they test `!== false` rather than
 * `=== true` and an unset attribute has to be distinguishable from an explicit `0`.
 */
export function boolValue(value: string | null | undefined): boolean | null {
	if (value === '1' || value === 'true') return true
	if (value === '0' || value === 'false') return false
	return null
}
