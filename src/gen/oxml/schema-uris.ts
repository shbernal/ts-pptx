/**
 * ts-pptx: OPC schema URI bases (write-side substrate)
 *
 * Six writers under `src/gen/` each declared their own private copy of the same two or three
 * `http://schemas.openxmlformats.org/…` prefixes. They are long, near-identical strings that
 * differ in one path segment, so a typo is easy to write and invisible on review — a wrong URI
 * does not throw, it produces a part PowerPoint silently ignores or rejects at open time.
 *
 * Only the *bases* live here. The rel type a single writer cares about (`oleObject`, `chartEx`,
 * the MS media rel, …) stays declared next to the code that emits it: those are read at the call
 * site, and hoisting them would trade a definition you can see for one you have to go and find.
 *
 * `src/read/api/rel-types.ts` names some of the same URIs for the read model. Sharing one home
 * across the gen/read boundary is a separate decision; this module is deliberately gen-internal.
 */

/** Root of every ECMA-376 schema URI. Private: callers want one of the two named prefixes below. */
const SCHEMA_BASE = 'http://schemas.openxmlformats.org/'

/** Prefix for `.rels` `Type` URIs — append the type name (`slide`, `image`, `theme`, …). */
export const OFFICE_REL = SCHEMA_BASE + 'officeDocument/2006/relationships/'

/** `xmlns` of every `<Relationships>` part. Also the prefix of the package-scoped rel types. */
export const PACKAGE_REL_NS = SCHEMA_BASE + 'package/2006/relationships'
