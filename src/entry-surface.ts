/**
 * ts-pptx: the authoring surface shared by the three main entries
 *
 * `index.ts` (runtime-agnostic), `node.ts` and `browser.ts` publish exactly the same API. They
 * differ only in which {@link RuntimeAdapter} their `TsPptx` subclass is constructed with —
 * and, on the browser entry, one extra live-DOM method. Everything else is this module.
 *
 * The three used to carry a verbatim copy of the list below. Sharing it means adding to the
 * public surface is one edit rather than three, and the three cannot silently diverge.
 *
 * Note `export *` and not `export type *` on the barrels: `types/index.js` also exports the
 * `textRun`/`textRuns` run-array *values*. `export type *` would drop them and crash any
 * consumer that imports them, while TypeScript — reading `index.d.ts` — stayed green.
 */

export * from './enums.js'
export * from './units.js'
export * from './clip.js'
export * from './types/index.js'
export {
	resetDiagnosticState,
	setDiagnosticHandler,
	type Diagnostic,
	type DiagnosticCode,
	type DiagnosticHandler,
} from './diagnostics.js'
export * from './entry-errors.js'
