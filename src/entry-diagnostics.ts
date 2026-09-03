/**
 * ts-pptx: the diagnostic surface, as every entry point republishes it
 *
 * Every warning the library reports without throwing. The `code` is API; the message is not —
 * the same contract `entry-errors.ts` states for the other half of the `codes.ts` vocabulary.
 *
 * This module exists for the same reason that one does, and closes the same gap on the other
 * half: a consumer of `ts-pptx/read`, `ts-pptx/measure`, `ts-pptx/script`, `ts-pptx/inspect`,
 * `ts-pptx/html`, `ts-pptx/math` or `ts-pptx/zip` gets `console.warn` output from those paths
 * — a chart point cache out of range, a picture the reader cannot resolve, a table span the
 * pager refuses — with no supported way to intercept it, because the handler was published
 * only by the three authoring entries. Bundling happens to put `diagnostics.js` in a shared
 * chunk today, so importing the handler from `.` does currently take effect on `./read`; that
 * is an artifact of chunking rather than a promise, and it drags the whole write path in for a
 * three-line function.
 *
 * Entries re-export it wholesale (`export * from './entry-diagnostics.js'`), so the surface
 * cannot drift between them.
 */

export {
	resetDiagnosticState,
	setDiagnosticHandler,
	type Diagnostic,
	type DiagnosticCode,
	type DiagnosticHandler,
} from './diagnostics.js'
