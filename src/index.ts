import PresentationCore from './presentation.js'
import { createNeutralRuntime } from './runtime/neutral.js'

/**
 * The runtime-agnostic entry: what a consumer gets from the bare `@shbernal/ts-pptx` specifier
 * when neither the `node` nor the `browser` export condition resolves — Deno, Bun, edge workers.
 * Node and browser consumers reach `ts-pptx/node` and `ts-pptx/browser` through those conditions
 * without naming them, and get the same class backed by an adapter that can reach their host.
 *
 * Authoring is identical on all three. The difference is only where the finished deck can go:
 * `write`, `stream` and `toParts` hand the bytes back here as everywhere, while `writeFile`
 * throws `runtime/file-output-unavailable` because there is no filesystem and no DOM to write
 * to. Live-DOM `tableToSlides` is likewise absent — it is defined on the browser entry, and the
 * DOM-agnostic form is the free `tableToSlides` on `ts-pptx/html`.
 */
export class TsPptx extends PresentationCore {
	constructor() {
		super(createNeutralRuntime())
	}
}

export { TsPptx as default }
export * from './enums.js'
export * from './units.js'
// Use `export *` (not `export type *`) so the value exports `textRun`/`textRuns`
// reach this entry; `export type *` would drop them and crash any consumer that
// imports them, while TypeScript (reading index.d.ts) stays green.
export * from './types/index.js'
export { setDiagnosticHandler, type Diagnostic, type DiagnosticCode, type DiagnosticHandler } from './diagnostics.js'

// Error taxonomy — every failure the library throws. The classes and their `code` are API;
// the message is not. Re-exported from every entry so `instanceof` works whichever subpath a
// consumer imports — they all resolve to one shared module, so the classes are identical.
export {
	TsPptxError,
	InvalidOptionError,
	UnsupportedFeatureError,
	PackageReadError,
	MediaError,
	InternalError,
	type TsPptxErrorOptions,
} from './errors.js'
export type {
	ErrorCode,
	TsPptxCode,
	InvalidOptionErrorCode,
	UnsupportedFeatureErrorCode,
	PackageReadErrorCode,
	MediaErrorCode,
	InternalErrorCode,
} from './codes.js'
