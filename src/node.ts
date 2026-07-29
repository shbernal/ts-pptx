import PresentationCore from './pptxgen.js'
import { createNodeRuntime } from './runtime/node.js'

export class TsPptx extends PresentationCore {
	constructor() {
		super(createNodeRuntime())
	}
}

export { TsPptx as default }
export * from './core-enums.js'
export * from './units.js'
// Use `export *` (not `export type *`) so the value exports `textRun`/`textRuns`
// reach this entry; `export type *` would drop them and crash any Node consumer
// that imports them, while TypeScript (reading index.d.ts) stays green.
export * from './core-interfaces.js'
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
