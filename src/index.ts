export { default, TsPptx } from './browser.js'
export * from './enums.js'
export * from './units.js'
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
