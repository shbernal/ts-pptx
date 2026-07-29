/**
 * The library's error taxonomy.
 *
 * Every failure ts-pptx raises is a {@link TsPptxError} carrying a stable {@link ErrorCode}, so a
 * consumer can tell *"you passed a bad coordinate"* from *"this font file is corrupt"* from
 * *"these bytes are not a package"* without matching on message substrings. The classes are a
 * deliberately flat, coarse bucket — five of them — and the code carries the specificity. Do not
 * add a class per throw site; add a code.
 *
 * **The class and the code are API. The message is not.** Branch on `instanceof` and on `code`;
 * the wording behind them is free to improve in any release. Every error remains an
 * `instanceof Error`, so existing `catch` blocks keep working.
 *
 * Which class to reach for:
 *
 * | class | the failure is | who fixes it |
 * |---|---|---|
 * | {@link InvalidOptionError} | the caller passed something unusable | the caller's code |
 * | {@link UnsupportedFeatureError} | a well-formed request this build/runtime/shape cannot express | the caller's expectations, or the environment |
 * | {@link PackageReadError} | the input bytes are not a readable package | the input file |
 * | {@link MediaError} | a referenced image/font/AV resource would not load or decode | the resource |
 * | {@link InternalError} | an invariant of the library itself did not hold | ts-pptx (file a bug) |
 *
 * @see `docs/errors.md` for the consumer-facing contract.
 * @see `codes.ts` for the code vocabulary, which is shared with the diagnostic (warning) surface.
 */

import type {
	ErrorCode,
	InternalErrorCode,
	InvalidOptionErrorCode,
	MediaErrorCode,
	PackageReadErrorCode,
	UnsupportedFeatureErrorCode,
} from './codes.js'

/** Extra context accepted by every {@link TsPptxError} constructor. */
export interface TsPptxErrorOptions extends ErrorOptions {
	/**
	 * Structured context for the condition, when a site has any to give. Mirrors `Diagnostic.detail`
	 * so the same condition looks the same whether it was warned or thrown.
	 */
	readonly detail?: Readonly<Record<string, unknown>>
}

/**
 * Base class for every error ts-pptx throws.
 *
 * Catch this to catch anything from the library; narrow with `instanceof` on a subclass, or branch
 * on {@link TsPptxError.code}, to react to something specific. It is never thrown directly — every
 * site picks one of the five subclasses.
 */
export class TsPptxError extends Error {
	/** Stable condition identifier — see {@link ErrorCode}. Branch on this, not on the text. */
	readonly code: ErrorCode

	/** Structured context for the condition, when the throwing site had any to give. */
	readonly detail?: Readonly<Record<string, unknown>>

	/**
	 * @param code - the stable condition identifier
	 * @param message - explanation for a human reading a stack trace; not part of the contract
	 * @param options - standard `cause`, plus optional structured `detail`
	 */
	constructor(code: ErrorCode, message: string, options?: TsPptxErrorOptions) {
		super(message, options)
		// `name` drives how a stack trace and `console.log` label the error, and the default from
		// `Error` would say "Error" for every subclass. Derive it from the constructor so a subclass
		// never has to remember to set it.
		this.name = new.target.name
		this.code = code
		if (options?.detail !== undefined) this.detail = options.detail
	}
}

/**
 * The caller passed something the library cannot use — a non-finite coordinate, an unknown enum
 * value, a required option left out, a value outside its legal range.
 *
 * This is the error behind the project's throw-rather-than-coerce policy: emitting a degenerate
 * result (a zero-size shape, a silently dropped option) hides the mistake, so the library refuses
 * instead. Reaching one means the calling code needs to change.
 */
export class InvalidOptionError extends TsPptxError {
	declare readonly code: InvalidOptionErrorCode
	constructor(code: InvalidOptionErrorCode, message: string, options?: TsPptxErrorOptions) {
		super(code, message, options)
	}
}

/**
 * The request is well-formed but cannot be expressed here: an optional peer dependency is not
 * installed, the runtime lacks a capability, or the shape/feature has no OOXML representation the
 * library emits.
 *
 * Distinct from {@link InvalidOptionError} because the caller did nothing wrong — the answer is
 * "not here", not "not like that" — and the fix is usually to install something or pick a different
 * approach rather than to correct a value.
 */
export class UnsupportedFeatureError extends TsPptxError {
	declare readonly code: UnsupportedFeatureErrorCode
	constructor(code: UnsupportedFeatureErrorCode, message: string, options?: TsPptxErrorOptions) {
		super(code, message, options)
	}
}

/**
 * The bytes handed to the library are not a package it can read, or a part inside one is
 * structurally malformed — a missing `[Content_Types].xml`, a relationship pointing at nothing, a
 * slide part with no root element.
 *
 * Always about *input*. A problem with something the library is being asked to produce is an
 * {@link InvalidOptionError}.
 */
export class PackageReadError extends TsPptxError {
	declare readonly code: PackageReadErrorCode
	constructor(code: PackageReadErrorCode, message: string, options?: TsPptxErrorOptions) {
		super(code, message, options)
	}
}

/**
 * A referenced image, font, or audio/video resource could not be fetched, read, or decoded.
 *
 * Kept apart from {@link PackageReadError} because the failure is in a resource the deck *points
 * at*, not in the package structure — a broken URL, an unreadable path, a corrupt font file — and
 * a consumer typically retries or substitutes rather than rejecting the deck.
 */
export class MediaError extends TsPptxError {
	declare readonly code: MediaErrorCode
	constructor(code: MediaErrorCode, message: string, options?: TsPptxErrorOptions) {
		super(code, message, options)
	}
}

/**
 * An invariant the library maintains itself did not hold.
 *
 * No consumer input should be able to produce one, which is the whole reason it is a separate
 * class: seeing it means the bug is in ts-pptx, and the useful response is to file it rather than
 * to keep adjusting the input.
 */
export class InternalError extends TsPptxError {
	declare readonly code: InternalErrorCode
	constructor(code: InternalErrorCode, message: string, options?: TsPptxErrorOptions) {
		super(code, message, options)
	}
}
