/**
 * The library's condition vocabulary — one registry covering both surfaces a problem can take.
 *
 * A **code** names a *condition*, in `area/condition` form. The same condition keeps the same code
 * whichever way it reaches the consumer: as a non-fatal {@link Diagnostic} (see `diagnostics.ts`)
 * or as a thrown {@link TsPptxError} (see `errors.ts`). `coord/non-finite` means "a coordinate was
 * `NaN`/`Infinity`" in both directions, and a consumer that special-cases it only has to learn the
 * one string.
 *
 * **A code is API; the message that carries it is not.** Adding a code is back-compatible;
 * removing or renaming one is a breaking change. Every union here is closed, so a new warning or
 * throw site cannot be written without naming its condition in this file — that is the enforcement
 * mechanism for keeping the vocabulary curated rather than accumulated. Reuse an existing code when
 * the condition is genuinely the same, even if the wording differs and even if it is reported from
 * a different entry point.
 *
 * The error codes are grouped by the {@link TsPptxError} subclass that carries them, not by area,
 * because that pairing is type-enforced: `new MediaError('coord/non-finite', …)` does not compile.
 * A condition that can genuinely arrive under two different classes needs two codes, because the
 * class is part of what the consumer branches on.
 *
 * This module is types only — it emits no runtime code, so importing it can never create an import
 * cycle and `diagnostics.ts` keeps the "safe to use from anywhere" property its header claims.
 */

/**
 * The stable identifier for a non-fatal condition the library reports and then works around.
 *
 * Codes are shared across entry points where the *condition* is the same even though the message
 * differs — an image whose natural size cannot be measured reports
 * `image/unmeasurable-natural-size` whether it came from `addImage` or from `Picture.setImage`.
 */
export type DiagnosticCode =
	// Selection Pane identity (`objectName`)
	| 'object-name/empty'
	| 'object-name/control-characters'
	| 'object-name/too-long'
	| 'object-name/duplicate'
	// Build animations
	| 'animation/target-index-out-of-range'
	| 'animation/target-not-found'
	| 'animation/target-missing'
	// Charts
	| 'chart/invalid-axis-time-unit'
	| 'chart/non-finite-value'
	| 'chart/error-bars-missing-values'
	| 'chart/axis-type-conflict'
	| 'chart/invalid-metadata'
	| 'chart/invalid-metadata-key'
	| 'chart/invalid-metadata-value'
	| 'chart/option-out-of-range'
	| 'chart/invalid-subtotal-index'
	| 'chart/symbol-size-out-of-range'
	| 'chart/invalid-grid-line-size'
	| 'chart/invalid-grid-line-style'
	| 'chart/invalid-grid-line-cap'
	| 'chart/layout-out-of-range'
	| 'chart/stock-series-count'
	// Comments
	| 'comment/missing-author'
	| 'comment/missing-text'
	// Connectors
	| 'connector/bends-ignored-for-straight'
	| 'connector/adj-out-of-range'
	| 'connector/unresolved-binding'
	// Groups
	| 'group/unsupported-child'
	| 'group/unrecognized-child'
	| 'group/no-children'
	| 'group/partial-frame'
	// Images and image fills
	| 'image-fill/missing-source'
	| 'image-fill/missing-base64-header'
	| 'image-fill/svg-unsupported'
	| 'image-fill/unresolved-media'
	| 'image/crop-and-sizing-conflict'
	| 'image/unmeasurable-natural-size'
	// Tables
	| 'table/invalid-border'
	| 'table/col-width-count-mismatch'
	| 'table/autopage-height-too-small'
	// Text and bullets
	| 'text/invalid-columns'
	| 'text/invalid-column-spacing'
	| 'text/invalid-fit-percentage'
	| 'text/char-spacing-out-of-range'
	| 'text/line-spacing-out-of-range'
	| 'font/size-out-of-range'
	| 'bullet/size-out-of-range'
	| 'bullet/image-embed-failed'
	| 'bullet/invalid-character-code'
	// Zooms
	| 'zoom/missing-target'
	| 'zoom/unresolved-target'
	| 'zoom/missing-section-title'
	| 'zoom/section-not-found'
	| 'zoom/section-empty'
	| 'zoom/no-sections-to-summarize'
	// Colour, shadow, geometry, locks
	| 'color/not-a-string'
	| 'color/invalid-value'
	| 'shadow/invalid-type'
	| 'shadow/angle-out-of-range'
	| 'shadow/transparency-out-of-range'
	| 'shadow/color-has-hash'
	| 'geometry/invalid-shape-adjust'
	| 'geometry/shape-adjust-overridden'
	| 'geometry/arc-node-point-ignored'
	| 'geometry/invalid-guide'
	| 'geometry/unknown-guide-operation'
	| 'geometry/invalid-connection-site'
	| 'object-lock/unsupported-on-shape'
	// Media, theme, masters, notes
	| 'media/load-failed'
	| 'theme/invalid-color-override'
	| 'master/invalid-text-style-font-size'
	| 'master/too-many-text-style-levels'
	| 'notes/hyperlink-slide-unsupported'
	// HTML table conversion
	| 'html/image-missing-source'
	// Deck-level authoring
	| 'section/missing-argument'
	| 'section/missing-title'
	| 'section/duplicate-title'
	| 'slide/section-not-found'
	| 'layout/invalid-definition'
	// Reading and measuring
	| 'inspect/group-transform-missing'
	| 'inspect/group-transform-degenerate'
	| 'measure/shrink-unmeasured'
	| 'measure/resize-unmeasured'
	| 'measure/heuristic-metrics'
	// Units and coordinates
	| 'margin/legacy-points'
	| 'transparency/out-of-range'
	| 'opacity/out-of-range'
	| 'line/width-out-of-range'
	| 'coord/bare-number-is-inches'

/**
 * Conditions carried by `InvalidOptionError`: the caller passed something the library cannot use.
 *
 * This is the largest group by design — the project's stated policy is to throw rather than coerce
 * (`AGENTS.md` → "Silent coercion of invalid input is a footgun"), so most throws are here.
 */
export type InvalidOptionErrorCode =
	// Units and coordinates
	| 'coord/non-finite'
	| 'coord/not-positive'
	| 'coord/invalid-format'
	// Geometry
	| 'geometry/arc-angle-non-finite'
	// Charts
	| 'chart/missing-type'
	// Embedded fonts
	| 'font/missing-typeface'
	| 'font/invalid-style-slot'
	| 'font/invalid-base64'
	| 'font/missing-source'
	// Deck-level authoring
	| 'layout/unknown'
	| 'master/missing-title'
	| 'table-style/missing-argument'
	| 'table-style/missing-name'
	// Math (LaTeX)
	| 'math/invalid-latex'
	// Zip / package I/O
	| 'zip/unsupported-input'
	| 'zip/unsupported-output'

/**
 * Conditions carried by `UnsupportedFeatureError`: the request is well-formed, but this build,
 * runtime, or shape cannot express it. Distinct from `InvalidOptionError` because the caller did
 * nothing wrong — the answer is "not here", not "not like that".
 */
export type UnsupportedFeatureErrorCode =
	// Optional dependencies that were not installed or did not load
	| 'math/missing-optional-peer'
	| 'font/opentype-unavailable'
	// Runtime capabilities
	| 'zip/filesystem-unavailable'

/**
 * Conditions carried by `PackageReadError`: the bytes handed to the library are not a package it
 * can read, or a part inside one is structurally malformed. Always about *input*, never about
 * something the library is asked to produce.
 */
export type PackageReadErrorCode = 'zip/not-a-zip-archive' | 'zip/file-read-failed'

/**
 * Conditions carried by `MediaError`: an image, font, or audio/video payload could not be fetched,
 * read, or decoded. Kept separate from `PackageReadError` because the failure is in a *referenced
 * resource*, not in the package structure, and the recovery is usually different (fix the URL or
 * the file, not the deck).
 */
export type MediaErrorCode =
	| 'font/fetch-failed'
	| 'font/read-failed'
	| 'media/fetch-failed'
	| 'media/read-failed'
	| 'media/decode-failed'
	| 'media/svg-preview-failed'

/**
 * Conditions carried by `InternalError`: an invariant the library maintains itself did not hold.
 * Reaching one is a bug in ts-pptx, not something a consumer can fix by changing their input —
 * which is exactly why it is worth being able to tell apart from the four classes above.
 */
export type InternalErrorCode = 'layout/default-not-registered'

/** Every condition the library can throw. See {@link TsPptxCode} for the whole vocabulary. */
export type ErrorCode =
	InvalidOptionErrorCode | UnsupportedFeatureErrorCode | PackageReadErrorCode | MediaErrorCode | InternalErrorCode

/** Every condition the library can report, warned or thrown. */
export type TsPptxCode = DiagnosticCode | ErrorCode
