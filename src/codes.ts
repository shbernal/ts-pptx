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
	| 'table/invalid-outer-border'
	| 'table/invalid-horz-overflow'
	| 'table/invalid-cell3d'
	| 'table/col-width-count-mismatch'
	| 'table/autopage-height-too-small'
	| 'table-style/region-overridden'
	// Text and bullets
	| 'text/invalid-columns'
	| 'text/invalid-column-spacing'
	| 'text/invalid-fit-percentage'
	| 'text/char-spacing-out-of-range'
	| 'text/line-spacing-out-of-range'
	| 'font/size-out-of-range'
	| 'bullet/size-out-of-range'
	| 'bullet/image-missing-base64-header'
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
	| 'border/unknown-key'
	| 'border/invalid-dash-type'
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
	| 'coord/negative'
	| 'coord/invalid-format'
	// Geometry
	| 'geometry/arc-angle-non-finite'
	// Charts
	| 'chart/missing-type'
	| 'chart/secondary-axis-unused'
	| 'chart/axis-count-mismatch'
	// Shapes and connectors
	| 'shape/missing-type'
	| 'shape/unknown-preset'
	| 'connector/missing-endpoints'
	| 'connector/invalid-type'
	| 'connector/invalid-bends'
	| 'connector/adj-count-mismatch'
	| 'connector/adj-non-finite'
	| 'connector/invalid-binding-name'
	| 'connector/invalid-connection-site'
	| 'line/invalid-cap'
	// Groups
	| 'group/missing-object-names'
	| 'group/invalid-object-name'
	| 'group/duplicate-object-name'
	| 'group/unresolved-object-name'
	| 'group/ambiguous-object-name'
	// Fills
	| 'gradient/angle-non-finite'
	| 'gradient/too-few-stops'
	| 'gradient/stop-position-non-finite'
	| 'gradient/stop-position-out-of-range'
	| 'gradient/rotate-with-shape-not-boolean'
	| 'gradient/scaled-not-boolean'
	| 'pattern-fill/missing-pattern'
	// Images
	| 'image/missing-source'
	| 'image/path-not-a-string'
	| 'image/data-not-a-string'
	| 'image/missing-base64-header'
	| 'image/crop-window-overflows'
	| 'image/crop-inset-out-of-range'
	| 'image/crop-insets-exceed-extent'
	// Media and OLE
	| 'media/missing-source'
	| 'media/missing-base64-header'
	| 'media/cover-missing-base64-header'
	| 'media/online-missing-link'
	| 'ole/missing-source'
	// Tables
	| 'table/rows-not-an-array'
	| 'table/rows-not-nested'
	// Editing a table in an existing deck (`ts-pptx/read`). Each names an attribute value
	// outside its schema enum, or a measurement that cannot be written — conditions the read
	// path throws on rather than drops, because a caller editing one attribute would
	// otherwise be left looking at an unchanged deck with nothing to explain it.
	| 'table/invalid-cell-anchor'
	| 'table/invalid-cell-vert'
	| 'table/invalid-cell-overflow'
	| 'table/invalid-cell-margin'
	| 'table/invalid-cell-border'
	| 'table/row-index-out-of-range'
	| 'table/column-index-out-of-range'
	| 'table/merge-range-invalid'
	// Hyperlinks
	| 'hyperlink/not-an-object'
	| 'hyperlink/missing-target'
	// HTML table conversion
	| 'html/no-document'
	| 'html/table-not-found'
	| 'html/table-has-no-cells'
	// Colours
	| 'color/invalid-hex'
	// Addressing an existing deck (`ts-pptx/read`)
	| 'slide/index-out-of-range'
	| 'slide/foreign-target'
	| 'shape/index-out-of-range'
	| 'layout/ambiguous-name'
	| 'layout/not-found'
	| 'layout/foreign-handle'
	// Moving slides, shapes, or masters between decks
	| 'import/slide-size-mismatch'
	| 'import/slide-size-unknown'
	| 'import/unresolved-slide-link'
	| 'import/destination-missing-master'
	| 'import/destination-missing-layout'
	// Replacing or adding picture content
	| 'image/undeterminable-extension'
	| 'image/undeterminable-type'
	| 'image/missing-content-type'
	| 'image/fit-needs-extent'
	// Text
	| 'font/size-not-positive'
	// Printing a deck back out as a script (`ts-pptx/script`). Both name an inconsistency in the
	// `DeckIr` handed to the printer — unreachable from an IR the library built itself, but the
	// IR is a public type a caller may construct or edit.
	| 'script/unresolved-asset-reference'
	| 'script/non-finite-literal'
	// OPC package operations asked of the library (distinct from a malformed package,
	// which is a `PackageReadError` — see `package/duplicate-relationship-id`)
	| 'package/duplicate-part-name'
	| 'part/not-xml'
	| 'relationship/duplicate-id'
	| 'relationship/not-found'
	| 'relationship/external-has-no-partname'
	// OOXML naming
	| 'oxml/unknown-namespace-prefix'
	| 'oxml/invalid-qname'
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
	// Requests with no OOXML expression the library emits
	| 'gradient/type-unsupported'
	| 'group/kind-not-groupable'
	| 'shape/element-unsupported'
	// Asked of a shape kind that has nowhere to put it
	| 'shape/fill-unsupported'
	| 'shape/line-unsupported'
	| 'shape/shape-properties-unsupported'
	| 'shape/no-text-frame'
	// Optional dependencies that were not installed or did not load
	| 'math/missing-optional-peer'
	| 'font/opentype-unavailable'
	// Runtime capabilities
	| 'zip/filesystem-unavailable'
	| 'runtime/file-output-unavailable'

/**
 * Conditions carried by `PackageReadError`: the bytes handed to the library are not a package it
 * can read, or a part inside one is structurally malformed. Always about *input*, never about
 * something the library is asked to produce.
 */
export type PackageReadErrorCode =
	// The bytes are not an archive, or the archive is not an OPC package
	| 'zip/not-a-zip-archive'
	| 'zip/file-read-failed'
	| 'package/not-an-opc-package'
	// `[Content_Types].xml`
	| 'package/content-types-invalid-root'
	| 'package/content-types-entry-incomplete'
	| 'package/part-content-type-missing'
	// `.rels` parts
	| 'package/relationships-invalid-root'
	| 'package/relationship-incomplete'
	| 'package/relationship-invalid-target-mode'
	| 'package/duplicate-relationship-id'
	| 'package/relationship-target-escapes-root'
	| 'package/relationship-target-missing'
	| 'package/office-document-relationship-invalid'
	// A part the package should contain is absent or structurally broken
	| 'package/part-missing'
	| 'package/part-has-no-root'
	| 'slide/no-shape-tree'
	| 'shape/no-non-visual-properties'
	| 'table/cell-has-no-text-body'

/**
 * Conditions carried by `MediaError`: an image, font, or audio/video payload could not be fetched,
 * read, or decoded. Kept separate from `PackageReadError` because the failure is in a *referenced
 * resource*, not in the package structure, and the recovery is usually different (fix the URL or
 * the file, not the deck).
 */
export type MediaErrorCode =
	// Shared with the diagnostic surface, and the reason the vocabulary is one registry: the same
	// load failure warns under `onMediaError: 'placeholder'` and throws under the default
	// fail-fast policy, so it must not be two different strings.
	| 'media/load-failed'
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
export type InternalErrorCode =
	| 'layout/default-not-registered'
	| 'slide/rel-index-out-of-range'
	| 'import/part-went-missing'
	| 'animation/timing-scaffold-failed'
	// A DOM node the read model was handed is detached. Reachable in principle through the
	// documented `part.dom` / `element_` escape hatch, but never from the library's own paths.
	| 'oxml/node-has-no-document'
	| 'oxml/node-has-no-parent'

/** Every condition the library can throw. See {@link TsPptxCode} for the whole vocabulary. */
export type ErrorCode =
	InvalidOptionErrorCode | UnsupportedFeatureErrorCode | PackageReadErrorCode | MediaErrorCode | InternalErrorCode

/** Every condition the library can report, warned or thrown. */
export type TsPptxCode = DiagnosticCode | ErrorCode
