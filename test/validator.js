// Test-time helper: validate a `.pptx` Buffer against Microsoft's OpenXmlValidator.
// Returns a Promise<Array> of diagnostics (empty array on a clean file).
//
// This is a thin adapter over `ooxml-validate`, the shared validation oracle this
// project and `ts-xlsx` both use. Everything this file used to own — the binary
// install, the batch queue, the conformance pin, the CI gate — belongs to that package
// now. The two repos had each grown their own validator pinned to a different Open XML
// SDK version, which meant they were enforcing different rule sets while appearing to
// enforce the same one; a single shared oracle is what removes that.
//
// The diagnostic shape comes from the package (`ValidationDiagnostic`):
//
//   { id, type, description, partUri, xpath }
//
// `id` is a stable machine code such as `Sch_UndeclaredAttribute`; `description` is
// upstream prose and can be reworded in any release, so assert on `id`. `partUri` and
// `xpath` are null on a package-level failure, where there is no part to point at.

import { FILE_FORMAT, validateBuffer, validatorAvailable, validatorPath } from 'ooxml-validate'

// The conformance target, pinned by the package at `Microsoft365` and re-exported so
// the suites keep reading it from one place.
//
// Worth knowing why that value, because the intuition runs the other way:
// `Microsoft365` is the STRONGEST available check, not merely the newest. The SDK's
// per-version schemas differ only in how much markup they MODEL — an older version
// does not reject newer constructs, it skips them — so error count is monotonically
// non-decreasing in version and validating anywhere lower can only lose coverage.
// `pnpm run schema:versions` prints the evidence; the package's own `probeFormats` is
// the executable form of the claim, re-checked against every SDK bump rather than
// measured once.

/**
 * Whether the oracle can be obtained.
 *
 * Unlike the old bare existence check this may *fetch* the binary on its first call —
 * resolution is the package's business now, and "installed" stopped being a question
 * about one path on disk. Prefer `validatorAvailable` at a call site that gates
 * assertions; this is for the two places that need the plain answer.
 */
async function isInstalled() {
	return (await validatorPath()) !== null
}

/**
 * Validate one in-memory deck.
 *
 * Batching, the one-child-at-a-time queue and the temp-file bookkeeping all live in
 * the package. Correlation between a buffer and its result is by the handle the
 * package tracks, never by position in a batch.
 *
 * @param {Uint8Array} buf
 * @param {import('ooxml-validate').FileFormat} [fileFormat]
 */
async function validateBuf(buf, fileFormat = FILE_FORMAT) {
	const result = await validateBuffer(buf, { ext: '.pptx', format: fileFormat })
	return result.errors
}

export { isInstalled, validatorAvailable, validateBuf, FILE_FORMAT }
