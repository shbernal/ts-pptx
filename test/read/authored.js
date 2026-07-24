// Shared write→read fidelity harness for the read-side-expansion batches.
//
// The round-trip matrix's thesis is that every read-side feature ships a
// *measured* fidelity number: author the feature with the write API (which
// already emits it), load the bytes back through the deep read model
// (src/read/api/*), and assert the extracted model. This module is that
// author→read step — factored out of the one-off IIFE that
// style-accessors.test.js hand-rolls, plus the slide-walking locators that
// chart.test.js / table.test.js each redefine.
//
// It deliberately does NOT wrap the per-feature assertions: each batch asserts
// its own getters against its own oracle. This is only the fixture-in-memory +
// locate plumbing. Loading the writer's bytes with the deep read model keeps the
// two sides independent — the write path (fflate serializer) and the read path
// (src/read/*) are separate code, so a bug in one can't mask a bug in the other.
//
// Not a test file (no `.test.` in the name) — vitest's default glob skips it.

import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { isInstalled, validateBuf } from '../validator.js'

/** True when the OOXML schema validator is installed; gate schema legs with this. */
export const validatorInstalled = await isInstalled()

/**
 * Author a deck in memory with the write API and load it into the deep read
 * model. `build` receives a fresh TsPptx instance — add slides / shapes /
 * charts / tables with the normal write API — and may be async.
 *
 * @param {(pres: InstanceType<typeof TsPptx>) => void | Promise<void>} build
 * @returns {Promise<{ presentation: Presentation, buf: Uint8Array, pres: InstanceType<typeof TsPptx> }>}
 */
export async function authorRead(build) {
	const pres = new TsPptx()
	await build(pres)
	// Under Node, stream() resolves to a Uint8Array (see test/helpers.js build()).
	const buf = /** @type {Uint8Array} */ (await pres.stream())
	const presentation = await Presentation.load(buf)
	return { presentation, buf, pres }
}

/** Every shape on every slide, in document (slide, then shape) order. */
export function allShapes(presentation) {
	return presentation.slides.flatMap((slide) => slide.shapes)
}

/** The first shape on any slide matching `predicate`, or null. */
export function firstShape(presentation, predicate) {
	return allShapes(presentation).find(predicate) ?? null
}

/** The first chart on any slide, or null. */
export function firstChart(presentation) {
	const frame = firstShape(presentation, (s) => s.shapeType === 'graphicFrame' && s.chart)
	return frame ? frame.chart : null
}

/** The first chartEx (waterfall/funnel/treemap/…) chart on any slide, or null. */
export function firstChartEx(presentation) {
	const frame = firstShape(presentation, (s) => s.shapeType === 'graphicFrame' && s.chartEx)
	return frame ? frame.chartEx : null
}

/** The first table on any slide, or null. */
export function firstTable(presentation) {
	const frame = firstShape(presentation, (s) => s.shapeType === 'graphicFrame' && s.table)
	return frame ? frame.table : null
}

/**
 * Schema-validate authored bytes. Returns the validator's error array (empty ⇒
 * valid). Gate the calling test with `test.skipIf(!validatorInstalled)` so the
 * suite stays green where the OOXML validator isn't installed.
 *
 * @param {Uint8Array} buf
 * @returns {Promise<unknown[]>}
 */
export async function schemaErrors(buf) {
	return validateBuf(Buffer.from(buf))
}
