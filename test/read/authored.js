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

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { validateBuf, validatorInstalled } from '../validator.js'

// Re-exported rather than recomputed: `validator.js` owns the fact, and this module is where
// most read-side tests already import from.
export { validatorInstalled }

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
	const buf = await pres.toBytes()
	const presentation = await Presentation.load(buf)
	return { presentation, buf, pres }
}

/**
 * `authorRead`, but with the PowerPoint-authored `ppt/tableStyles.xml` from
 * `fixtures/table-styles.pptx` spliced in before the read.
 *
 * The read side's table style graph — `Table.resolvedStyle`, and the header/banding shading
 * `TableCell.resolvedFill` inherits — needs a deck whose styles part actually defines the
 * style the table names. The write API cannot produce one: it emits the part as a bare
 * default-id stub, because PowerPoint resolves `<a:tableStyleId>` against its own gallery and
 * never reads a definition out of the package (`defineTableStyle()` was removed for exactly
 * that reason). Real decks get theirs from PowerPoint, so the fixture's part is both the only
 * available oracle and the more honest one.
 *
 * Author the table with `tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1` — the fixture defines
 * that GUID, with `firstRow` shading and `band1H`/`band1V` banding.
 *
 * @param {(pres: InstanceType<typeof TsPptx>) => void | Promise<void>} build
 * @returns {Promise<{ presentation: Presentation, buf: Uint8Array, pres: InstanceType<typeof TsPptx> }>}
 */
export async function authorReadWithFixtureStyles(build) {
	const pres = new TsPptx()
	await build(pres)
	const authored = /** @type {Uint8Array} */ (await pres.toBytes())

	const fixture = await JSZip.loadAsync(
		await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'table-styles.pptx'))
	)
	const zip = await JSZip.loadAsync(authored)
	zip.file('ppt/tableStyles.xml', await fixture.file('ppt/tableStyles.xml').async('string'))
	const buf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })

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
 * Schema-validate authored bytes. Returns the oracle's diagnostics (empty ⇒ valid).
 * Gate the calling test with `test.skipIf(!validatorInstalled)` so the suite stays
 * green where the oracle cannot be obtained.
 *
 * @param {Uint8Array} buf
 * @returns {Promise<readonly import('ooxml-validate').ValidationDiagnostic[]>}
 */
export async function schemaErrors(buf) {
	return validateBuf(Buffer.from(buf))
}
