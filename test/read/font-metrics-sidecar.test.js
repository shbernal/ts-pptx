// Keeps `fixtures/autofit-font-metrics.json` honest.
//
// That sidecar is what lets the autofit and CJK oracles run where their fonts cannot be
// installed: it carries the advance width of every code point the committed cases measure,
// per face, recorded from the genuine fonts by `fixtures/authoring/build-font-metrics.mjs`.
// Derived data standing in for a font is only worth as much as the check that it still
// matches the font, so this suite is that check, in three parts:
//
//   1. **Complete.** Every (face, code point) the committed cases reach has an entry. A
//      case added or edited without regenerating the sidecar fails here, with the face and
//      the character named, rather than at the point where a measurement silently uses a
//      default advance somewhere downstream.
//   2. **Faithful.** Wherever a genuine font resolves, every recorded advance is re-derived
//      from it and compared. This is the part that cannot run everywhere, and it is why the
//      CI matrix includes a Windows leg: that runner has Arial, Calibri, Tahoma and Malgun
//      Gothic for real, so four of the six faces are re-verified on every push. Aptos and
//      Aptos SemiBold ship with Microsoft 365, so their verification stays on a workstation
//      that has Office installed.
//   3. **Declared.** `FONT_ORACLES_GENUINE` lists the families a runner is expected to have
//      installed. It is how a runner image that quietly drops a font fails the leg instead
//      of falling through to the sidecar and reporting green, which is the exact failure
//      this whole arrangement exists to prevent.
import { describe, test, expect } from 'vitest'
import {
	deriveFace,
	diffFace,
	faceLabel,
	genuineMetrics,
	GENUINE_REQUIRED,
	neededFaces,
	readSidecar,
	resolveGenuineFontFile,
} from './font-oracle.js'

const needed = neededFaces()
const sidecar = readSidecar()

/** @param {{ family: string, bold?: boolean, italic?: boolean }} face */
function entryFor(face) {
	return sidecar.faces.find(
		(f) =>
			f.family.toLowerCase() === face.family.toLowerCase() && !!f.bold === !!face.bold && !!f.italic === !!face.italic
	)
}

describe('font metrics sidecar: complete', () => {
	test('the schema is the one this suite reads', () => {
		expect(sidecar.schema).toBe('font-metrics@1')
	})

	test('every face the committed cases measure has an entry', () => {
		const recorded = sidecar.faces.map((f) => faceLabel(f)).sort()
		expect(recorded).toEqual(needed.map((f) => faceLabel(f)).sort())
	})

	for (const face of needed) {
		test(`${faceLabel(face)}: every code point the cases use is recorded`, () => {
			const entry = entryFor(face)
			expect(entry, `no sidecar entry for ${faceLabel(face)}`).toBeTruthy()
			const missing = face.codepoints
				.filter((cp) => entry.advances[String(cp)] === undefined)
				.map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`)
			// A non-empty list here means the sidecar is stale, not that the model is wrong:
			// pnpm run font-metrics:build, on a machine with the fonts.
			expect(missing).toEqual([])
		})
	}
})

describe('font metrics sidecar: faithful to the fonts it was recorded from', () => {
	let verified = 0

	for (const face of needed) {
		const installed = resolveGenuineFontFile(face)
		test.skipIf(!installed)(`${faceLabel(face)}: recorded advances match the installed font`, async () => {
			const metrics = await genuineMetrics(face)
			expect(metrics, `${faceLabel(face)} resolved to ${installed} but would not parse`).toBeTruthy()
			const entry = entryFor(face)
			expect(entry, `no sidecar entry for ${faceLabel(face)}`).toBeTruthy()
			// Compare the whole entry, not a sample: a drifting advance is one number, and a
			// spot check is how it survives.
			expect(diffFace(entry, deriveFace(face, metrics))).toEqual([])
			verified++
		})
	}

	test('what this machine was able to verify (informational)', () => {
		const installed = needed.filter((f) => resolveGenuineFontFile(f)).map((f) => faceLabel(f))
		console.info(
			`font metrics sidecar: ${verified}/${needed.length} faces re-derived from installed fonts` +
				(installed.length > 0 ? ` (${installed.join(', ')})` : ' (none installed here)')
		)
		expect(verified).toBe(installed.length)
	})
})

// Only registered when FONT_ORACLES_GENUINE is set: a `describe` with no tests in it is a
// failure in vitest, and this list is empty on any machine that has not declared one.
describe.skipIf(GENUINE_REQUIRED.length === 0)('font metrics sidecar: declared installs', () => {
	for (const family of GENUINE_REQUIRED) {
		test(`${family} is installed on this runner, as declared`, () => {
			const file = resolveGenuineFontFile({ family })
			expect(
				file,
				`FONT_ORACLES_GENUINE names ${family}, but no installed file resolved for it. ` +
					`Either the runner image dropped the font, or the list is out of date.`
			).toBeTruthy()
		})
	}
})
