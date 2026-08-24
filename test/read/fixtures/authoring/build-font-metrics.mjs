// Writes `../autofit-font-metrics.json`: the advance widths the autofit and CJK oracles
// measure with, recorded from the genuine fonts so those oracles can run where the fonts
// cannot be installed.
//
// The oracles compare this repo's measured-fit model against what desktop PowerPoint baked,
// which is only a comparison if the model charges the same widths PowerPoint did. Five of
// the six faces involved ship with Windows or with Microsoft 365 and none of them can be
// committed, so on a hosted runner the oracles used to resolve nothing and pass anyway. The
// sidecar this writes is the way out: it carries the raw `hmtx` advance of every code point
// the committed cases actually use, per face, and nothing else. Two numbers per character,
// not a font.
//
// It is derived data, and the repo treats it that way. `../../font-metrics-sidecar.test.js`
// re-derives every entry from the installed font wherever one resolves and fails on any
// drift, so this file cannot be hand-edited into agreement with a model that has moved.
//
// **Run it on a machine with all six faces installed** (a Windows box with Microsoft 365:
// Aptos and Aptos SemiBold are Office fonts, installed per-user under `%LOCALAPPDATA%`).
// It refuses to write a partial sidecar, because a face silently dropped here becomes a
// suite that skips a case there.
//
//     pnpm run font-metrics:build
//     pnpm exec oxfmt --write "test/read/fixtures/*.json"   # see this directory's README
//
// The second command is not optional: every builder here emits raw `JSON.stringify(…, '\t')`
// and the sidecars are committed in the repo formatter's output.

import { writeFileSync } from 'node:fs'
import { relative } from 'node:path'
import {
	deriveFace,
	faceLabel,
	genuineMetrics,
	neededFaces,
	resolveGenuineFontFile,
	SIDECAR_PATH,
} from '../../font-oracle.js'

const faces = neededFaces()

const missing = faces.filter((face) => !resolveGenuineFontFile(face))
if (missing.length > 0) {
	console.error(`Cannot write the metrics sidecar: ${missing.length} of ${faces.length} faces did not resolve.\n`)
	for (const face of missing) console.error(`  - ${faceLabel(face)}`)
	console.error(
		`\nEvery face the committed cases use has to be installed, or the sidecar would be\n` +
			`written with a hole in it and the oracle would skip those cases without saying so.\n` +
			`Aptos and Aptos SemiBold come with Microsoft 365; the rest ship with Windows.`
	)
	process.exit(1)
}

const entries = []
for (const face of faces) {
	const metrics = await genuineMetrics(face)
	entries.push(deriveFace(face, /** @type {NonNullable<typeof metrics>} */ (metrics)))
}

const doc = {
	schema: 'font-metrics@1',
	note:
		'Advance widths in font design units for exactly the code points the committed autofit ' +
		'and CJK cases measure, read from the genuine installed faces by ' +
		'test/read/fixtures/authoring/build-font-metrics.mjs. Regenerate rather than edit; ' +
		'test/read/font-metrics-sidecar.test.js re-derives these from the installed fonts.',
	faces: entries,
}

writeFileSync(SIDECAR_PATH, `${JSON.stringify(doc, null, '\t')}\n`)

const codepoints = entries.reduce((n, e) => n + Object.keys(e.advances).length, 0)
console.log(`Wrote ${relative(process.cwd(), SIDECAR_PATH)}: ${entries.length} faces, ${codepoints} advances.`)
for (const entry of entries) {
	const uncovered = entry.uncovered.length > 0 ? `, ${entry.uncovered.length} uncovered` : ''
	console.log(`  ${faceLabel(entry)}: ${Object.keys(entry.advances).length} advances${uncovered}`)
}
