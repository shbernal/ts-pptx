#!/usr/bin/env node
/**
 * What the measurement oracles will measure with on this machine.
 *
 * The autofit and CJK oracles (`test/read/autofit-calibration-oracle.test.js`,
 * `test/read/cjk-line-breaking-oracle.test.js`) need the genuine faces PowerPoint used
 * when it baked their fixtures, and fall back to the committed metrics sidecar where a
 * machine does not have one. Both are legitimate, but they are not the same claim, and
 * the difference used to be invisible: the suites resolved whatever was there and
 * reported green either way.
 *
 * This prints which source each face resolves to, and turns a runner's expectations into
 * a gate:
 *
 *   node scripts/font-oracle-probe.mjs     # report; exit 1 if a declared family is absent
 *
 * `FONT_ORACLES_GENUINE` is the declaration: a comma-separated list of families this
 * machine is expected to have installed for real. Set it on a CI leg chosen for its font
 * set, and the leg fails when the image stops carrying one instead of silently falling
 * back to recorded advances. Unset, this is a report and always exits 0.
 *
 * On GitHub Actions the table is appended to the job summary as well as printed, so the
 * font situation of a run is readable without opening the log.
 */

import { appendFileSync } from 'node:fs'
import { relative } from 'node:path'
import { ROOT } from './script-utils.mjs'
import {
	faceLabel,
	GENUINE_REQUIRED,
	neededFaces,
	readSidecar,
	resolveGenuineFontFile,
	SIDECAR_PATH,
} from '../test/read/font-oracle.js'

const sidecar = readSidecar()
const faces = neededFaces()

/** @param {{ family: string, bold?: boolean, italic?: boolean }} face */
const hasSidecarEntry = (face) =>
	sidecar.faces.some(
		(f) =>
			f.family.toLowerCase() === face.family.toLowerCase() && !!f.bold === !!face.bold && !!f.italic === !!face.italic
	)

const rows = faces.map((face) => {
	const file = resolveGenuineFontFile(face)
	return {
		face: faceLabel(face),
		source: file ? 'installed font' : hasSidecarEntry(face) ? 'metrics sidecar' : 'NOTHING',
		detail: file ?? (hasSidecarEntry(face) ? 'recorded advances' : 'no font, no sidecar entry'),
		codepoints: face.codepoints.length,
	}
})

const table = [
	'| Face | Source | Code points | Resolved from |',
	'| --- | --- | --- | --- |',
	...rows.map((r) => `| ${r.face} | ${r.source} | ${r.codepoints} | \`${r.detail}\` |`),
].join('\n')

const installed = rows.filter((r) => r.source === 'installed font').length
const heading = `Font oracles on ${process.platform}: ${installed}/${rows.length} faces from installed fonts`

console.log(`${heading}\n\n${table}\n`)

if (process.env.GITHUB_STEP_SUMMARY) {
	appendFileSync(process.env.GITHUB_STEP_SUMMARY, `### ${heading}\n\n${table}\n\n`)
}

// The declaration is the gate. A family named here that does not resolve means the image
// changed under us, and every case using it would have quietly measured from the sidecar.
const undeclared = GENUINE_REQUIRED.filter((family) => !resolveGenuineFontFile({ family }))
if (undeclared.length > 0) {
	console.error(
		`\nFONT_ORACLES_GENUINE declares ${GENUINE_REQUIRED.length} installed families, but these did not resolve:\n` +
			undeclared.map((f) => `  - ${f}`).join('\n') +
			`\n\nEither the runner image dropped the font, or the list is out of date. Do not "fix"\n` +
			`this by trimming the list without saying so: the sidecar would still answer, and the\n` +
			`leg would go green having verified nothing against a real font.`
	)
	process.exit(1)
}

const nothing = rows.filter((r) => r.source === 'NOTHING')
if (nothing.length > 0) {
	console.error(
		`\n${nothing.length} face(s) resolve to neither a font nor a sidecar entry. Regenerate the\n` +
			`sidecar on a machine that has them: pnpm run font-metrics:build\n` +
			`(${relative(ROOT, SIDECAR_PATH)})`
	)
	process.exit(1)
}
