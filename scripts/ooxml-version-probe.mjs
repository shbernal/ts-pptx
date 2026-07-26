#!/usr/bin/env node
// Validate decks against EVERY Office version the OOXML validator accepts and report
// the coverage profile. Run with: pnpm run schema:versions [--file <deck.pptx>]
//
// WHAT THIS IS NOT
//
// This is not a "which PowerPoint version can open the deck" check, and it cannot be
// turned into one. The intuition that the oldest version validating clean is the
// deck's floor is exactly backwards. The SDK's per-version schemas differ in how much
// markup they MODEL, not in what they accept: markup a version has never heard of is
// skipped, not rejected. So a chartEx (Office 2016) deck validates clean at Office2007
// precisely because Office2007's schema set cannot see it.
//
// The evidence is reproducible from this script:
//
//   fixture                             O2007 O2010 O2013 O2016 O2019 O2021  M365
//   base (plain text slide)                 0     0     0     0     0     0     0
//   chartEx pareto (2016 feature)           0     0     0     4     4     4     4
//   classic bar chart (2007 feature)        0     0     0     0     0     0     0
//   core-construct corruption               1     1     1     1     1     1     1
//
// Row 2 rises with version (newer schemas model chartEx and flag the known
// cx:axisId divergence); row 4 is flat (a bogus attribute on <p:sp> is a core
// ECMA-376 error every schema generation catches). Error count is monotonically
// non-decreasing in version, which is why `test/validator.js` pins FILE_FORMAT to
// Microsoft365: it is the strongest available check, and any lower version can only
// lose coverage.
//
// WHAT IT IS FOR
//
//  1. Re-verifying that monotonicity after bumping tools/ooxml-validator/version.json.
//     A DECREASE along the axis would break the premise the pin rests on, so the probe
//     exits non-zero on one.
//  2. Localizing a known divergence to the schema generation that introduced it. The
//     version where a fixture's count first moves is the generation that started
//     modelling that markup — that is how the pareto row above dates chartEx to 2016.
//
// It is deliberately NOT part of `verify`: 7 validator spawns per fixture, and it
// asserts nothing about emitted markup that `test:schema` does not already assert
// at Microsoft365.

import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import JSZip from 'jszip'
import TsPptx from '../dist/node.js'
import { isInstalled, runValidatorOnFile, FILE_FORMATS, VALIDATOR } from '../test/validator.js'

// Fixtures chosen to span the coverage axis, not to cover features — `test:schema`
// does that. Each one exists to make a different row shape observable: flat-clean,
// rising (version-gated markup), and flat-dirty (core error caught everywhere).
const FIXTURES = [
	{
		name: 'base (plain text slide)',
		build: (p) => {
			p.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1 })
		},
	},
	{
		name: 'classic bar chart (2007 feature)',
		build: (p) => {
			p.addSlide().addChart([{ name: 'D', labels: ['a', 'b', 'c'], values: [3, 2, 1] }], {
				type: 'bar',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
			})
		},
	},
	{
		name: 'chartEx pareto (2016 feature)',
		build: (p) => {
			p.addSlide().addChart([{ name: 'D', labels: ['a', 'b', 'c'], values: [3, 2, 1] }], {
				type: 'pareto',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
			})
		},
	},
	{
		// The control row. Without a fixture that is invalid at EVERY version, a table of
		// zeros is indistinguishable from a validator that silently stopped working.
		name: 'core-construct corruption (control)',
		build: (p) => {
			p.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1 })
		},
		corrupt: async (zip) => {
			const xml = await zip.file('ppt/slides/slide1.xml').async('string')
			zip.file('ppt/slides/slide1.xml', xml.replace('<p:sp>', '<p:sp bogusAttr="1">'))
		},
	},
]

async function buildFixture(fixture, dir) {
	const pres = new TsPptx()
	fixture.build(pres)
	let buf = /** @type {Uint8Array} */ (await pres.stream())
	if (fixture.corrupt) {
		const zip = await JSZip.loadAsync(buf)
		await fixture.corrupt(zip)
		buf = await zip.generateAsync({ type: 'nodebuffer' })
	}
	const file = path.join(dir, fixture.name.replace(/\W+/g, '_') + '.pptx')
	await fs.writeFile(file, buf)
	return file
}

function shortLabel(version) {
	return version.replace('Microsoft365', 'M365').replace('Office', 'O')
}

async function main() {
	if (!(await isInstalled())) {
		console.error('OOXMLValidatorCLI not installed at ' + VALIDATOR)
		console.error('Run: ./tools/ooxml-validator/install.sh')
		process.exit(1)
	}

	const fileArgIndex = process.argv.indexOf('--file')
	const explicitFile = fileArgIndex !== -1 ? process.argv[fileArgIndex + 1] : null
	if (fileArgIndex !== -1 && !explicitFile) {
		console.error('--file requires a path to a .pptx')
		process.exit(1)
	}

	const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'TsPptx-versions-'))
	try {
		const targets = explicitFile
			? [{ name: path.basename(explicitFile), file: path.resolve(explicitFile) }]
			: await Promise.all(FIXTURES.map(async (f) => ({ name: f.name, file: await buildFixture(f, dir) })))

		const nameWidth = Math.max(...targets.map((t) => t.name.length), 7)
		console.log('fixture'.padEnd(nameWidth), FILE_FORMATS.map((v) => shortLabel(v).padStart(6)).join(''))

		let regressions = 0
		for (const target of targets) {
			const counts = []
			for (const version of FILE_FORMATS) counts.push((await runValidatorOnFile(target.file, version)).length)
			console.log(target.name.padEnd(nameWidth), counts.map((c) => String(c).padStart(6)).join(''))

			// A decrease means a newer schema generation stopped modelling markup an older
			// one flagged. That inverts the premise behind pinning to Microsoft365, so it is
			// a hard failure rather than a note.
			for (let i = 1; i < counts.length; i++) {
				if (counts[i] < counts[i - 1]) {
					console.error(
						`\n  MONOTONICITY BROKEN on "${target.name}": ` +
							`${FILE_FORMATS[i - 1]}=${counts[i - 1]} -> ${FILE_FORMATS[i]}=${counts[i]}.\n` +
							'  test/validator.js pins FILE_FORMAT=Microsoft365 on the premise that error count\n' +
							'  never decreases with version. Re-read that comment before changing anything.'
					)
					regressions++
				}
			}
		}

		if (!explicitFile) {
			const control = targets.find((t) => t.name.includes('control'))
			console.log('\nThe control row must be non-zero at every version; an all-zero table means the')
			console.log('validator is not actually running, not that the fixtures are clean.')
			if (!control) console.error('WARNING: control fixture missing from this run.')
		}

		process.exit(regressions > 0 ? 1 : 0)
	} finally {
		await fs.rm(dir, { recursive: true, force: true })
	}
}

await main()
