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
// non-decreasing in version, which is why `ooxml-validate` pins FILE_FORMAT to
// Microsoft365: it is the strongest available check, and any lower version can only
// lose coverage.
//
// WHAT IT IS FOR
//
//  1. Re-verifying that monotonicity after an `ooxml-validate` bump — which is also an
//     Open XML SDK bump, since the package pins one. A DECREASE along the axis would
//     break the premise the pin rests on, so the probe exits non-zero on one.
//  2. Localizing a known divergence to the schema generation that introduced it. The
//     version where a fixture's count first moves is the generation that started
//     modelling that markup — that is how the pareto row above dates chartEx to 2016.
//
// The counting is `ooxml-validate`'s own `probeFormats`, so this script is the
// project-specific half — the fixtures that make each row shape observable — over a
// claim the package re-checks for every consumer. What is measured here is TsPptx's
// output; what is asserted is the package's pin.
//
// It is deliberately NOT part of `verify`: seven validation passes per fixture, and it
// asserts nothing about emitted markup that `test:schema` does not already assert
// at Microsoft365.

import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import JSZip from 'jszip'
import { probeFormats, validatorPath } from 'ooxml-validate'
import TsPptx from '../dist/node.js'
import { parseCliOrExit } from './script-utils.mjs'

// Fixtures chosen to span the coverage axis, not to cover features — `test:schema`
// does that. Each one exists to make a different row shape observable: flat-clean,
// rising (version-gated markup), and flat-dirty (core error caught everywhere).
/**
 * @typedef {object} Fixture
 * @property {string} name
 * @property {(pres: import('../dist/node.js').TsPptx) => void} build
 * @property {(zip: import('jszip')) => Promise<void>} [corrupt] damage the built package on purpose
 */
/** @type {Fixture[]} */
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
			const slide = zip.file('ppt/slides/slide1.xml')
			if (!slide) throw new Error('built package has no ppt/slides/slide1.xml to corrupt')
			const xml = await slide.async('string')
			zip.file('ppt/slides/slide1.xml', xml.replace('<p:sp>', '<p:sp bogusAttr="1">'))
		},
	},
]

/**
 * @param {Fixture} fixture
 * @param {string} dir
 * @returns {Promise<string>} path to the built deck
 */
async function buildFixture(fixture, dir) {
	const pres = new TsPptx()
	fixture.build(pres)
	let buf = await pres.toBytes()
	if (fixture.corrupt) {
		const zip = await JSZip.loadAsync(buf)
		await fixture.corrupt(zip)
		buf = await zip.generateAsync({ type: 'nodebuffer' })
	}
	const file = path.join(dir, fixture.name.replace(/\W+/g, '_') + '.pptx')
	await fs.writeFile(file, buf)
	return file
}

/**
 * @param {string} version
 * @returns {string}
 */
function shortLabel(version) {
	return version.replace('Microsoft365', 'M365').replace('Office', 'O')
}

async function main() {
	// Arguments first, validator second: `--help` has to work on a machine that has
	// never fetched the oracle, which is exactly where someone reads it.
	//
	// `parseArgs` rejects a `--file` with no value itself, so the hand-rolled
	// "requires a path" check that used to live here is gone with the indexOf form.
	const { values } = parseCliOrExit(process.argv.slice(2), {
		usage: `Validate decks against every Office version the OOXML validator accepts.

  pnpm run schema:versions
  pnpm run schema:versions -- --file path/to/deck.pptx

Options:
  --file <path>  probe one existing deck instead of the built-in fixtures
  -h, --help     show this message`,
		options: { file: { type: 'string' } },
	})
	const explicitFile = values.file ?? null

	if ((await validatorPath()) === null) {
		console.error('the ooxml-validate oracle could not be obtained.')
		console.error('It is fetched from GitHub Releases on first use; see docs/testing.md.')
		process.exit(1)
	}

	const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'TsPptx-versions-'))
	try {
		const targets = explicitFile
			? [{ name: path.basename(explicitFile), file: path.resolve(explicitFile) }]
			: await Promise.all(FIXTURES.map(async (f) => ({ name: f.name, file: await buildFixture(f, dir) })))

		// One call per conformance target for the whole set, not one per fixture: the
		// package batches, so probing every fixture together costs seven oracle runs in
		// total rather than seven per row.
		const probe = await probeFormats(targets.map((t) => t.file))
		const rowsByFile = new Map(probe.rows.map((row) => [row.file, row]))

		const nameWidth = Math.max(...targets.map((t) => t.name.length), 7)
		console.log('fixture'.padEnd(nameWidth), probe.formats.map((v) => shortLabel(v).padStart(6)).join(''))

		for (const target of targets) {
			const row = rowsByFile.get(target.file)
			if (!row) {
				console.error(`\n  NO RESULT for "${target.name}" (${target.file}).`)
				continue
			}
			console.log(target.name.padEnd(nameWidth), row.counts.map((c) => String(c).padStart(6)).join(''))

			// A decrease means a newer schema generation stopped modelling markup an older
			// one flagged. That inverts the premise behind pinning to Microsoft365, so it is
			// a hard failure rather than a note.
			for (let i = 1; i < row.counts.length; i++) {
				const previous = row.counts[i - 1] ?? 0
				const current = row.counts[i] ?? 0
				if (current < previous) {
					console.error(
						`\n  MONOTONICITY BROKEN on "${target.name}": ` +
							`${probe.formats[i - 1]}=${previous} -> ${probe.formats[i]}=${current}.\n` +
							'  ooxml-validate pins FILE_FORMAT=Microsoft365 on the premise that error count\n' +
							'  never decreases with version. Re-read that comment before changing anything.'
					)
				}
			}
		}

		if (!explicitFile) {
			const control = targets.find((t) => t.name.includes('control'))
			console.log('\nThe control row must be non-zero at every version; an all-zero table means the')
			console.log('validator is not actually running, not that the fixtures are clean.')
			if (!control) console.error('WARNING: control fixture missing from this run.')
		}

		process.exit(probe.violated ? 1 : 0)
	} finally {
		await fs.rm(dir, { recursive: true, force: true })
	}
}

await main()
