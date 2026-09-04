#!/usr/bin/env node
/**
 * PowerPoint desktop COM smoke test (Windows only).
 *
 * CI's XML/schema checks can't catch three classes of desktop regression:
 *   1. Files that PowerPoint reports as corrupt and offers to "repair" (0x80070570),
 *      which the schema validator happily accepts.
 *   2. hlinkClick navigation actions that are schema-valid but that PowerPoint fails
 *      to parse back into a live slide-show action.
 *   3. custGeom `<a:cxnLst>` connection sites that are schema-valid but that PowerPoint
 *      never binds a connector to — i.e. `<a:stCxn>` points at a site PowerPoint can't
 *      resolve, so the connector opens as unconnected floating geometry.
 *   4. embedded OLE objects (`<p:oleObj>`) that are schema-valid but that PowerPoint
 *      silently drops on open, leaving the slide with no shape where the object was.
 *   5. embedded 3D models (`<am3d:model3d>`) that PowerPoint accepts but does not draw —
 *      the model collapses to its fallback picture, or renders as an empty frame. Nothing
 *      upstream can catch this: the SDK validator does not descend into an `mc:Choice` at
 *      all, so every mutation inside the `am3d` subtree validates clean.
 *   6. preset adjustment guides (`<a:gd>`) driven past the range the preset's own handle
 *      allows. These are inert rather than corrupting, and this is the only thing that can
 *      say so: `Shape.Adjustments` reports the value as *stored*, so only the pixels answer
 *      what PowerPoint paints.
 *
 * This script drives the real PowerPoint application over COM (via cscript) to check
 * all six. By default it runs five generated decks from the built `dist/`:
 *   - a navigation deck, reading each button's `ActionSettings(ppMouseClick).Action`
 *     back out and asserting each jump resolved to the correct PpActionType enum; and
 *   - a custGeom deck with connection sites + a connector bound to site index 1, reading
 *     the connector's `ConnectorFormat` back out and asserting it begin-connected to the
 *     custom shape at a real connection site; and
 *   - an OLE deck with an embedded OPC package and a generic OLE blob, reading each
 *     shape's `OLEFormat.ProgID` back out to prove the embedded object survived the open; and
 *   - a 3D-model deck with one embedded `.glb` over a deliberately magenta preview, reading
 *     `Shape.Type`/`Model3DFormat.CameraPositionZ` back out AND exporting the slide to PNG,
 *     because the read-back only proves PowerPoint resolved a model — the pixels are what
 *     prove it drew one (no magenta ⇒ not the fallback; not blank ⇒ not an empty frame); and
 *   - a preset-geometry deck whose slides pair an out-of-range adjustment guide with the same
 *     shape at the bound, exported to PNG and compared pixel for pixel. Each pair carries a
 *     third, in-range slide that must paint *differently*, so a run that rendered nothing
 *     fails instead of satisfying every equality.
 * Point it at any deck with `--file` to run only the corruption-open check.
 *
 *   node scripts/powerpoint-com-smoke.mjs                 # nav + custGeom + OLE + prstGeom + 3D-model checks
 *   node scripts/powerpoint-com-smoke.mjs --keep          # ...and keep the generated .pptx files
 *   node scripts/powerpoint-com-smoke.mjs --file deck.pptx # corruption-open check on an existing deck
 *
 * Requirements: Windows with PowerPoint installed. No-ops with a clear message elsewhere.
 */
import os from 'node:os'
import fs from 'node:fs/promises'
import path from 'node:path'
import { spawn, spawnSync } from 'node:child_process'
import { decodePng } from './png-utils.mjs'
import { parseCliOrExit, skipOrFail } from './script-utils.mjs'
import {
	EXPECTED_ACTION,
	EXPECTED_OLE_PROGID,
	GEOM_EXPECTED_SITE,
	GEOM_START_IDX,
	GEOM_TARGET_NAME,
	MODEL3D_EXPECTED_CAMERA_Z,
	MODEL3D_FRAME_IN,
	MODEL3D_LAYOUT_IN,
	MODEL3D_SHAPE_NAME,
	MODEL3D_SHAPE_TYPE,
	PRSTGEOM_CASES,
} from './com/contract.mjs'
import {
	generateGeomDeck,
	generateModel3dDeck,
	generateNavDeck,
	generateOleDeck,
	generatePresetGeomDeck,
} from './com/decks.mjs'
import {
	buildGeomVbs,
	buildModel3dVbs,
	buildNavVbs,
	buildOleVbs,
	buildPresetGeomVbs,
	vbsFooter,
	vbsOpenHeader,
} from './com/vbs.mjs'

// --- args -------------------------------------------------------------------
const USAGE = `PowerPoint COM smoke — open generated decks in desktop PowerPoint (Windows only).

  pnpm run test:com
  pnpm run test:com -- --keep
  pnpm run test:com -- --file path/to/deck.pptx

Options:
  --keep          leave the generated decks on disk for inspection
  --file <path>   open an existing deck instead of generating the corpus
  -h, --help      show this message

Environment:
  TSPPTX_COM_SMOKE   set to "required" to fail, not SKIP, when PowerPoint is unavailable`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: { keep: { type: 'boolean', default: false }, file: { type: 'string' } },
})
const KEEP = values.keep
const EXISTING_FILE = values.file ?? null

if (os.platform() !== 'win32') {
	process.exit(
		skipOrFail('TSPPTX_COM_SMOKE', 'the PowerPoint COM smoke is Windows-only (platform: ' + os.platform() + ').')
	)
}
// --- 2. clear the PowerPoint Resiliency key ---------------------------------
// A prior crash can leave the file in the Disabled/Resiliency list, so PowerPoint refuses
// to open it (or opens in reduced-functionality mode) and the smoke gives a false failure.
function clearResiliency() {
	for (const ver of ['16.0', '15.0', '14.0']) {
		spawnSync('reg', ['delete', `HKCU\\Software\\Microsoft\\Office\\${ver}\\PowerPoint\\Resiliency`, '/f'], {
			stdio: 'ignore',
		})
	}
}

/**
 * @param {string} vbsFile
 * @returns {Promise<{code: number, out: string, err: string}>}
 */
function runCscript(vbsFile) {
	return new Promise((resolve) => {
		const child = spawn('cscript', ['//nologo', '//B', vbsFile], { stdio: ['ignore', 'pipe', 'pipe'] })
		let out = ''
		let err = ''
		child.stdout.on('data', (d) => (out += d))
		child.stderr.on('data', (d) => (err += d))
		// A `null` code means the process was killed by a signal; report it the way the
		// spawn-error path does, so the retry in `driveDeck` treats the two alike.
		child.on('close', (code) => resolve({ code: code ?? -1, out, err }))
		child.on('error', (e) => resolve({ code: -1, out, err: String(e) }))
	})
}

/**
 * Write a VBS for `file`, drive PowerPoint (retry once), and return the raw cscript result.
 * @param {string} label
 * @param {string} file the deck to open
 * @param {(file: string) => string} buildVbs
 * @returns {Promise<{code: number, out: string, err: string}>}
 */
async function driveDeck(label, file, buildVbs) {
	const vbsFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-${label}-${process.pid}.vbs`)
	await fs.writeFile(vbsFile, buildVbs(file))
	// cscript can transiently fail if PowerPoint is mid-launch; retry once.
	/** @type {{code: number, out: string, err: string}} */
	let result = { code: -1, out: '', err: 'cscript was never run' }
	for (let attempt = 1; attempt <= 2; attempt++) {
		clearResiliency()
		result = await runCscript(vbsFile)
		if (result.code !== -1 && !/OPEN_ERR/.test(result.out)) break
		if (attempt === 1) console.log(`[${label}] first attempt failed; retrying once...`)
	}
	await fs.rm(vbsFile, { force: true })
	return result
}

// --- 4. verifiers -----------------------------------------------------------
/**
 * Returns { skip, failures }. Handles the NO_POWERPOINT / OPEN_ERR / OPEN_OK common cases.
 * @param {string} label
 * @param {string[]} lines cscript's stdout, split into lines
 * @param {string} out cscript's raw stdout
 * @returns {{skip: boolean, failures: string[]}}
 */
function checkOpen(label, lines, out) {
	if (/NO_POWERPOINT/.test(out)) return { skip: true, failures: [] }
	if (!lines.some((l) => l.startsWith('OPEN_OK'))) {
		const openErr = lines.find((l) => l.startsWith('OPEN_ERR')) || '(no OPEN_OK line)'
		return {
			skip: false,
			failures: [`[${label}] PowerPoint failed to open the deck (possible corruption / 0x80070570): ${openErr}`],
		}
	}
	const openOk = lines.find((l) => l.startsWith('OPEN_OK')) ?? ''
	console.log(`[${label}] ` + openOk.replace('OPEN_OK\t', 'opened OK; slide count = '))
	return { skip: false, failures: [] }
}

/**
 * @param {string[]} lines
 * @returns {string[]} one message per failure
 */
function verifyNav(lines) {
	/** @type {string[]} */
	const failures = []
	/** @type {Map<string | undefined, number>} */
	const seen = new Map()
	for (const l of lines) {
		if (!l.startsWith('ACTION\t')) continue
		const [, , name, num] = l.split('\t')
		seen.set(name, Number(num))
	}
	for (const [jump, wantNum] of Object.entries(EXPECTED_ACTION)) {
		const got = seen.get(jump)
		if (got === undefined) failures.push(`button "${jump}": no navigation action read back (dead geometry?)`)
		else if (got !== wantNum) failures.push(`button "${jump}": resolved to PpActionType ${got}, expected ${wantNum}`)
		else console.log(`  OK  ${jump} -> PpActionType ${got}`)
	}
	return failures
}

/**
 * @param {string[]} lines
 * @returns {string[]}
 */
function verifyGeom(lines) {
	/** @type {string[]} */
	const failures = []
	const conns = lines.filter((l) => l.startsWith('CONN\t')).map((l) => l.split('\t'))
	if (!conns.length) {
		failures.push('no connector shape found on the custGeom deck (connector dropped?)')
		return failures
	}
	// The deck has exactly one bound connector.
	const [, name, connected, targetName, site] = conns[0] ?? []
	if (String(connected) !== '-1') {
		failures.push(
			`connector "${name}" is not begin-connected (BeginConnected=${connected}); the connector did not snap to the custom shape's site`
		)
	} else if (targetName !== GEOM_TARGET_NAME) {
		failures.push(`connector "${name}" snapped to "${targetName}", expected the custGeom shape "${GEOM_TARGET_NAME}"`)
	} else if (Number(site) !== GEOM_EXPECTED_SITE) {
		failures.push(
			`connector "${name}" begin-connected to "${targetName}" but at site ${site}, expected ${GEOM_EXPECTED_SITE} (OOXML idx ${GEOM_START_IDX})`
		)
	} else {
		console.log(`  OK  connector "${name}" begin-connected to "${targetName}" at connection site ${site}`)
	}
	return failures
}

/**
 * @param {string[]} lines
 * @returns {string[]}
 */
function verifyOle(lines) {
	/** @type {string[]} */
	const failures = []
	/** @type {Map<string | undefined, string | undefined>} */
	const seen = new Map()
	for (const l of lines) {
		if (!l.startsWith('OLE\t')) continue
		const [, name, progId] = l.split('\t')
		seen.set(name, progId)
	}
	for (const [name, wantProgId] of Object.entries(EXPECTED_OLE_PROGID)) {
		const got = seen.get(name)
		if (got === undefined) {
			failures.push(`OLE object "${name}": no embedded object read back (PowerPoint discarded the graphicFrame?)`)
		} else if (got !== wantProgId) {
			failures.push(`OLE object "${name}": resolved to progId "${got}", expected "${wantProgId}"`)
		} else {
			console.log(`  OK  ${name} -> progId ${got}`)
		}
	}
	return failures
}

/**
 * Verifier for the 3D-model deck. Async because it reads back the exported PNG — which is the
 * only check here that distinguishes "PowerPoint resolved a model" from "PowerPoint drew one".
 * @param {string[]} lines
 * @returns {Promise<string[]>}
 */
async function verifyModel3d(lines) {
	/** @type {string[]} */
	const failures = []
	const row = lines
		.map(/** @param {string} l */ (l) => l.split('\t'))
		.find(/** @param {string[]} p */ (p) => p[0] === 'M3D' && p[1] === MODEL3D_SHAPE_NAME)
	if (!row) {
		failures.push(
			`3D model "${MODEL3D_SHAPE_NAME}": no such shape (PowerPoint discarded the graphicFrame, or resolved only the fallback picture)`
		)
		return failures
	}
	const [, name, type, camZ] = row
	if (Number(type) !== MODEL3D_SHAPE_TYPE) {
		failures.push(
			`3D model "${name}": Shape.Type ${type}, expected ${MODEL3D_SHAPE_TYPE} (msoShape3DModel) — PowerPoint did not bind it as a live model`
		)
	} else if (!camZ || Math.abs(Number(camZ) - MODEL3D_EXPECTED_CAMERA_Z) > 0.001) {
		failures.push(
			`3D model "${name}": CameraPositionZ read back as "${camZ}", expected ~${MODEL3D_EXPECTED_CAMERA_Z} — the am3d:camera subtree did not parse`
		)
	} else {
		console.log(`  OK  ${name} -> Shape.Type ${type} (msoShape3DModel), CameraPositionZ ${camZ}`)
	}

	// The render check.
	const exportErr = lines.find((l) => l.startsWith('EXPORT_ERR'))
	const exportLine = lines.find((l) => l.startsWith('EXPORT\t'))
	if (exportErr || !exportLine) {
		failures.push(`3D model: slide export failed (${exportErr || 'no EXPORT line'})`)
		return failures
	}
	const pngPath = exportLine.split('\t')[1]
	if (!pngPath) {
		failures.push('3D model: EXPORT line carries no path')
		return failures
	}
	let img
	try {
		img = decodePng(await fs.readFile(pngPath))
	} catch (e) {
		failures.push(`3D model: could not read the exported PNG (${String(e)})`)
		return failures
	} finally {
		if (!KEEP) await fs.rm(pngPath, { force: true })
	}

	// The frame, in exported pixels, from the layout the deck was built at.
	const sx = img.w / MODEL3D_LAYOUT_IN.w
	const sy = img.h / MODEL3D_LAYOUT_IN.h
	const x0 = Math.round(MODEL3D_FRAME_IN.x * sx)
	const y0 = Math.round(MODEL3D_FRAME_IN.y * sy)
	const x1 = Math.min(img.w - 1, Math.round((MODEL3D_FRAME_IN.x + MODEL3D_FRAME_IN.w) * sx))
	const y1 = Math.min(img.h - 1, Math.round((MODEL3D_FRAME_IN.y + MODEL3D_FRAME_IN.h) * sy))
	let fallbackPx = 0
	let drawnPx = 0
	for (let y = y0; y <= y1; y++) {
		for (let x = x0; x <= x1; x++) {
			const [r, g, b] = img.rgb(x, y)
			// A tolerance band around `MODEL3D_PREVIEW_RGB`, not the exact value: PowerPoint's PNG
			// export resamples the preview, so an exact match would miss an edge-antialiased one.
			if (r > 200 && g < 60 && b > 200) fallbackPx++
			if (r < 250 || g < 250 || b < 250) drawnPx++
		}
	}
	const area = (x1 - x0 + 1) * (y1 - y0 + 1)
	if (fallbackPx > 0) {
		failures.push(
			`3D model: ${fallbackPx}/${area} exported pixels are the magenta preview — PowerPoint drew the mc:Fallback picture, not the model`
		)
	} else if (drawnPx < area / 20) {
		failures.push(
			`3D model: only ${drawnPx}/${area} pixels in the frame were drawn — the model rendered as an empty frame (unreadable .glb payload?)`
		)
	} else {
		console.log(`  OK  model rasterized: ${drawnPx}/${area} pixels drawn in-frame, 0 fallback pixels`)
	}
	return failures
}

/**
 * Verifier for the preset-geometry deck.
 *
 * The claim: a finite adjustment guide outside the preset's own handle range is *inert*. PowerPoint
 * stores it verbatim -- it neither repairs the package nor rewrites the value on re-save -- and the
 * preset's guide formula pins it, so the shape it paints is the shape it paints at the bound. That
 * is why `genXmlPresetGeom` emits a finite adjustment as written instead of clamping it against a
 * per-preset range table, which lives in the preset shape definitions rather than in ECMA-376 and
 * would have to be invented here.
 *
 * Each `same` pair asserts two exported slides are pixel-identical; each `differs` pair asserts two
 * are not. The second is the sensitivity check: without it, six slides that failed to render
 * anything would satisfy every equality in the set.
 * @param {string[]} lines
 * @returns {Promise<string[]>}
 */
async function verifyPresetGeom(lines) {
	/** @type {string[]} */
	const failures = []
	const exportErr = lines.find((l) => l.startsWith('EXPORT_ERR'))
	if (exportErr) {
		failures.push(`preset geometry: slide export failed (${exportErr})`)
		return failures
	}
	/** @type {Map<string, string>} */
	const pngByLabel = new Map()
	for (const line of lines) {
		if (!line.startsWith('PNG')) continue
		const [, slide, png] = line.split('	')
		const label = PRSTGEOM_CASES[Number(slide) - 1]?.label
		if (label && png) pngByLabel.set(label, png)
	}
	if (pngByLabel.size !== PRSTGEOM_CASES.length) {
		failures.push(`preset geometry: exported ${pngByLabel.size} slides, expected ${PRSTGEOM_CASES.length}`)
		return failures
	}

	// Each slide's pixels as one comparable string. The shapes are flat fills at identical rects, so
	// an exact comparison is the right one -- there is no resampling or antialiasing difference to
	// tolerate between two renders of the same geometry.
	/** @type {Map<string, string>} */
	const pixels = new Map()
	try {
		for (const [label, png] of pngByLabel) {
			const img = decodePng(await fs.readFile(png))
			/** @type {number[]} */
			const flat = []
			for (let y = 0; y < img.h; y++) for (let x = 0; x < img.w; x++) flat.push(...img.rgb(x, y))
			pixels.set(label, `${img.w}x${img.h}:` + flat.join(','))
		}
	} catch (e) {
		failures.push(`preset geometry: could not read an exported PNG (${String(e)})`)
		return failures
	} finally {
		if (!KEEP) for (const png of pngByLabel.values()) await fs.rm(png, { force: true })
	}

	for (const testCase of PRSTGEOM_CASES) {
		const mine = pixels.get(testCase.label)
		if (testCase.same) {
			if (mine !== pixels.get(testCase.same)) {
				failures.push(
					`preset geometry: "${testCase.label}" paints differently from "${testCase.same}" — PowerPoint did NOT pin the out-of-range guide, so emitting one verbatim is no longer safe`
				)
			} else {
				console.log(`  OK  ${testCase.label} paints exactly as ${testCase.same}`)
			}
		}
		if (testCase.differs) {
			if (mine === pixels.get(testCase.differs)) {
				failures.push(
					`preset geometry: "${testCase.label}" paints identically to "${testCase.differs}", but they are two different IN-range values — the comparison cannot see a geometry change, so the equalities above prove nothing`
				)
			} else {
				console.log(`  OK  ${testCase.label} paints differently from ${testCase.differs} (comparison is sensitive)`)
			}
		}
	}
	return failures
}

// --- 5. orchestrate ---------------------------------------------------------
async function main() {
	/** @type {{label:string, file:string, generated:boolean, buildVbs:Function, verify:Function}[]} */
	/**
	 * One deck to drive: how to build its VBS, and how to read the result back.
	 * @typedef {object} Spec
	 * @property {string} label
	 * @property {string} file
	 * @property {boolean} generated whether this run created the deck and may delete it
	 * @property {(file: string) => string} buildVbs
	 * @property {(lines: string[]) => string[] | Promise<string[]>} verify
	 */
	/** @type {Spec[]} */
	const specs = []

	if (EXISTING_FILE) {
		// Corruption-open check only — no read-back verifier for an arbitrary deck.
		const fileSpec = {
			label: 'file',
			file: path.resolve(EXISTING_FILE),
			generated: false,
			buildVbs: /** @param {string} f */ (f) => vbsOpenHeader(f) + vbsFooter(),
			verify: () => [],
		}
		specs.push(fileSpec)
		console.log('Using deck: ' + fileSpec.file)
	} else {
		const navFile = await generateNavDeck()
		console.log('Generated nav deck: ' + navFile)
		specs.push({ label: 'nav', file: navFile, generated: true, buildVbs: buildNavVbs, verify: verifyNav })

		const geomFile = await generateGeomDeck()
		console.log('Generated custGeom deck: ' + geomFile)
		specs.push({ label: 'geom', file: geomFile, generated: true, buildVbs: buildGeomVbs, verify: verifyGeom })

		const oleFile = await generateOleDeck()
		console.log('Generated OLE deck: ' + oleFile)
		specs.push({ label: 'ole', file: oleFile, generated: true, buildVbs: buildOleVbs, verify: verifyOle })

		const prstGeomFile = await generatePresetGeomDeck()
		console.log('Generated preset-geometry deck: ' + prstGeomFile)
		specs.push({
			label: 'prstgeom',
			file: prstGeomFile,
			generated: true,
			buildVbs: buildPresetGeomVbs,
			verify: verifyPresetGeom,
		})

		const model3dFile = await generateModel3dDeck()
		console.log('Generated 3D-model deck: ' + model3dFile)
		specs.push({
			label: 'model3d',
			file: model3dFile,
			generated: true,
			buildVbs: buildModel3dVbs,
			verify: verifyModel3d,
		})
	}

	const failures = []
	for (const [index, spec] of specs.entries()) {
		const result = await driveDeck(spec.label, spec.file, spec.buildVbs)
		if (spec.generated && !KEEP) await fs.rm(spec.file, { force: true })
		else if (spec.generated) console.log('Kept deck: ' + spec.file)

		const lines = result.out.split(/\r?\n/).filter(Boolean)
		const open = checkOpen(spec.label, lines, result.out)
		if (open.skip) {
			// `NO_POWERPOINT` on the FIRST deck means this machine has no PowerPoint, which is the
			// SKIP this script is allowed to report. On any later deck it means something else --
			// a transient COM failure of the kind `driveDeck` already retries once for, or
			// PowerPoint dying part-way through -- and `process.exit(0)` inside the loop threw away
			// every failure the decks before it had already collected. The run reported SKIP and
			// passed, which is the one thing a gate must never do.
			if (index === 0) {
				process.exit(
					skipOrFail('TSPPTX_COM_SMOKE', 'PowerPoint is not installed / not COM-registered on this machine.')
				)
			}
			failures.push(
				`[${spec.label}] PowerPoint stopped answering COM part-way through the run ` +
					`(after ${index} deck(s) had been driven); the checks below it did not run.`
			)
			break
		}
		failures.push(...open.failures)
		if (!open.failures.length) failures.push(...(await spec.verify(lines)))
		if (result.err.trim()) console.error(`[${spec.label}] stderr: ` + result.err.trim())
	}

	if (failures.length) {
		console.error('\nPowerPoint COM smoke FAILED:')
		for (const f of failures) console.error('  - ' + f)
		process.exit(1)
	}
	console.log('\nPowerPoint COM smoke PASSED.')
}

main().catch((e) => {
	console.error('PowerPoint COM smoke errored: ' + (e?.stack || e))
	process.exit(1)
})
