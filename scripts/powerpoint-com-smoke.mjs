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
 *
 * This script drives the real PowerPoint application over COM (via cscript) to check
 * all three. By default it runs two generated decks from the built `dist/`:
 *   - a navigation deck, reading each button's `ActionSettings(ppMouseClick).Action`
 *     back out and asserting each jump resolved to the correct PpActionType enum; and
 *   - a custGeom deck with connection sites + a connector bound to site index 1, reading
 *     the connector's `ConnectorFormat` back out and asserting it begin-connected to the
 *     custom shape at a real connection site.
 * Point it at any deck with `--file` to run only the corruption-open check.
 *
 *   node scripts/powerpoint-com-smoke.mjs                 # generated nav + custGeom checks
 *   node scripts/powerpoint-com-smoke.mjs --keep          # ...and keep the generated .pptx files
 *   node scripts/powerpoint-com-smoke.mjs --file deck.pptx # corruption-open check on an existing deck
 *
 * Requirements: Windows with PowerPoint installed. No-ops with a clear message elsewhere.
 */
import os from 'node:os'
import fs from 'node:fs/promises'
import path from 'node:path'
import { spawn, spawnSync } from 'node:child_process'
import { fileURLToPath, pathToFileURL } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const ROOT = path.resolve(__dirname, '..')

// --- args -------------------------------------------------------------------
const argv = process.argv.slice(2)
const KEEP = argv.includes('--keep')
const fileArgIdx = argv.indexOf('--file')
const EXISTING_FILE = fileArgIdx !== -1 ? argv[fileArgIdx + 1] : null

if (os.platform() !== 'win32') {
	console.log('SKIP: PowerPoint COM smoke is Windows-only (platform: ' + os.platform() + ').')
	process.exit(0)
}

// PpActionType enum values PowerPoint resolves `ppaction://hlinkshowjump?jump=<x>` into.
// The whole point of the read-back: prove the jump became a live navigation action, not a
// dead ppActionHyperlink (7) / ppActionNone (0).
const EXPECTED_ACTION = {
	nextslide: 1, // ppActionNextSlide
	previousslide: 2, // ppActionPreviousSlide
	firstslide: 3, // ppActionFirstSlide
	lastslide: 4, // ppActionLastSlide
	lastslideviewed: 5, // ppActionLastSlideViewed
	endshow: 6, // ppActionEndShow
}

// The custGeom deck binds the connector's start to connection site index 1 (0-based, OOXML
// `<a:stCxn idx="1">`) of the shape named `freeform`. PowerPoint numbers connection sites
// 1-based, so the resolved `BeginConnectionSite` is idx+1 = 2.
const GEOM_TARGET_NAME = 'freeform'
const GEOM_START_IDX = 1
const GEOM_EXPECTED_SITE = GEOM_START_IDX + 1

async function loadPptxgen() {
	const { default: pptxgen } = await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)
	return pptxgen
}

// --- 1a. navigation deck ----------------------------------------------------
/** Build a focused nav deck from the built dist and return its path. */
async function generateNavDeck() {
	const pptxgen = await loadPptxgen()
	const pptx = new pptxgen()
	pptx.defineLayout({ name: 'SMOKE', width: 10, height: 5.63 })
	pptx.layout = 'SMOKE'

	// A few content slides so every jump has somewhere to land.
	for (let i = 1; i <= 3; i++) {
		const s = pptx.addSlide()
		s.addText(`Content slide ${i}`, { x: 0.5, y: 0.4, w: 9, h: 0.8, fontSize: 28, bold: true })
	}

	// One button per navigation jump. `objectName` == the jump value, so the COM read-back
	// can map each shape's resolved Action back to what we asked for.
	const nav = pptx.addSlide()
	nav.addText('Navigation buttons', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true })
	const buttons = [
		['actionButtonBeginning', 'firstslide', 'Home'],
		['actionButtonBackPrevious', 'previousslide', 'Back'],
		['actionButtonForwardNext', 'nextslide', 'Next'],
		['actionButtonEnd', 'lastslide', 'End'],
		['actionButtonReturn', 'lastslideviewed', 'Return'],
		['actionButtonEnd', 'endshow', 'Stop'],
	]
	buttons.forEach(([shape, action, tooltip], i) => {
		nav.addShape(shape, {
			x: 0.5 + i * 1.5,
			y: 2,
			w: 1.1,
			h: 1.1,
			objectName: action,
			hyperlink: { action, tooltip },
		})
	})

	const outFile = path.join(os.tmpdir(), `pptxgenjs-com-smoke-nav-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
}

// --- 1b. custGeom connection-site deck --------------------------------------
/** Build a custGeom deck whose connector binds to the custom shape's connection site. */
async function generateGeomDeck() {
	const pptxgen = await loadPptxgen()
	const pptx = new pptxgen()
	pptx.defineLayout({ name: 'SMOKE', width: 10, height: 5.63 })
	pptx.layout = 'SMOKE'

	const s = pptx.addSlide()
	s.addText('custGeom connection sites', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true })

	// A freeform rectangle carrying real guides / connection sites / adjust handles — the same
	// shape the schema fixture `custGeom-connection-sites` builds.
	s.addShape(pptx.ShapeType.custGeom, {
		x: 1,
		y: 1.5,
		w: 3,
		h: 2,
		objectName: GEOM_TARGET_NAME,
		fill: { color: '4472C4' },
		line: { color: '2F528F', width: 1 },
		points: [
			{ x: 0, y: 0 },
			{ x: 3, y: 0 },
			{ x: 3, y: 2 },
			{ x: 0, y: 2, close: true },
		],
		guides: [{ name: 'w2', formula: '*/ w 1 2' }],
		connectionSites: [
			{ ang: 0, x: 3, y: 1 }, // idx 0: right-middle
			{ ang: 90, x: 'w2', y: 0 }, // idx 1: top, at the guide-driven x
			{ ang: 180, x: 0, y: 1 }, // idx 2: left-middle
		],
		adjustHandles: [
			{ x: 'w2', y: 0, gdRefX: 'w2', minX: 0, maxX: 3 },
			{ x: 3, y: 2, gdRefAng: 'w2', minAng: 0, maxAng: 90 },
		],
	})

	// Bind an elbow connector's start to connection site #1 of the freeform shape.
	s.addConnector({
		type: 'elbow',
		x1: 4,
		y1: 2.5,
		x2: 7,
		y2: 4,
		color: 'FF0000',
		width: 2,
		startShape: GEOM_TARGET_NAME,
		startShapeIdx: GEOM_START_IDX,
	})

	const outFile = path.join(os.tmpdir(), `pptxgenjs-com-smoke-geom-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
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

// --- 3. the VBScripts that drive PowerPoint ---------------------------------
/** Shared header: create the app, open the deck headless+read-only, emit OPEN_OK/OPEN_ERR. */
function vbsOpenHeader(pptxFile) {
	// WithWindow:=msoFalse (0) keeps it headless; ReadOnly avoids touching the file.
	return `Option Explicit
Dim ppt, pres, sld, shp
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "NO_POWERPOINT\t" & Err.Description
  WScript.Quit 3
End If
Err.Clear
Set pres = ppt.Presentations.Open("${pptxFile.replace(/\\/g, '\\\\')}", -1, 0, 0)
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "OPEN_ERR\t" & Hex(Err.Number) & "\t" & Err.Description
  ppt.Quit
  WScript.Quit 2
End If
WScript.StdOut.WriteLine "OPEN_OK\t" & pres.Slides.Count
`
}

function vbsFooter() {
	return `pres.Close
ppt.Quit
WScript.StdOut.WriteLine "DONE"
WScript.Quit 0
`
}

function buildNavVbs(pptxFile) {
	// Emits tab-separated `ACTION` lines (slideIdx, objectName, resolvedActionNum).
	return (
		vbsOpenHeader(pptxFile) +
		`Dim act
For Each sld In pres.Slides
  For Each shp In sld.Shapes
    Set act = shp.ActionSettings(1) ' ppMouseClick = 1
    If Not act Is Nothing Then
      If act.Action <> 0 Then
        WScript.StdOut.WriteLine "ACTION\t" & sld.SlideIndex & "\t" & shp.Name & "\t" & act.Action
      End If
    End If
  Next
Next
` +
		vbsFooter()
	)
}

function buildGeomVbs(pptxFile) {
	// Emits one tab-separated `CONN` line per connector shape:
	//   CONN <name> <beginConnected(-1/0)> <beginConnectedShapeName> <beginConnectionSite>
	return (
		vbsOpenHeader(pptxFile) +
		`Dim cf, tgt
For Each sld In pres.Slides
  For Each shp In sld.Shapes
    If shp.Connector Then
      Set cf = shp.ConnectorFormat
      tgt = ""
      If cf.BeginConnected Then tgt = cf.BeginConnectedShape.Name
      WScript.StdOut.WriteLine "CONN\t" & shp.Name & "\t" & cf.BeginConnected & "\t" & tgt & "\t" & cf.BeginConnectionSite
    End If
  Next
Next
` +
		vbsFooter()
	)
}

function runCscript(vbsFile) {
	return new Promise((resolve) => {
		const child = spawn('cscript', ['//nologo', '//B', vbsFile], { stdio: ['ignore', 'pipe', 'pipe'] })
		let out = ''
		let err = ''
		child.stdout.on('data', (d) => (out += d))
		child.stderr.on('data', (d) => (err += d))
		child.on('close', (code) => resolve({ code, out, err }))
		child.on('error', (e) => resolve({ code: -1, out, err: String(e) }))
	})
}

/** Write a VBS for `file`, drive PowerPoint (retry once), and return the raw cscript result. */
async function driveDeck(label, file, buildVbs) {
	const vbsFile = path.join(os.tmpdir(), `pptxgenjs-com-smoke-${label}-${process.pid}.vbs`)
	await fs.writeFile(vbsFile, buildVbs(file))
	// cscript can transiently fail if PowerPoint is mid-launch; retry once.
	let result
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
/** Returns { skip, failures }. Handles the NO_POWERPOINT / OPEN_ERR / OPEN_OK common cases. */
function checkOpen(label, lines, out) {
	if (/NO_POWERPOINT/.test(out)) return { skip: true, failures: [] }
	if (!lines.some((l) => l.startsWith('OPEN_OK'))) {
		const openErr = lines.find((l) => l.startsWith('OPEN_ERR')) || '(no OPEN_OK line)'
		return {
			skip: false,
			failures: [`[${label}] PowerPoint failed to open the deck (possible corruption / 0x80070570): ${openErr}`],
		}
	}
	console.log(
		`[${label}] ` + lines.find((l) => l.startsWith('OPEN_OK')).replace('OPEN_OK\t', 'opened OK; slide count = ')
	)
	return { skip: false, failures: [] }
}

function verifyNav(lines) {
	const failures = []
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

function verifyGeom(lines) {
	const failures = []
	const conns = lines.filter((l) => l.startsWith('CONN\t')).map((l) => l.split('\t'))
	if (!conns.length) {
		failures.push('no connector shape found on the custGeom deck (connector dropped?)')
		return failures
	}
	// The deck has exactly one bound connector.
	const [, name, connected, targetName, site] = conns[0]
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

// --- 5. orchestrate ---------------------------------------------------------
async function main() {
	/** @type {{label:string, file:string, generated:boolean, buildVbs:Function, verify:Function}[]} */
	const specs = []

	if (EXISTING_FILE) {
		// Corruption-open check only — no read-back verifier for an arbitrary deck.
		specs.push({
			label: 'file',
			file: path.resolve(EXISTING_FILE),
			generated: false,
			buildVbs: (f) => vbsOpenHeader(f) + vbsFooter(),
			verify: () => [],
		})
		console.log('Using deck: ' + specs[0].file)
	} else {
		const navFile = await generateNavDeck()
		console.log('Generated nav deck: ' + navFile)
		specs.push({ label: 'nav', file: navFile, generated: true, buildVbs: buildNavVbs, verify: verifyNav })

		const geomFile = await generateGeomDeck()
		console.log('Generated custGeom deck: ' + geomFile)
		specs.push({ label: 'geom', file: geomFile, generated: true, buildVbs: buildGeomVbs, verify: verifyGeom })
	}

	const failures = []
	for (const spec of specs) {
		const result = await driveDeck(spec.label, spec.file, spec.buildVbs)
		if (spec.generated && !KEEP) await fs.rm(spec.file, { force: true })
		else if (spec.generated) console.log('Kept deck: ' + spec.file)

		const lines = result.out.split(/\r?\n/).filter(Boolean)
		const open = checkOpen(spec.label, lines, result.out)
		if (open.skip) {
			console.log('SKIP: PowerPoint is not installed / not COM-registered on this machine.')
			process.exit(0)
		}
		failures.push(...open.failures)
		if (!open.failures.length) failures.push(...spec.verify(lines))
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
