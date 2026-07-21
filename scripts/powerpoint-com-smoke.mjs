#!/usr/bin/env node
/**
 * PowerPoint desktop COM smoke test (Windows only).
 *
 * CI's XML/schema checks can't catch two classes of desktop regression:
 *   1. Files that PowerPoint reports as corrupt and offers to "repair" (0x80070570),
 *      which the schema validator happily accepts.
 *   2. hlinkClick navigation actions that are schema-valid but that PowerPoint fails
 *      to parse back into a live slide-show action.
 *
 * This script drives the real PowerPoint application over COM (via cscript) to check
 * both. By default it generates a focused action-button navigation deck from the built
 * `dist/`, opens it headless, and reads each button's `ActionSettings(ppMouseClick).Action`
 * back out — asserting each jump resolved to the correct PpActionType enum, not a dead
 * `ppActionHyperlink`/`ppActionNone`. Point it at any deck with `--file` to run only the
 * corruption check.
 *
 *   node scripts/powerpoint-com-smoke.mjs                 # generate nav deck + verify actions
 *   node scripts/powerpoint-com-smoke.mjs --keep          # ...and keep the generated .pptx
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

// --- 1. deck under test -----------------------------------------------------
/** Build a focused nav deck from the built dist and return its path (+ expected button map). */
async function generateNavDeck() {
	const { default: pptxgen } = await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)
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

	const outFile = path.join(os.tmpdir(), `pptxgenjs-com-smoke-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return { file: outFile, expected: EXPECTED_ACTION }
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

// --- 3. the VBScript that drives PowerPoint ---------------------------------
function buildVbs(pptxFile) {
	// Emits tab-separated `ACTION` lines (slideIdx, objectName, resolvedActionNum) for the
	// caller to assert on, plus `OPEN_OK` / `OPEN_ERR` / `DONE` control lines.
	return `Option Explicit
Dim ppt, pres, sld, shp, act
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "NO_POWERPOINT\t" & Err.Description
  WScript.Quit 3
End If
Err.Clear
' WithWindow:=msoFalse (0) keeps it headless; ReadOnly avoids touching the file.
Set pres = ppt.Presentations.Open("${pptxFile.replace(/\\/g, '\\\\')}", -1, 0, 0)
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "OPEN_ERR\t" & Hex(Err.Number) & "\t" & Err.Description
  ppt.Quit
  WScript.Quit 2
End If
WScript.StdOut.WriteLine "OPEN_OK\t" & pres.Slides.Count
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
pres.Close
ppt.Quit
WScript.StdOut.WriteLine "DONE"
WScript.Quit 0
`
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

// --- 4. orchestrate ---------------------------------------------------------
async function main() {
	let file = EXISTING_FILE ? path.resolve(EXISTING_FILE) : null
	let expected = null
	let generated = false
	if (!file) {
		const deck = await generateNavDeck()
		file = deck.file
		expected = deck.expected
		generated = true
		console.log('Generated nav deck: ' + file)
	} else {
		console.log('Using deck: ' + file)
	}

	const vbsFile = path.join(os.tmpdir(), `pptxgenjs-com-smoke-${process.pid}.vbs`)
	await fs.writeFile(vbsFile, buildVbs(file))

	// cscript can transiently fail if PowerPoint is mid-launch; retry once.
	let result
	for (let attempt = 1; attempt <= 2; attempt++) {
		clearResiliency()
		result = await runCscript(vbsFile)
		if (result.code !== -1 && !/OPEN_ERR/.test(result.out)) break
		if (attempt === 1) console.log('First attempt failed; retrying once...')
	}

	// cleanup
	await fs.rm(vbsFile, { force: true })
	if (generated && !KEEP) await fs.rm(file, { force: true })
	else if (generated) console.log('Kept deck: ' + file)

	const lines = result.out.split(/\r?\n/).filter(Boolean)
	const failures = []

	if (/NO_POWERPOINT/.test(result.out)) {
		console.log('SKIP: PowerPoint is not installed / not COM-registered on this machine.')
		process.exit(0)
	}
	if (!lines.some((l) => l.startsWith('OPEN_OK'))) {
		const openErr = lines.find((l) => l.startsWith('OPEN_ERR')) || '(no OPEN_OK line)'
		failures.push('PowerPoint failed to open the deck (possible corruption / 0x80070570): ' + openErr)
	} else {
		console.log(lines.find((l) => l.startsWith('OPEN_OK')).replace('OPEN_OK\t', 'Opened OK; slide count = '))
	}

	// Read back each button's resolved action and compare to what we asked for.
	if (expected) {
		const seen = new Map()
		for (const l of lines) {
			if (!l.startsWith('ACTION\t')) continue
			const [, , name, num] = l.split('\t')
			seen.set(name, Number(num))
		}
		for (const [jump, wantNum] of Object.entries(expected)) {
			const got = seen.get(jump)
			if (got === undefined) failures.push(`button "${jump}": no navigation action read back (dead geometry?)`)
			else if (got !== wantNum) failures.push(`button "${jump}": resolved to PpActionType ${got}, expected ${wantNum}`)
			else console.log(`  OK  ${jump} -> PpActionType ${got}`)
		}
	}

	if (failures.length) {
		console.error('\nPowerPoint COM smoke FAILED:')
		for (const f of failures) console.error('  - ' + f)
		if (result.err.trim()) console.error('stderr: ' + result.err.trim())
		process.exit(1)
	}
	console.log('\nPowerPoint COM smoke PASSED.')
}

main().catch((e) => {
	console.error('PowerPoint COM smoke errored: ' + (e?.stack || e))
	process.exit(1)
})
