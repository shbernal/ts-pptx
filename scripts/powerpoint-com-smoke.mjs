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
 *
 * This script drives the real PowerPoint application over COM (via cscript) to check
 * all five. By default it runs four generated decks from the built `dist/`:
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
 *     prove it drew one (no magenta ⇒ not the fallback; not blank ⇒ not an empty frame).
 * Point it at any deck with `--file` to run only the corruption-open check.
 *
 *   node scripts/powerpoint-com-smoke.mjs                 # nav + custGeom + OLE + 3D-model checks
 *   node scripts/powerpoint-com-smoke.mjs --keep          # ...and keep the generated .pptx files
 *   node scripts/powerpoint-com-smoke.mjs --file deck.pptx # corruption-open check on an existing deck
 *
 * Requirements: Windows with PowerPoint installed. No-ops with a clear message elsewhere.
 */
import os from 'node:os'
import fs from 'node:fs/promises'
import path from 'node:path'
import zlib from 'node:zlib'
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

// The OLE deck's objects, by `objectName` → the ProgID PowerPoint must resolve each back to.
// A schema-valid `<p:oleObj>` that PowerPoint quietly discards enumerates as *no shape at all*,
// so the read-back is the only thing proving the embedded object survived the round trip.
const EXPECTED_OLE_PROGID = {
	OlePackage: 'Excel.Sheet.12', // embedded OPC package → `.../package` rel
	OleBlob: 'Package', // generic OLE blob → `.../oleObject` rel, `.bin` part
}
/** An empty (but structurally valid) ZIP — enough of an "xlsx" for PowerPoint to bind the OLE server. */
const EMPTY_ZIP_B64 = 'UEsFBgAAAAAAAAAAAAAAAAAAAAAAAA=='

// The 3D-model deck. `msoShape3DModel` = 30 — a shape PowerPoint bound as a live model rather
// than leaving as an inert graphicFrame or collapsing to the fallback picture. The camera is the
// library's default, in metres, and PowerPoint exposes it as `Model3DFormat.CameraPositionZ`; a
// reading that far off means the `am3d:camera` subtree did not parse.
const MODEL3D_SHAPE_NAME = 'Cube3D'
const MODEL3D_SHAPE_TYPE = 30
const MODEL3D_EXPECTED_CAMERA_Z = 2.2630334
/**
 * The model deck's preview picture is solid magenta, which no lit gray model can produce. Type 30
 * plus a good camera still only proves PowerPoint resolved the *model*, not that it drew one — an
 * unreadable payload renders as the fallback or as nothing. So the deck is exported to PNG and the
 * pixels are read: any magenta means we are looking at the fallback, and an all-white frame means
 * nothing was drawn at all. Both absent ⇒ the `.glb` really rasterized.
 */
const MODEL3D_PREVIEW_RGB = [255, 0, 255]
/** Export size, and the frame's slide-relative rect (inches) — see `generateModel3dDeck`. */
const MODEL3D_EXPORT = { w: 960, h: 540 }
const MODEL3D_FRAME_IN = { x: 2, y: 1, w: 4, h: 3 }

async function loadTsPptx() {
	const { default: TsPptx, ShapeType } = await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)
	return { TsPptx, ShapeType }
}

// --- 1a. navigation deck ----------------------------------------------------
/** Build a focused nav deck from the built dist and return its path. */
async function generateNavDeck() {
	const { TsPptx } = await loadTsPptx()
	const pptx = new TsPptx()
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

	const outFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-nav-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
}

// --- 1b. custGeom connection-site deck --------------------------------------
/** Build a custGeom deck whose connector binds to the custom shape's connection site. */
async function generateGeomDeck() {
	const { TsPptx, ShapeType } = await loadTsPptx()
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'SMOKE', width: 10, height: 5.63 })
	pptx.layout = 'SMOKE'

	const s = pptx.addSlide()
	s.addText('custGeom connection sites', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true })

	// A freeform rectangle carrying real guides / connection sites / adjust handles — the same
	// shape the schema fixture `custGeom-connection-sites` builds.
	s.addShape(ShapeType.custGeom, {
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

	const outFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-geom-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
}

// --- 1c. OLE / embedded-object deck -----------------------------------------
/** Build a deck with one embedded-package and one generic-blob OLE object. */
async function generateOleDeck() {
	const { TsPptx } = await loadTsPptx()
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'SMOKE', width: 10, height: 5.63 })
	pptx.layout = 'SMOKE'

	const s = pptx.addSlide()
	s.addText('Embedded OLE objects', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true })
	// Embedded OPC package (xlsx): `.../package` rel + an `xlsx` content-type Default.
	s.addOleObject({ data: EMPTY_ZIP_B64, extn: 'xlsx', objectName: 'OlePackage', x: 0.5, y: 1.5, w: 3, h: 2 })
	// Generic OLE-server blob: `.../oleObject` rel + a `.bin` part.
	s.addOleObject({ data: 'AAECAwQFBgc=', objectName: 'OleBlob', x: 4.5, y: 1.5, w: 2, h: 2 })

	const outFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-ole-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
}

// --- 1d. 3D-model deck ------------------------------------------------------
/**
 * A 1x1 solid-colour PNG, built here so nothing about the preview can be mistaken for model pixels.
 * @param {readonly number[]} rgb - the fill colour
 */
function solidPngBase64(rgb) {
	const raw = Buffer.from([0, rgb[0] ?? 0, rgb[1] ?? 0, rgb[2] ?? 0])
	const crcTable = Int32Array.from({ length: 256 }, (_, n) => {
		let c = n
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		return c
	})
	const crc = (buf) => {
		let c = -1
		for (const byte of buf) c = (crcTable[(c ^ byte) & 0xff] ?? 0) ^ (c >>> 8)
		return (c ^ -1) >>> 0
	}
	const chunk = (type, data) => {
		const len = Buffer.alloc(4)
		len.writeUInt32BE(data.length)
		const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
		const crcBuf = Buffer.alloc(4)
		crcBuf.writeUInt32BE(crc(body))
		return Buffer.concat([len, body, crcBuf])
	}
	const ihdr = Buffer.alloc(13)
	ihdr.writeUInt32BE(1, 0)
	ihdr.writeUInt32BE(1, 4)
	ihdr[8] = 8 // bit depth
	ihdr[9] = 2 // colour type: truecolour
	return Buffer.concat([
		Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
		chunk('IHDR', ihdr),
		chunk('IDAT', zlib.deflateSync(raw)),
		chunk('IEND', Buffer.alloc(0)),
	]).toString('base64')
}

/** Build a deck with one embedded 3D model over a magenta preview. */
async function generateModel3dDeck() {
	const { TsPptx } = await loadTsPptx()
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'SMOKE3D', width: 10, height: 5.625 })
	pptx.layout = 'SMOKE3D'

	const cube = await fs.readFile(path.join(ROOT, 'test', 'read', 'fixtures', 'authoring', 'assets', 'cube.glb'))
	pptx.addSlide().addModel3d({
		data: cube.toString('base64'),
		preview: { data: 'image/png;base64,' + solidPngBase64(MODEL3D_PREVIEW_RGB) },
		objectName: MODEL3D_SHAPE_NAME,
		...MODEL3D_FRAME_IN,
	})

	const outFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-model3d-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
}

/**
 * Decode an 8-bit PNG (truecolour, palette, or grey, +/- alpha) into `{w, h, rgb(x,y)}`.
 *
 * Hand-rolled rather than pulled in as a dependency: this script is the only consumer, PowerPoint
 * writes whichever of those colour types suits the slide (truecolour for the rendered model, an
 * indexed palette for a blank one), and `node:zlib` already does the hard part.
 * @param {Buffer} bytes - the PNG file
 */
function decodePng(bytes) {
	/** Indexed read that satisfies `noUncheckedIndexedAccess`; out-of-range is 0, as for a `Buffer`. */
	const at = (/** @type {Uint8Array} */ buf, /** @type {number} */ i) => buf[i] ?? 0
	let off = 8
	let w = 0
	let h = 0
	let colour = 0
	/** @type {Buffer | null} */
	let palette = null
	/** @type {Buffer[]} */
	const idat = []
	while (off + 8 <= bytes.length) {
		const len = bytes.readUInt32BE(off)
		const type = bytes.toString('ascii', off + 4, off + 8)
		const data = bytes.subarray(off + 8, off + 8 + len)
		if (type === 'IHDR') {
			w = data.readUInt32BE(0)
			h = data.readUInt32BE(4)
			if (at(data, 8) !== 8) throw new Error(`unsupported PNG bit depth ${at(data, 8)}`)
			colour = at(data, 9)
			if (at(data, 12) !== 0) throw new Error('interlaced PNG not supported')
		} else if (type === 'PLTE') palette = Buffer.from(data)
		else if (type === 'IDAT') idat.push(data)
		else if (type === 'IEND') break
		off += 12 + len
	}
	const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colour]
	if (!channels) throw new Error(`unsupported PNG colour type ${colour}`)
	if (colour === 3 && !palette) throw new Error('indexed PNG with no PLTE chunk')
	const pal = palette ?? Buffer.alloc(0)
	const raw = zlib.inflateSync(Buffer.concat(idat))
	const stride = w * channels
	const out = Buffer.alloc(h * stride)
	let pos = 0
	for (let y = 0; y < h; y++) {
		const filter = at(raw, pos++)
		const line = raw.subarray(pos, pos + stride)
		pos += stride
		const cur = out.subarray(y * stride, (y + 1) * stride)
		const prior = y > 0 ? out.subarray((y - 1) * stride, y * stride) : Buffer.alloc(stride)
		for (let i = 0; i < stride; i++) {
			const a = i >= channels ? at(cur, i - channels) : 0
			const b = at(prior, i)
			const c = i >= channels ? at(prior, i - channels) : 0
			let v = at(line, i)
			if (filter === 1) v += a
			else if (filter === 2) v += b
			else if (filter === 3) v += (a + b) >> 1
			else if (filter === 4) {
				const p = a + b - c
				const pa = Math.abs(p - a)
				const pb = Math.abs(p - b)
				const pc = Math.abs(p - c)
				v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c
			}
			cur[i] = v & 0xff
		}
	}
	/**
	 * @param {number} x
	 * @param {number} y
	 * @returns {[number, number, number]}
	 */
	const rgb = (x, y) => {
		const i = y * stride + x * channels
		if (colour === 3) {
			const p = at(out, i) * 3
			return [at(pal, p), at(pal, p + 1), at(pal, p + 2)]
		}
		if (colour === 0 || colour === 4) return [at(out, i), at(out, i), at(out, i)]
		return [at(out, i), at(out, i + 1), at(out, i + 2)]
	}
	return { w, h, rgb }
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
/**
 * Shared header: create the app, open the deck, emit OPEN_OK/OPEN_ERR.
 *
 * Default is headless + read-only. Pass `withWindow` for checks that read OLE objects back:
 * a windowless PowerPoint does not instantiate embedded objects, so `Shapes` comes back without
 * them — indistinguishable from PowerPoint having discarded them. (A deck PowerPoint authored
 * itself enumerates zero shapes under a headless open too, which is how that was pinned down.)
 */
function vbsOpenHeader(pptxFile, withWindow = false) {
	// WithWindow:=msoFalse (0) keeps it headless; ReadOnly avoids touching the file.
	const openArgs = withWindow ? '0, 0, -1' : '-1, 0, 0'
	return `Option Explicit
Dim ppt, pres, sld, shp
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "NO_POWERPOINT\t" & Err.Description
  WScript.Quit 3
End If
Err.Clear
Set pres = ppt.Presentations.Open("${pptxFile.replace(/\\/g, '\\\\')}", ${openArgs})
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

function buildOleVbs(pptxFile) {
	// Emits one tab-separated `OLE` line per embedded-object shape: name, resolved ProgID.
	// msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10.
	return (
		vbsOpenHeader(pptxFile, true) +
		`Dim i, j
For i = 1 To pres.Slides.Count
  Set sld = pres.Slides(i)
  For j = 1 To sld.Shapes.Count
    Set shp = sld.Shapes(j)
    If shp.Type = 7 Or shp.Type = 10 Then
      WScript.StdOut.WriteLine "OLE\t" & shp.Name & "\t" & shp.OLEFormat.ProgID
    End If
  Next
Next
` +
		vbsFooter()
	)
}

function buildModel3dVbs(pptxFile) {
	// Emits `M3D <name> <Shape.Type> <Model3D.CameraPositionZ>` per shape, then exports slide 1 to
	// PNG so the caller can prove the model actually rasterized. A window is needed here for the
	// same reason as OLE: a headless PowerPoint does not instantiate the 3D renderer.
	const png = pptxFile.replace(/\.pptx$/i, '.png').replace(/\\/g, '\\\\')
	return (
		vbsOpenHeader(pptxFile, true) +
		`Dim j, camZ
Set sld = pres.Slides(1)
For j = 1 To sld.Shapes.Count
  Set shp = sld.Shapes(j)
  camZ = ""
  On Error Resume Next
  camZ = shp.Model3D.CameraPositionZ
  On Error Goto 0
  WScript.StdOut.WriteLine "M3D" & vbTab & shp.Name & vbTab & shp.Type & vbTab & camZ
Next
Err.Clear
sld.Export "${png}", "PNG", ${MODEL3D_EXPORT.w}, ${MODEL3D_EXPORT.h}
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "EXPORT_ERR" & vbTab & Hex(Err.Number) & vbTab & Err.Description
Else
  WScript.StdOut.WriteLine "EXPORT" & vbTab & "${png}"
End If
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
	const vbsFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-${label}-${process.pid}.vbs`)
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

function verifyOle(lines) {
	const failures = []
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
 */
async function verifyModel3d(lines) {
	const failures = []
	const row = lines.map((l) => l.split('\t')).find((p) => p[0] === 'M3D' && p[1] === MODEL3D_SHAPE_NAME)
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
	let img
	try {
		img = decodePng(await fs.readFile(pngPath))
	} catch (e) {
		failures.push(`3D model: could not read the exported PNG (${String(e)})`)
		return failures
	} finally {
		if (!KEEP) await fs.rm(pngPath, { force: true })
	}

	// The frame, in exported pixels. The deck's layout is 10 x 5.625 in, exported at 960 x 540.
	const sx = img.w / 10
	const sy = img.h / 5.625
	const x0 = Math.round(MODEL3D_FRAME_IN.x * sx)
	const y0 = Math.round(MODEL3D_FRAME_IN.y * sy)
	const x1 = Math.min(img.w - 1, Math.round((MODEL3D_FRAME_IN.x + MODEL3D_FRAME_IN.w) * sx))
	const y1 = Math.min(img.h - 1, Math.round((MODEL3D_FRAME_IN.y + MODEL3D_FRAME_IN.h) * sy))
	let fallbackPx = 0
	let drawnPx = 0
	for (let y = y0; y <= y1; y++) {
		for (let x = x0; x <= x1; x++) {
			const [r, g, b] = img.rgb(x, y)
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

// --- 5. orchestrate ---------------------------------------------------------
async function main() {
	/** @type {{label:string, file:string, generated:boolean, buildVbs:Function, verify:Function}[]} */
	const specs = []

	if (EXISTING_FILE) {
		// Corruption-open check only — no read-back verifier for an arbitrary deck.
		const fileSpec = {
			label: 'file',
			file: path.resolve(EXISTING_FILE),
			generated: false,
			buildVbs: (f) => vbsOpenHeader(f) + vbsFooter(),
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
