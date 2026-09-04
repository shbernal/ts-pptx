/**
 * The five decks the PowerPoint COM smoke drives, built from the current `dist/`.
 *
 * Each one is a minimal deck carrying exactly the construct under test -- navigation actions,
 * custGeom connection sites, embedded OLE objects, an embedded 3D model, out-of-range preset
 * adjustment guides -- written to the OS temp directory and handed back by path. Nothing here
 * asserts: the verifiers in the entry point read PowerPoint's answers, and `contract.mjs` is what
 * the two sides agree on.
 *
 * `dist/node.js` is imported lazily rather than at module load, so importing this module (a test,
 * a `--help`) does not require a build.
 */

import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

import { solidPngBase64 } from '../png-utils.mjs'
import { ROOT } from '../script-utils.mjs'
import {
	EMPTY_ZIP_B64,
	GEOM_START_IDX,
	GEOM_TARGET_NAME,
	MODEL3D_FRAME_IN,
	MODEL3D_LAYOUT_IN,
	MODEL3D_PREVIEW_RGB,
	MODEL3D_SHAPE_NAME,
	PRSTGEOM_CASES,
	PRSTGEOM_FRAME_IN,
	PRSTGEOM_LAYOUT_IN,
} from './contract.mjs'

async function loadTsPptx() {
	const { default: TsPptx, ShapeType } = await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)
	return { TsPptx, ShapeType }
}

// --- 1a. navigation deck ----------------------------------------------------
/** Build a focused nav deck from the built dist and return its path. */
export async function generateNavDeck() {
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
export async function generateGeomDeck() {
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
export async function generateOleDeck() {
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
/** Build a deck with one embedded 3D model over a magenta preview. */
export async function generateModel3dDeck() {
	const { TsPptx } = await loadTsPptx()
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'SMOKE3D', width: MODEL3D_LAYOUT_IN.w, height: MODEL3D_LAYOUT_IN.h })
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

// --- 1e. preset-geometry adjustment deck ------------------------------------
/**
 * Build a deck with one shape per {@link PRSTGEOM_CASES} row, each on its own slide at the same
 * rect so the exported PNGs are directly comparable.
 *
 * The guides are written through `shapeAdjust`, whose `value` is scaled by 100000 on the way to
 * `<a:gd fmla="val N">` -- so the contract's raw units divide by that here rather than being
 * spelled twice.
 */
export async function generatePresetGeomDeck() {
	const { TsPptx } = await loadTsPptx()
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'PRSTGEOM', width: PRSTGEOM_LAYOUT_IN.w, height: PRSTGEOM_LAYOUT_IN.h })
	pptx.layout = 'PRSTGEOM'

	for (const testCase of PRSTGEOM_CASES) {
		pptx.addSlide().addShape(testCase.shape, {
			...PRSTGEOM_FRAME_IN,
			objectName: testCase.label,
			fill: { color: '4472C4' },
			line: { type: 'none' },
			shapeAdjust: Object.entries(testCase.adj).map(([name, raw]) => ({ name, value: raw / 100000 })),
		})
	}

	const outFile = path.join(os.tmpdir(), `ts-pptx-com-smoke-prstgeom-${process.pid}.pptx`)
	await pptx.writeFile({ fileName: outFile })
	return outFile
}
