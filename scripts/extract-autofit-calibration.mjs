#!/usr/bin/env node
// Extract the autofit calibration table from the PowerPoint-authored fixture
// decks into test/read/fixtures/autofit-calibration.json.
//
// For each `test/read/fixtures/autofit-*.cases.json` manifest it pairs the
// sibling `.pptx` (the oracle), reads PowerPoint's baked outputs per case-id
// shape (ext.cy/off.y, normAutofit fontScale/lnSpcReduction, resolved typeface),
// and merges the LibreOffice cross-measure from `<deck>.lo.json` when present
// (produced on Windows by .tmp/measure-lo.py — LibreOffice is not a CI dep).
//
// The .pptx files remain the source of truth; this JSON is regenerable from them
// (PowerPoint columns are pure-Node/cross-platform; the LibreOffice column needs
// the Windows measurement step).
import { readFileSync, writeFileSync, existsSync, readdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..', 'test', 'read', 'fixtures')
const LO_DIR = process.argv.includes('--lo-dir')
	? process.argv[process.argv.indexOf('--lo-dir') + 1]
	: resolve(dirname(fileURLToPath(import.meta.url)), '..', '.tmp')

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', isArray: () => false })

const asArray = (v) => (v == null ? [] : Array.isArray(v) ? v : [v])

// Recursively collect every node under `tag` (namespaced key) in a parsed tree.
function collect(node, tag, out = []) {
	if (node == null || typeof node !== 'object') return out
	for (const [k, v] of Object.entries(node)) {
		if (k === tag) for (const item of asArray(v)) out.push(item)
		for (const item of asArray(v)) if (item && typeof item === 'object') collect(item, tag, out)
	}
	return out
}

function int(v) {
	return v == null ? null : parseInt(v, 10)
}

function readSlideShapes(xml) {
	const doc = parser.parse(xml)
	const shapes = collect(doc, 'p:sp')
	const byName = {}
	for (const sp of shapes) {
		const nv = sp['p:nvSpPr']?.['p:cNvPr']
		const name = nv?.['@_name']
		if (!name) continue
		const spPr = sp['p:spPr']
		const xfrm = spPr?.['a:xfrm']
		const off = xfrm?.['a:off']
		const ext = xfrm?.['a:ext']
		const txBody = sp['p:txBody']
		const bodyPr = txBody?.['a:bodyPr']
		const norm = bodyPr?.['a:normAutofit']
		const hasSp = bodyPr?.['a:spAutoFit'] !== undefined
		const autofit = norm !== undefined ? 'normAutofit' : hasSp ? 'spAutoFit' : 'none'
		// first run's resolved latin typeface + size
		const firstP = asArray(txBody?.['a:p'])[0]
		const firstR = asArray(firstP?.['a:r'])[0]
		const rPr = firstR?.['a:rPr']
		const latin = rPr?.['a:latin']
		byName[name] = {
			offXEmu: int(off?.['@_x']),
			offYEmu: int(off?.['@_y']),
			extCxEmu: int(ext?.['@_cx']),
			extCyEmu: int(ext?.['@_cy']),
			autofit,
			fontScale: norm && norm['@_fontScale'] != null ? int(norm['@_fontScale']) : null,
			lnSpcReduction: norm && norm['@_lnSpcReduction'] != null ? int(norm['@_lnSpcReduction']) : null,
			bodyWrap: bodyPr?.['@_wrap'] ?? null,
			bodyAnchor: bodyPr?.['@_anchor'] ?? null,
			lInsEmu: int(bodyPr?.['@_lIns']),
			tInsEmu: int(bodyPr?.['@_tIns']),
			rInsEmu: int(bodyPr?.['@_rIns']),
			bInsEmu: int(bodyPr?.['@_bIns']),
			resolvedTypeface: latin?.['@_typeface'] ?? null,
			runSizeHundredths: rPr?.['@_sz'] != null ? int(rPr['@_sz']) : null,
		}
	}
	return byName
}

function readDeck(pptxPath) {
	const zip = unzipSync(readFileSync(pptxPath))
	const slideNames = Object.keys(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => parseInt(a.match(/(\d+)/)[1], 10) - parseInt(b.match(/(\d+)/)[1], 10))
	const out = {}
	for (const n of slideNames) Object.assign(out, readSlideShapes(strFromU8(zip[n])))
	return out
}

// Flatten a case's primary inputs for the table (full case kept under `case`).
function caseInputs(c) {
	const firstRun = c.paragraphs?.[0]?.runs?.[0] ?? {}
	const firstPara = c.paragraphs?.[0] ?? {}
	return {
		kind: c.kind,
		font: firstRun.font ?? null,
		sizePt: firstRun.sizePt ?? null,
		bold: !!firstRun.bold,
		italic: !!firstRun.italic,
		wPt: c.wPt,
		hPt: c.hPt,
		wrap: !!c.wrap,
		anchor: c.anchor ?? null,
		insetsPt: c.insetsPt ?? null,
		lineSpacingPct: firstPara.lineSpacingPct ?? null,
		lineSpacingPts: firstPara.lineSpacingPts ?? null,
		spaceBeforePts: firstPara.spaceBeforePts ?? null,
		spaceAfterPts: firstPara.spaceAfterPts ?? null,
		charSpacingPts: firstRun.charSpacingPts ?? null,
		paragraphCount: c.paragraphs?.length ?? 0,
		runCount: (c.paragraphs ?? []).reduce((s, p) => s + (p.runs?.length ?? 0), 0),
		textSample: (firstRun.text ?? '').slice(0, 80),
	}
}

function main() {
	const manifests = readdirSync(FIX).filter((f) => /^autofit-.*\.cases\.json$/.test(f))
	if (!manifests.length) {
		console.error('no autofit-*.cases.json manifests found')
		process.exit(1)
	}
	// Preserve any previously-committed LibreOffice column so re-running on a box
	// without the (uncommitted) <deck>.lo.json files does not clobber it. The LO
	// measurement is a Windows+LibreOffice step (scripts/measure-autofit-lo.py).
	const outPath = resolve(FIX, 'autofit-calibration.json')
	const priorLo = {}
	if (existsSync(outPath)) {
		try {
			const prev = JSON.parse(readFileSync(outPath, 'utf8'))
			for (const d of prev.decks ?? []) for (const c of d.cases ?? []) if (c.libreoffice) priorLo[c.id] = c.libreoffice
		} catch {
			/* ignore a malformed prior file */
		}
	}
	const decks = []
	for (const mf of manifests.sort()) {
		const deckName = mf.replace(/\.cases\.json$/, '')
		const spec = JSON.parse(readFileSync(resolve(FIX, mf), 'utf8'))
		const pptx = resolve(FIX, `${deckName}.pptx`)
		if (!existsSync(pptx)) {
			console.error(`skip ${deckName}: ${deckName}.pptx missing`)
			continue
		}
		const pp = readDeck(pptx)
		const loPath = resolve(LO_DIR, `${deckName}.lo.json`)
		const lo = existsSync(loPath) ? JSON.parse(readFileSync(loPath, 'utf8')) : null
		const records = spec.cases.map((c) => {
			const ppOut = pp[c.id] ?? null
			const loOut = lo?.[c.id] ?? null
			const loMerged = loOut ? { hEmu: loOut.hEmu, hPt: loOut.hPt, wPt: loOut.wPt } : (priorLo[c.id] ?? null)
			return {
				id: c.id,
				slide: c.slide,
				note: c.note ?? null,
				inputs: caseInputs(c),
				powerpoint: ppOut,
				libreoffice: loMerged,
			}
		})
		decks.push({
			deck: deckName,
			notes: spec.notes ?? null,
			fontsRequired: spec.fontsRequired ?? null,
			libreofficeMeasured: !!lo || records.some((r) => r.libreoffice),
			caseCount: records.length,
			cases: records,
		})
		console.log(`${deckName}: ${records.length} cases, PP parsed, LO ${lo ? 'merged' : 'absent'}`)
	}
	const result = {
		schema: 'autofit-calibration@1',
		generatedFrom: 'PowerPoint-authored fixtures in test/read/fixtures/ (see README)',
		units: {
			emu: 'English Metric Units (914400/inch, 12700/pt)',
			fontScale: 'per-mille of original (100000 = 100%)',
			lnSpcReduction: 'per-mille reduction',
		},
		decks,
	}
	const out = resolve(FIX, 'autofit-calibration.json')
	writeFileSync(out, JSON.stringify(result, null, 2) + '\n')
	console.log(`wrote ${out}`)
}

main()
