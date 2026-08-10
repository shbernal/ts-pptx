// Extract the autofit calibration table from the PowerPoint-authored fixture
// decks into the parent directory's autofit-calibration.json.
//
// For each `../autofit-*.cases.json` manifest it pairs the sibling `.pptx` (the
// oracle), reads PowerPoint's baked outputs per case-id shape (ext.cy/off.y,
// normAutofit fontScale/lnSpcReduction, resolved typeface), and merges the
// LibreOffice cross-measure from `<deck>.lo.json` when present (produced on
// Windows by the sibling measure-lo.py — LibreOffice is not a CI dep).
//
// The .pptx files remain the source of truth; this JSON is regenerable from them.
// Unlike the author-*.ps1 recipes here, this one needs no desktop app: the
// PowerPoint columns are pure-Node/cross-platform, and only the LibreOffice
// column depends on the Windows measurement step.
import { readFileSync, writeFileSync, existsSync, readdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'
import { DOMParser, MIME_TYPE, onErrorStopParsing } from '@xmldom/xmldom'
import { parseCliOrExit, ROOT } from '../../../../scripts/script-utils.mjs'

// This script lives in test/read/fixtures/authoring/, so the fixtures dir is its parent.
const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..')

const USAGE = `Extract the autofit calibration table from the PowerPoint-authored fixtures.

  node test/read/fixtures/authoring/extract-autofit-calibration.mjs
  node test/read/fixtures/authoring/extract-autofit-calibration.mjs --lo-dir .tmp

Options:
  --lo-dir <path>  where the LibreOffice cross-measure <deck>.lo.json files live
                   (default <repo>/.tmp)
  -h, --help       show this message`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: { 'lo-dir': { type: 'string' } },
})
const LO_DIR = values['lo-dir'] ?? resolve(ROOT, '.tmp')

// The two namespaces this script reads. `@xmldom/xmldom` is already a runtime
// dependency of the library, so no parser is pulled in for this script alone.
const NS = {
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
}
const ELEMENT_NODE = 1

function splitQName(qname) {
	const colon = qname.indexOf(':')
	const uri = NS[qname.slice(0, colon)]
	if (!uri) throw new Error(`unknown namespace prefix in qname: ${qname}`)
	return { uri, local: qname.slice(colon + 1) }
}

/**
 * The first *direct child* element matching a prefixed qname, or `null`.
 *
 * Deliberately not `getElementsByTagNameNS`, which searches descendants: every
 * lookup below is a child step (`p:spPr` → `a:xfrm` → `a:off`), and a descendant
 * search would happily match an `a:off` belonging to some nested subtree.
 */
function child(parent, qname) {
	if (!parent) return null
	const { uri, local } = splitQName(qname)
	for (let node = parent.firstChild; node; node = node.nextSibling)
		if (node.nodeType === ELEMENT_NODE && node.localName === local && node.namespaceURI === uri) return node
	return null
}

/** Every descendant element matching a prefixed qname, in document order. */
function descendants(root, qname) {
	const { uri, local } = splitQName(qname)
	const list = root.getElementsByTagNameNS(uri, local)
	const out = []
	for (let i = 0; i < list.length; i++) out.push(list[i])
	return out
}

/** An unprefixed attribute value, or `null` when the element or attribute is absent. */
function attr(element, name) {
	if (!element || !element.hasAttribute(name)) return null
	return element.getAttribute(name)
}

function int(v) {
	return v == null ? null : parseInt(v, 10)
}

function readSlideShapes(xml) {
	const doc = new DOMParser({ onError: onErrorStopParsing }).parseFromString(xml, MIME_TYPE.XML_TEXT)
	const byName = {}
	// Shapes nest (a group's children are `p:sp` too), so this one lookup is a
	// descendant search — the field reads below are all child steps.
	for (const sp of descendants(doc, 'p:sp')) {
		const name = attr(child(child(sp, 'p:nvSpPr'), 'p:cNvPr'), 'name')
		if (!name) continue
		const xfrm = child(child(sp, 'p:spPr'), 'a:xfrm')
		const off = child(xfrm, 'a:off')
		const ext = child(xfrm, 'a:ext')
		const txBody = child(sp, 'p:txBody')
		const bodyPr = child(txBody, 'a:bodyPr')
		const norm = child(bodyPr, 'a:normAutofit')
		const autofit = norm ? 'normAutofit' : child(bodyPr, 'a:spAutoFit') ? 'spAutoFit' : 'none'
		// first run's resolved latin typeface + size
		const rPr = child(child(child(txBody, 'a:p'), 'a:r'), 'a:rPr')
		const latin = child(rPr, 'a:latin')
		byName[name] = {
			offXEmu: int(attr(off, 'x')),
			offYEmu: int(attr(off, 'y')),
			extCxEmu: int(attr(ext, 'cx')),
			extCyEmu: int(attr(ext, 'cy')),
			autofit,
			fontScale: int(attr(norm, 'fontScale')),
			lnSpcReduction: int(attr(norm, 'lnSpcReduction')),
			bodyWrap: attr(bodyPr, 'wrap'),
			bodyAnchor: attr(bodyPr, 'anchor'),
			lInsEmu: int(attr(bodyPr, 'lIns')),
			tInsEmu: int(attr(bodyPr, 'tIns')),
			rInsEmu: int(attr(bodyPr, 'rIns')),
			bInsEmu: int(attr(bodyPr, 'bIns')),
			resolvedTypeface: attr(latin, 'typeface'),
			runSizeHundredths: int(attr(rPr, 'sz')),
		}
	}
	return byName
}

function readDeck(pptxPath) {
	const zip = unzipSync(readFileSync(pptxPath))
	const slideNames = Object.keys(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => parseInt(a.match(/(\d+)/)?.[1] ?? '0', 10) - parseInt(b.match(/(\d+)/)?.[1] ?? '0', 10))
	const out = {}
	for (const n of slideNames) {
		const data = zip[n]
		if (data) Object.assign(out, readSlideShapes(strFromU8(data)))
	}
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
	// measurement is a Windows+LibreOffice step (the sibling measure-lo.py).
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
	// Tab-indented to match what the repo formatter leaves in the committed file, so
	// a regeneration that changed nothing produces no diff at all.
	writeFileSync(outPath, JSON.stringify(result, null, '\t') + '\n')
	console.log(`wrote ${outPath}`)
}

main()
