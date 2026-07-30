import { readFile, writeFile } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import JSZip from 'jszip'

// This script lives in test/read/fixtures/authoring/, so the fixtures dir is its parent.
const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..')

const fixture = resolve(FIX, 'slide-animation-presets.pptx')
const bytes = await readFile(fixture)
const sha256 = createHash('sha256').update(bytes).digest('hex')
const zip = await JSZip.loadAsync(bytes)
const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')

// --- verbatim full timing tree (same extraction the regression test uses) ---
const timingXml = slideXml.match(/<p:timing>[\s\S]*<\/p:timing>/)[0]

// --- shapeIds: id -> name from each <p:cNvPr id=".." name=".."> on a top-level sp ---
const shapeIds = {}
for (const m of slideXml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)) shapeIds[m[1]] = m[2]

// --- balanced <p:cTn ...> ... </p:cTn> extractor ---
function elementAt(s, startIdx, tag) {
	// startIdx points at '<tag'
	let i = startIdx
	let depth = 0
	while (i < s.length) {
		// find next relevant token from i
		const nextOpen = s.indexOf(`<${tag}`, i)
		const nextClose = s.indexOf(`</${tag}>`, i)
		// also detect self-closing of the immediate open
		if (nextOpen !== -1 && (nextClose === -1 || nextOpen < nextClose)) {
			// determine if this open tag is self-closing
			const gt = s.indexOf('>', nextOpen)
			const selfClose = s[gt - 1] === '/'
			if (!selfClose) depth++
			i = gt + 1
			if (selfClose && depth === 0) return s.slice(startIdx, gt + 1)
		} else if (nextClose !== -1) {
			depth--
			i = nextClose + `</${tag}>`.length
			if (depth === 0) return s.slice(startIdx, i)
		} else {
			break
		}
	}
	throw new Error('unbalanced ' + tag)
}

function innerOf(elXml, tag) {
	const open = elXml.indexOf(`<${tag}`)
	const full = elementAt(elXml, open, tag)
	const firstGt = full.indexOf('>')
	if (full[firstGt - 1] === '/') return '' // self-closing
	return full.slice(firstGt + 1, full.length - `</${tag}>`.length)
}

// --- per-preset effect nodes: every <p:cTn ... presetID=...> ---
const presetByName = {
	'entr-fadeIn': 'fadeIn',
	'entr-flyIn': 'flyIn',
	'entr-appear': 'appear',
	'entr-wipe': 'wipe',
	'emph-grow': 'grow',
	'emph-spin': 'spin',
	'exit-fadeOut': 'fadeOut',
	'exit-flyOut': 'flyOut',
}

const effects = []
const presetTemplates = {}
const cTnOpenRe =
	/<p:cTn id="\d+" presetID="(\d+)" presetClass="(entr|emph|exit)" presetSubtype="(\d+)" fill="hold" grpId="0" nodeType="(\w+)">/g
for (const m of timingXml.matchAll(cTnOpenRe)) {
	const cTnFull = elementAt(timingXml, m.index, 'p:cTn')
	const spid = cTnFull.match(/<p:spTgt spid="(\d+)"/)[1]
	const childTnLst = innerOf(cTnFull, 'p:childTnLst')
	const shapeName = shapeIds[spid]
	const presetName = presetByName[shapeName]
	const key = { presetID: Number(m[1]), presetClass: m[2], presetSubtype: Number(m[3]), nodeType: m[4] }
	effects.push({ ...key, grpId: 0, spid: Number(spid), shapeName, presetName })
	presetTemplates[presetName] = {
		key,
		effectParXml: `<p:par>${cTnFull}</p:par>`,
		behaviorsXml: childTnLst,
		bldPXml: `<p:bldP spid="${spid}" grpId="0"/>`,
	}
}

const bldLstXml = timingXml.match(/<p:bldLst>[\s\S]*<\/p:bldLst>/)[0]
const animationSpids = [...bldLstXml.matchAll(/<p:bldP spid="(\d+)"/g)].map((x) => Number(x[1]))

const oracle = {
	deck: 'slide-animation-presets',
	schema: 'slide-animation-oracle@1',
	application: 'Microsoft Office PowerPoint',
	appVersion: '16.0000',
	sha256,
	notes:
		'PowerPoint-authored preset-template oracle for docs/animations-and-transitions.md (Phase 2 capability B: expand the write-side preset set). One blank 16:9 slide, one labeled text box per preset, one on-click effect each, spanning all three preset classes: entrance fadeIn (presetID 10), flyIn (2/sub4), appear (1), wipe (22/sub4, filter "wipe(down)"); emphasis grow/shrink (6) and spin (8, animRot by="21600000" on attr "r"); exit fadeOut (10) and flyOut (2/sub4). fadeIn/flyIn/grow/fadeOut reconfirm the Phase 1 ANIM_PRESETS templates byte-for-byte; appear/wipe/spin/flyOut are the NEW templates to add. presetTemplates maps each preset name -> { key:(presetID,presetClass,presetSubtype,nodeType), effectParXml (verbatim <p:par> effect node), behaviorsXml (verbatim childTnLst inner — the parameterizable write template, by spid/dur/id), bldPXml }. NOTE on Pulse: the recommended downstream set named emphasis "pulse", but the opacity-based Pulse is NOT reachable through PowerPoint COM AddEffect (the legacy MsoAnimEffect enum exposes no opacity-pulse; probed 1..150 — none emit a style.opacity node), so the emphasis pair authored here is grow + spin and Pulse is recorded as a gap. Exit flyOut differs from entrance flyIn in two PowerPoint-authored ways captured verbatim: the ppt_x/ppt_y strVal expressions carry NO leading "#" (e.g. "ppt_x", "1+ppt_h/2") and the behavior <p:cTn> omits fill="hold".',
	shapeIds,
	effects,
	animationSpids,
	bldList: { spids: animationSpids, xml: bldLstXml },
	presetTemplates,
	timingXml,
}

await writeFile(resolve(FIX, 'slide-animation-presets.oracle.json'), JSON.stringify(oracle, null, '\t') + '\n')
console.log('sha256', sha256)
console.log('presets', Object.keys(presetTemplates).join(', '))
console.log('spids', animationSpids.join(','))
console.log('effects', effects.length)
