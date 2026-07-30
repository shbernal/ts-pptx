import { readFile, writeFile } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import JSZip from 'jszip'

// This script lives in test/read/fixtures/authoring/, so the fixtures dir is its parent.
const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..')

const bytes = await readFile(resolve(FIX, 'import-animation-merge.pptx'))
const sha256 = createHash('sha256').update(bytes).digest('hex')
const zip = await JSZip.loadAsync(bytes)

function elementAt(s, startIdx, tag) {
	let i = startIdx
	let depth = 0
	while (i < s.length) {
		const nextOpen = s.indexOf(`<${tag}`, i)
		const nextClose = s.indexOf(`</${tag}>`, i)
		if (nextOpen !== -1 && (nextClose === -1 || nextOpen < nextClose)) {
			const gt = s.indexOf('>', nextOpen)
			const selfClose = s[gt - 1] === '/'
			if (!selfClose) depth++
			i = gt + 1
			if (selfClose && depth === 0) return s.slice(startIdx, gt + 1)
		} else if (nextClose !== -1) {
			depth--
			i = nextClose + `</${tag}>`.length
			if (depth === 0) return s.slice(startIdx, i)
		} else break
	}
	throw new Error('unbalanced ' + tag)
}

async function slideModel(n) {
	const xml = await zip.file(`ppt/slides/slide${n}.xml`).async('string')
	const shapeIds = {}
	for (const m of xml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)) shapeIds[m[1]] = m[2]
	const timingXml = xml.match(/<p:timing>[\s\S]*<\/p:timing>/)[0]
	const bldLstXml = timingXml.match(/<p:bldLst>[\s\S]*<\/p:bldLst>/)[0]
	const animationSpids = [...bldLstXml.matchAll(/<p:bldP spid="(\d+)"/g)].map((x) => Number(x[1]))
	const effects = []
	const re =
		/<p:cTn id="\d+" presetID="(\d+)" presetClass="(entr|emph|exit)" presetSubtype="(\d+)" fill="hold" grpId="0" nodeType="(\w+)">/g
	for (const m of timingXml.matchAll(re)) {
		const cTnFull = elementAt(timingXml, m.index, 'p:cTn')
		const spid = Number(cTnFull.match(/<p:spTgt spid="(\d+)"/)[1])
		effects.push({
			presetID: Number(m[1]),
			presetClass: m[2],
			presetSubtype: Number(m[3]),
			nodeType: m[4],
			grpId: 0,
			spid,
			shapeName: shapeIds[String(spid)],
		})
	}
	return { shapeIds, effects, animationSpids, bldList: { spids: animationSpids, xml: bldLstXml }, timingXml }
}

const source = { slide: 1, ...(await slideModel(1)) }
const merged = { slide: 2, ...(await slideModel(2)) }

const oracle = {
	deck: 'import-animation-merge',
	schema: 'slide-animation-oracle@1',
	application: 'Microsoft Office PowerPoint',
	appVersion: '16.0000',
	sha256,
	notes:
		"PowerPoint-authored cross-slide animation-merge oracle for docs/animations-and-transitions.md (Phase 2 capability A: carry a build animation through importShape, remapping into the destination p:timing; backlog dn-1863). Two blank 16:9 slides. Slide 1 \"Source\" (spid 2) has an entrance Fade-on-click (presetID 10), mirroring slide-animation-basic. Slide 2 \"HostExisting\" (spid 2) has its own entrance Fly-on-click (presetID 2 subtype 4); then in PowerPoint \"Source\" was copied from slide 1 and pasted onto slide 2 WITH its animation. PowerPoint's observed merge behavior (the ground truth importShape's remap+merge must reproduce): (1) the pasted shape takes the next free shape id on the destination — Source 2 -> 3, HostExisting keeps 2; (2) the carried build's every <p:spTgt spid> and <p:bldP spid> is renumbered from the source spid (2) to the new destination spid (3); (3) the carried build is APPENDED as a new click group after the host's existing build(s), and its <p:bldP> appended after the host's in <p:bldLst>; (4) every <p:cTn id> in the merged tree is renumbered sequentially in document order (the host's fly effect keeps ids 3-8, the carried fade effect becomes ids 9-13). The `source` block pins slide 1, `merged` pins slide 2's post-merge tree, and `mergeMap` records the spid remap + build ordering. Assert no dangling spTgt/bldP on either side after a programmatic remap.",
	source,
	merged,
	mergeMap: {
		hostShape: { name: 'HostExisting', spid: 2, presetID: 2, presetClass: 'entr' },
		carriedShape: { name: 'Source', sourceSpid: 2, mergedSpid: 3, presetID: 10, presetClass: 'entr' },
		spidRemap: { 2: 3 },
		buildOrder: ['HostExisting (spid 2, fly)', 'Source (spid 3, fade — carried, appended last)'],
		bldListOrder: [2, 3],
	},
}

await writeFile(resolve(FIX, 'import-animation-merge.oracle.json'), JSON.stringify(oracle, null, '\t') + '\n')
console.log('sha256', sha256)
console.log('source spids', source.animationSpids.join(','), '| merged spids', merged.animationSpids.join(','))
console.log('source effects', JSON.stringify(source.effects.map((e) => `${e.shapeName}:${e.presetID}@${e.spid}`)))
console.log('merged effects', JSON.stringify(merged.effects.map((e) => `${e.shapeName}:${e.presetID}@${e.spid}`)))
