import { readFileSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import { unzipSync, strFromU8 } from 'fflate'

// This script lives in test/read/fixtures/authoring/, so the fixtures dir is its parent.
const HERE = dirname(fileURLToPath(import.meta.url))
const FIX = resolve(HERE, '..')

function load(path) {
	const bytes = readFileSync(path)
	const zip = unzipSync(bytes)
	const txt = (p) => strFromU8(zip[p])
	const sha256 = createHash('sha256').update(bytes).digest('hex')
	const app = txt('docProps/app.xml')
	const application = (app.match(/<Application>([^<]*)<\/Application>/) || [])[1] ?? null
	const appVersion = (app.match(/<AppVersion>([^<]*)<\/AppVersion>/) || [])[1] ?? null
	return { zip, txt, sha256, application, appVersion }
}

// Ordered slide parts via presentation.xml sldIdLst -> rels.
function slideOrder(txt) {
	const pres = txt('ppt/presentation.xml')
	const rels = txt('ppt/_rels/presentation.xml.rels')
	const relMap = {}
	for (const m of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?>/g)) relMap[m[1]] = m[2]
	const order = []
	for (const m of pres.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/)[1].matchAll(/r:id="([^"]+)"/g)) {
		let t = relMap[m[1]]
		if (!t.startsWith('ppt/')) t = 'ppt/' + t.replace(/^\/?/, '')
		order.push(t)
	}
	return order
}

function decodeTransition(tx) {
	const open = tx.match(/<p:transition\b([^>]*?)\/?>/)
	const attrs = open ? open[1] : ''
	const at = (k) => (attrs.match(new RegExp('\\b' + k + '="([^"]+)"')) || [])[1] ?? null
	const body = tx.replace(/^<p:transition\b[^>]*?>/, '').replace(/<\/p:transition>\s*$/, '')
	const child = body.match(/<(?:([\w]+):)?([\w]+)\b([^>]*?)\/?>/)
	let element = null,
		ns = 'p'
	const variant = {}
	if (child) {
		ns = child[1] || 'p'
		element = child[2]
		for (const a of (child[3] || '').matchAll(/([\w:]+)="([^"]*)"/g)) variant[a[1]] = a[2]
	}
	return {
		element,
		ns,
		variant,
		spd: at('spd'),
		p14dur: at('p14:dur'),
		advClick: at('advClick'),
		advTm: at('advTm'),
	}
}

function extractTransitionRegion(xml) {
	let m = xml.match(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/)
	if (m) return { xml: m[0], wrapped: true }
	m = xml.match(/<p:transition\b[^>]*\/>/) || xml.match(/<p:transition\b[\s\S]*?<\/p:transition>/)
	if (m) return { xml: m[0], wrapped: false }
	return null
}

function firstTransition(s) {
	const m = s.match(/<p:transition\b[\s\S]*?<\/p:transition>|<p:transition\b[^>]*\/>/)
	return m ? m[0] : null
}

// ---------- slide-transition oracle ----------
function buildTransition() {
	const f = load(`${FIX}/slide-transition.pptx`)
	const order = slideOrder(f.txt)
	const slides = order.map((part, i) => {
		const xml = f.txt(part)
		const region = extractTransitionRegion(xml)
		const out = {
			slide: i + 1,
			part: part.replace('ppt/slides/', ''),
			wrapped: region?.wrapped ?? false,
			transitionXml: region?.xml ?? null,
		}
		if (!region) return out
		if (region.wrapped) {
			const ch = region.xml.match(/<mc:Choice\b[^>]*Requires="([^"]+)"[\s\S]*?>([\s\S]*?)<\/mc:Choice>/)
			const fb = region.xml.match(/<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/)
			const choiceTx = ch ? firstTransition(ch[2]) : null
			const fbTx = fb ? firstTransition(fb[1]) : null
			out.requires = ch ? ch[1] : null
			out.choice = choiceTx ? decodeTransition(choiceTx) : null
			out.fallback = fbTx ? decodeTransition(fbTx) : null
		} else {
			out.fallback = decodeTransition(region.xml)
			out.choice = null
		}
		const d = out.choice ?? out.fallback
		out.decoded = {
			type: d.element,
			ns: d.ns,
			variant: d.variant,
			// spd defaults to "fast" when the attribute is absent (ECMA-376 ST_TransitionSpeed).
			speed: out.fallback?.spd ?? d.spd ?? 'fast',
			speedAttrPresent: (out.fallback?.spd ?? d.spd) != null,
			durationMs: d.p14dur != null ? Number(d.p14dur) : null,
			// advClick defaults to true when absent.
			advanceOnClick: (out.fallback?.advClick ?? d.advClick) !== '0',
			advanceAfterMs: (out.fallback?.advTm ?? d.advTm) != null ? Number(out.fallback?.advTm ?? d.advTm) : null,
		}
		return out
	})

	// Full probed PpEntryEffect -> element table (write-side preset table).
	const probe = JSON.parse(readFileSync(resolve(HERE, 'entryeffect-table.json'), 'utf8'))
	const entryEffectTable = probe
		.filter((r) => r.element != null)
		.map((r) => ({
			entryEffect: r.entryEffect,
			element: r.element,
			ns: r.ns ?? 'p',
			variant: r.variant ?? {},
			// true when the effect has no base ECMA-376 p:transition child of its own kind:
			// PowerPoint expresses it only inside mc:Choice (p14/p15/p159) with a <p:fade/> fallback.
			modernOnly: (r.ns ?? 'p') !== 'p',
		}))
		.sort((a, b) => a.entryEffect - b.entryEffect)

	const oracle = {
		deck: 'slide-transition',
		schema: 'slide-transition-oracle@1',
		application: f.application,
		appVersion: f.appVersion,
		sha256: f.sha256,
		notes:
			'PowerPoint-authored slide-transition oracle for docs/animations-and-transitions.md (Phase 1 transition read+write). ' +
			'Six slides, each with one distinct p:transition authored via SlideShowTransition: fade (fast bucket, spd attr absent = default fast), ' +
			'push dir="d" (exact 1.25s -> mc:AlternateContent p14:dur="1250"), wipe dir="u" (med bucket), cut (fast bucket), ' +
			'dissolve (exact 2.0s -> p14:dur="2000"), and fade (med bucket, advClick="0" + advTm="3000" timed auto-advance). ' +
			'Position in CT_Slide: between p:clrMapOvr and p:timing. PowerPoint emits a bare <p:transition> when the duration matches a ' +
			'speed bucket and the mc:AlternateContent form (mc:Choice Requires="p14" carrying p14:dur, mc:Fallback without) only for an ' +
			'off-bucket exact duration. spd is omitted entirely for the default "fast" bucket. ' +
			'entryEffectTable is the full 158-row probed PpEntryEffect(int) -> {element, ns, variant attrs, modernOnly} mapping (the int 0 = no ' +
			'transition is excluded), captured by iterating every accepted SlideShowTransition.EntryEffect value 0..4096 (159 valid incl. 0); ' +
			'it is the write-side preset table (the family<<8 packing intuition is wrong, e.g. 1537 -> dissolve). ns "p" is base ECMA-376; ' +
			'"p14"/"p15"/"p159" elements appear only inside mc:Choice with a base fade fallback.',
		slides,
		entryEffectTable,
	}
	writeFileSync(`${FIX}/slide-transition.oracle.json`, JSON.stringify(oracle, null, '\t') + '\n')
	console.log(
		'slide-transition.oracle.json',
		'— slides:',
		slides.length,
		'sha:',
		f.sha256.slice(0, 12),
		'table rows:',
		entryEffectTable.length
	)
	return oracle
}

// ---------- animation oracle (basic + rich) ----------
function shapeNames(xml) {
	const map = {}
	for (const m of xml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)) map[m[1]] = m[2]
	return map
}
function effects(timing) {
	const rows = []
	const re = /<p:cTn id="\d+"([^>]*presetClass="[^"]*"[^>]*)>([\s\S]*?)<p:spTgt spid="(\d+)"/g
	let m
	while ((m = re.exec(timing))) {
		const a = m[1]
		const g = (k) => (a.match(new RegExp(k + '="([^"]*)"')) || [])[1]
		rows.push({
			presetID: Number(g('presetID')),
			presetClass: g('presetClass'),
			presetSubtype: Number(g('presetSubtype')),
			nodeType: g('nodeType'),
			grpId: Number(g('grpId')),
			spid: Number(m[3]),
		})
	}
	return rows
}
function buildAnimation(deck, noteHead) {
	const f = load(`${FIX}/${deck}.pptx`)
	const xml = f.txt('ppt/slides/slide1.xml')
	const names = shapeNames(xml)
	const timing = (xml.match(/<p:timing>[\s\S]*<\/p:timing>/) || [null])[0]
	const bldLst = (xml.match(/<p:bldLst>[\s\S]*?<\/p:bldLst>/) || [null])[0]
	const eff = effects(timing).map((e) => ({ ...e, shapeName: names[String(e.spid)] ?? null }))
	const bldSpids = [...(bldLst ?? '').matchAll(/<p:bldP\b[^>]*spid="(\d+)"/g)].map((m) => Number(m[1]))
	const spidsAll = [...timing.matchAll(/<p:spTgt spid="(\d+)"/g)].map((m) => Number(m[1]))
	const animationSpids = [...new Set(spidsAll)].sort((a, b) => a - b)

	const oracle = {
		deck,
		schema: 'slide-animation-oracle@1',
		application: f.application,
		appVersion: f.appVersion,
		sha256: f.sha256,
		notes: noteHead,
		shapeIds: names,
		effects: eff,
		animationSpids,
		bldList: { spids: bldSpids, xml: bldLst },
		timingXml: timing,
	}
	writeFileSync(`${FIX}/${deck}.oracle.json`, JSON.stringify(oracle, null, '\t') + '\n')
	console.log(
		`${deck}.oracle.json`,
		'— effects:',
		eff.length,
		'spids:',
		animationSpids.join(','),
		'sha:',
		f.sha256.slice(0, 12)
	)
	return oracle
}

buildTransition()
buildAnimation(
	'slide-animation-basic',
	'PowerPoint-authored basic-animation oracle for docs/animations-and-transitions.md (Phase 1 animation opaque-preserve + spid enumerate). ' +
		'One blank 16:9 slide whose text box "fade-target" (shape id/spid 2) carries a single entrance effect: Fade on click ' +
		'(MsoAnimEffect=10). The p:timing tree is CT_TimeNodeList: tnLst > par > cTn(tmRoot) > seq(mainSeq) with one nested par/cTn ' +
		'carrying presetID="10" presetClass="entr" presetSubtype="0" nodeType="clickEffect", a p:set of style.visibility + p:animEffect ' +
		'filter="fade", both targeting p:spTgt spid="2"; the sibling p:bldLst holds <p:bldP spid="2" grpId="0"/>. The (presetID, ' +
		'presetClass, presetSubtype) triple and the referenced spid are the structured data the read model extracts; the tree itself is ' +
		'preserved opaquely.'
)
buildAnimation(
	'slide-animation-rich',
	'PowerPoint-authored rich-animation oracle for docs/animations-and-transitions.md (Phase 1 spid enumerate/remap/prune + write-side ' +
		'preset templates). One blank 16:9 slide with four text boxes, each with one effect spanning all three preset classes and all ' +
		'three triggers: "ent-fade-click" (spid 2) entrance Fade presetID=10 nodeType="clickEffect"; "ent-fly-after" (spid 3) entrance ' +
		'Fly presetID=2 subtype=4 nodeType="afterEffect"; "emph-grow-with" (spid 4) emphasis Grow/Shrink presetID=6 nodeType="withEffect" ' +
		'(MsoAnimEffect=59 — the entrance-block value 36 is NOT an emphasis); "exit-fade-click" (spid 5) exit Fade presetID=10 ' +
		'nodeType="clickEffect" (authored via AddEffect then Effect.Exit=msoTrue). p:bldLst has one <p:bldP spid grpId> per shape ' +
		'(spids 2,3,4,5). This is the source deck for the write-side preset templates keyed by (presetID, presetClass, presetSubtype, ' +
		'nodeType) and the enumerate/remap/prune spid-coherence tests.'
)
