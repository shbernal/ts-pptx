// Probe; produces no committed fixture. Measures how a SmartArt data model's authored
// points reach the drawing cache PowerPoint bakes alongside them, which is what a re-texter
// has to mirror: `ppt/diagrams/data{N}.xml` holds the nodes, `ppt/diagrams/drawing{N}.xml`
// holds a copy of every drawn string, and every renderer that is not PowerPoint draws the
// copy.
//
// The rule under test, per authored point:
//
//   dgm:pt/@modelId = N
//     -> dgm:cxn[@type="presOf"][@srcId=N] -> @destId = P, @destOrd = k
//          -> dsp:sp/@modelId = P
//               -> the point's text is paragraph k of that sp's dsp:txBody
//
//   node test/read/fixtures/authoring/probe-smartart-mapping.mjs [deck.pptx ...]
//
// With no arguments it probes ../smartart-families.pptx and ../mixed.pptx. Exits non-zero
// if any authored point with text fails to resolve, so it doubles as a regression check on
// a re-authored fixture.

import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import JSZip from 'jszip'

// This script lives in test/read/fixtures/authoring/, so the fixtures dir is its parent.
const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..')

const DRAWING_EXT_URI = 'http://schemas.microsoft.com/office/drawing/2008/diagram'

const decks = process.argv.slice(2)
if (decks.length === 0) decks.push(resolve(FIX, 'smartart-families.pptx'), resolve(FIX, 'mixed.pptx'))

/** `<a:t>` payloads of one element, in document order — good enough to identify a string. */
function textOf(xml) {
	return [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decode(m[1])).join('')
}

/** One entry per `<a:p>`, so a body's paragraphs can be indexed by `destOrd`. */
function paragraphsOf(xml) {
	return [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>|<a:p\/>/g)].map((m) => textOf(m[1] ?? ''))
}

function decode(s) {
	return s
		.replaceAll('&lt;', '<')
		.replaceAll('&gt;', '>')
		.replaceAll('&quot;', '"')
		.replaceAll('&apos;', "'")
		.replaceAll('&amp;', '&')
}

/** Every `<tag …>…</tag>` (or self-closed) at any depth, as raw substrings. */
function elements(xml, tag) {
	const out = []
	const open = new RegExp(`<${tag}(\\s[^>]*?)?(/)?>`, 'g')
	for (let m = open.exec(xml); m; m = open.exec(xml)) {
		if (m[2]) {
			out.push(m[0])
			continue
		}
		// Nested same-tag elements do not occur for any tag this probe reads (dgm:pt, dgm:cxn,
		// dsp:sp), so a plain scan to the matching close tag is exact.
		const end = xml.indexOf(`</${tag}>`, open.lastIndex)
		out.push(xml.slice(m.index, end + tag.length + 3))
		open.lastIndex = end
	}
	return out
}

function attrOf(xml, name) {
	const m = xml.match(new RegExp(`\\s${name}="([^"]*)"`))
	return m ? decode(m[1]) : null
}

async function relationships(zip, partName) {
	const dir = partName.slice(0, partName.lastIndexOf('/'))
	const base = partName.slice(partName.lastIndexOf('/') + 1)
	const file = zip.file(`${dir}/_rels/${base}.rels`)
	if (!file) return {}
	const xml = await file.async('string')
	const map = {}
	for (const rel of elements(xml, 'Relationship')) {
		const target = attrOf(rel, 'Target')
		map[attrOf(rel, 'Id')] = target.startsWith('/') ? target.slice(1) : normalize(`${dir}/${target}`)
	}
	return map
}

function normalize(path) {
	const out = []
	for (const seg of path.split('/')) {
		if (seg === '.' || seg === '') continue
		if (seg === '..') out.pop()
		else out.push(seg)
	}
	return out.join('/')
}

/** Every diagram in the deck, as `{ slide, dataPart, drawingPart }`. */
async function diagrams(zip) {
	const found = []
	const slideNames = Object.keys(zip.files)
		.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
		.sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))
	for (const slideName of slideNames) {
		const xml = await zip.file(slideName).async('string')
		const rels = await relationships(zip, slideName)
		for (const relIds of elements(xml, 'dgm:relIds')) {
			const dataPart = rels[attrOf(relIds, 'r:dm')]
			const dataXml = await zip.file(dataPart).async('string')
			found.push({
				slide: Number(slideName.match(/\d+/)[0]),
				dataPart,
				dataXml,
				drawingPart: drawingPartOf(dataXml, rels),
			})
		}
	}
	return found
}

/**
 * The drawing part is named by an MS extension on the data model (`dgm:extLst/a:ext/@uri` =
 * the 2008 diagram uri, then `dsp:dataModelExt/@relId`) rather than by `dgm:relIds`, so a
 * slide with two diagrams still resolves each to its own cache. The relationship id belongs
 * to the part holding the *frame*, not to the data part — a data part has no `_rels` at all.
 * Same route `Diagram.drawingPart` takes in `src/read/api/diagram.ts`.
 */
function drawingPartOf(dataXml, frameRels) {
	for (const ext of elements(dataXml, 'a:ext')) {
		if (attrOf(ext, 'uri') !== DRAWING_EXT_URI) continue
		const relId = attrOf(ext, 'relId')
		if (relId) return frameRels[relId] ?? null
	}
	return null
}

const failures = []
let probed = 0

for (const deck of decks) {
	const zip = await JSZip.loadAsync(await readFile(deck))
	console.log(`\n${'='.repeat(78)}\n${deck}`)

	for (const diagram of await diagrams(zip)) {
		const points = elements(diagram.dataXml, 'dgm:pt').map((xml) => ({
			modelId: attrOf(xml, 'modelId'),
			type: attrOf(xml, 'type') ?? 'node',
			text: textOf(xml),
		}))
		const layoutId = (() => {
			const doc = elements(diagram.dataXml, 'dgm:pt').find((xml) => attrOf(xml, 'type') === 'doc')
			return doc ? attrOf(doc, 'loTypeId') : null
		})()

		// `presName` is the layout engine's own label for what a pres point draws (`rootText`,
		// `rootConnector`, `sibTrans`, …). It is the readable half of the mapping: it says which
		// of a node's several presentations is the one with the words in it.
		const presNames = new Map()
		for (const xml of elements(diagram.dataXml, 'dgm:pt')) {
			if (attrOf(xml, 'type') !== 'pres') continue
			presNames.set(attrOf(xml, 'modelId'), attrOf(xml, 'presName'))
		}

		const presOf = new Map()
		for (const cxn of elements(diagram.dataXml, 'dgm:cxn')) {
			if (attrOf(cxn, 'type') !== 'presOf') continue
			const src = attrOf(cxn, 'srcId')
			if (!presOf.has(src)) presOf.set(src, [])
			presOf.get(src).push({
				destId: attrOf(cxn, 'destId'),
				destOrd: Number(attrOf(cxn, 'destOrd') ?? 0),
				sourceOrd: Number(attrOf(cxn, 'srcOrd') ?? 0),
			})
		}
		for (const list of presOf.values()) list.sort((a, b) => a.sourceOrd - b.sourceOrd)

		// A `dsp:sp` with no `dsp:txBody` at all is a drawn shape that cannot hold text —
		// a connector, or a picture frame. Distinct from one whose body is empty.
		const shapes = new Map()
		if (diagram.drawingPart) {
			const drawingXml = await zip.file(diagram.drawingPart).async('string')
			for (const sp of elements(drawingXml, 'dsp:sp')) {
				const body = sp.match(/<dsp:txBody>[\s\S]*?<\/dsp:txBody>/)
				shapes.set(attrOf(sp, 'modelId'), body ? paragraphsOf(body[0]) : null)
			}
		}

		const counts = {}
		for (const p of points) counts[p.type] = (counts[p.type] ?? 0) + 1
		console.log(`\n--- slide ${diagram.slide}  ${diagram.dataPart} -> ${diagram.drawingPart ?? '(no drawing part)'}`)
		console.log(`    layout ${layoutId}`)
		console.log(`    points ${points.length} ${JSON.stringify(counts)}   presOf ${presOf.size}   dsp:sp ${shapes.size}`)

		for (const point of points) {
			if (point.type === 'pres' || point.type === 'doc') continue
			probed++
			const cxns = presOf.get(point.modelId) ?? []
			const label = `${point.type.padEnd(8)} ${point.modelId.slice(0, 9)}… ${JSON.stringify(point.text).padEnd(10)}`

			// Every presentation of the point, in `srcOrd` order — the org-chart box and the
			// connector under it are two presentations of one node, and only the first has text.
			const arms = cxns.map(({ destId, destOrd, sourceOrd }) => {
				const paragraphs = shapes.get(destId)
				return {
					destId,
					destOrd,
					sourceOrd,
					presName: presNames.get(destId),
					drawn: paragraphs === undefined ? undefined : (paragraphs?.[destOrd] ?? null),
					// undefined: the pres point draws nothing. null: it draws a shape with no text body.
					kind: paragraphs === undefined ? 'no sp' : paragraphs === null ? 'sp, no txBody' : 'sp',
				}
			})

			const rendered = arms
				.map(
					(a) =>
						`[srcOrd ${a.sourceOrd}] ${a.presName ?? '?'} ${a.kind}` +
						(a.kind === 'sp' ? ` para[${a.destOrd}]=${JSON.stringify(a.drawn)}` : '')
				)
				.join('  |  ')
			console.log(`    ${label} -> ${rendered || 'no presOf'}`)

			if (!point.text) continue
			const hits = arms.filter((a) => a.drawn === point.text)
			if (hits.length === 1) continue
			failures.push({
				deck,
				slide: diagram.slide,
				layoutId,
				point,
				reason:
					hits.length === 0 ? `text drawn by none of ${arms.length} presentations` : `text drawn by ${hits.length}`,
			})
		}
	}
}

console.log(`\n${'='.repeat(78)}`)
console.log(
	`probed ${probed} authored points; ${failures.length} with text did not resolve to exactly one drawn paragraph`
)
for (const f of failures) {
	console.log(
		`  slide ${f.slide} ${f.layoutId?.replace(/.*\//, '')}  ${f.point.type} ${JSON.stringify(f.point.text)}  ${f.reason}`
	)
}
process.exitCode = failures.length === 0 ? 0 : 1
