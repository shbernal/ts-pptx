// 3D-model (`am3d:model3d`) survival tests for `ts-pptx/read`.
//
// A 3D model reads back as an inert `p:graphicFrame` — there is no typed accessor, by design.
// What must NOT happen is a silent drop: the `am3d:model3d` subtree lives inside an
// `mc:AlternateContent` under a `graphicData@uri` nothing in the read model recognizes, and it
// depends on TWO relationships (the `.glb` payload and the preview image) whose rel types match
// none of the image/audio/video sniffs. Either could be dropped without anything else failing,
// and the deck would still open — just with no model in it.
//
// The oracle is `fixtures/model3d.pptx`, authored by desktop PowerPoint via `Shapes.Add3DModel`.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation, OpcPackage } from '../../dist/read.js'
import { bytesEqual, assert, assertEqual, partBodies } from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { FIXTURES, fixturePath } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

const MODEL3D_REL = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const AM3D_NS = 'http://schemas.microsoft.com/office/drawing/2017/model3d'

function text(bodies, name) {
	const bytes = bodies.get(name)
	assert(bytes, `part ${name} is present`)
	return new TextDecoder().decode(bytes)
}

/** The `<am3d:model3d …>…</am3d:model3d>` subtree of a slide body, or `undefined`. */
function model3dSubtree(body) {
	const start = body.indexOf('<am3d:model3d ')
	const end = body.indexOf('</am3d:model3d>')
	return start === -1 || end === -1 ? undefined : body.slice(start, end + '</am3d:model3d>'.length)
}

/** Rel `Type`s in a slide's `.rels`, keyed by Id. */
function relTypes(relsXml) {
	const out = new Map()
	for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
		const id = /\bId="([^"]+)"/.exec(m[0])?.[1]
		const type = /\bType="([^"]+)"/.exec(m[0])?.[1]
		if (id && type) out.set(id, type)
	}
	return out
}

describe('3D model: PowerPoint-authored fixture', () => {
	test('load → save leaves the model3d slide, its subtree and both media parts byte-identical', async () => {
		const input = await readFile(fixturePath('model3d'))
		const before = await partBodies(input)
		const saved = await (await OpcPackage.load(input)).save()
		const after = await partBodies(saved)

		assertEqual(
			[...after.keys()].sort().join('\n'),
			[...before.keys()].sort().join('\n'),
			'the part-name set is unchanged (neither the .glb nor the preview is dropped)'
		)
		for (const name of ['ppt/slides/slide1.xml', 'ppt/slides/_rels/slide1.xml.rels', 'ppt/media/model3d1.glb']) {
			assert(bytesEqual(before.get(name), after.get(name)), `${name} is byte-identical after the round-trip`)
		}
	})

	test('a model reads back as an inert graphicFrame — no typed accessor, but not dropped either', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('model3d')))
		const shapes = pres.slides[0].shapes
		assertEqual(shapes.length, 1, 'the model is enumerated as one shape')
		assertEqual(shapes[0].shapeType, 'graphicFrame', 'it surfaces as a graphicFrame')
		assertEqual(shapes[0].name, 'Cube3D', 'carrying its objectName')
		// Deliberate v1 scope: there is no `Model3d` shape class and no camera accessor. What this
		// pins is that the shape is *visible* to the read model, so a caller enumerating shapes sees
		// it rather than a silently shorter list.
	})

	test('the fixture pins the rel graph the emitter reproduces', async () => {
		const bodies = await partBodies(await readFile(fixturePath('model3d')))
		const body = text(bodies, 'ppt/slides/slide1.xml')
		const rels = relTypes(text(bodies, 'ppt/slides/_rels/slide1.xml.rels'))

		// The graphicFrame sits in an mc:Choice gated on the am3d namespace, with a picture fallback.
		assert(body.includes(`<mc:Choice xmlns:am3d="${AM3D_NS}" Requires="am3d">`), 'mc:Choice declares + requires am3d')
		assert(body.includes(`<a:graphicData uri="${AM3D_NS}">`), 'graphicData@uri is the am3d namespace')
		assert(body.includes('<mc:Fallback><p:pic>'), 'a picture fallback is present')

		// Two rels: the payload under the MS 2017/06 model3d type, the preview as a plain image.
		const modelRid = /<am3d:model3d r:embed="(rId\d+)"/.exec(body)?.[1]
		const previewRid = /<am3d:blip r:embed="(rId\d+)"/.exec(body)?.[1]
		assert(modelRid && previewRid, `body carries both rIds (${modelRid} / ${previewRid})`)
		assertEqual(rels.get(modelRid), MODEL3D_REL, 'the payload rel uses the MS 2017/06 model3d type')
		assertEqual(rels.get(previewRid), IMAGE_REL, 'the preview rel is an ordinary image rel')

		// The fallback picture reuses the preview rel rather than carrying a second copy.
		const fallbackRid = /<mc:Fallback>[\s\S]*?<a:blip r:embed="(rId\d+)"/.exec(body)?.[1]
		assertEqual(fallbackRid, previewRid, 'am3d:raster and the mc:Fallback picture share one image rel')
	})

	test('importSlide carries the model3d subtree and BOTH rels into another deck', async () => {
		const source = await Presentation.load(await readFile(fixturePath('model3d')))
		const targetBytes = await readFile(fixturePath('theme-colors'))
		const target = await Presentation.load(targetBytes)

		const imported = await target.importSlide(source, 0)
		assert(imported, 'the slide was imported')

		const out = await target.save()
		const after = await partBodies(out)
		const reopened = await Presentation.load(out)
		const slide = reopened.slides[reopened.slides.length - 1]
		const body = text(after, slide.partName.slice(1))

		// The subtree survives verbatim apart from rId renumbering, which is what the import does.
		const sourceBody = text(await partBodies(await readFile(fixturePath('model3d'))), 'ppt/slides/slide1.xml')
		const expected = model3dSubtree(sourceBody)
		const actual = model3dSubtree(body)
		assert(expected, 'the fixture has an am3d:model3d subtree to compare against')
		assert(actual, 'the imported slide still has an am3d:model3d subtree')
		assertEqual(
			actual.replace(/rId\d+/g, 'rId#'),
			expected.replace(/rId\d+/g, 'rId#'),
			'the am3d:model3d subtree survives the import unchanged (modulo rId renumbering)'
		)

		// Both rels came along, and both resolve to real parts.
		const relsXml = text(after, `ppt/slides/_rels/${slide.partName.split('/').pop()}.rels`)
		const rels = relTypes(relsXml)
		const modelRid = /<am3d:model3d r:embed="(rId\d+)"/.exec(body)[1]
		const previewRid = /<am3d:blip r:embed="(rId\d+)"/.exec(body)[1]
		assertEqual(rels.get(modelRid), MODEL3D_REL, 'the model3d rel survives the import')
		assertEqual(rels.get(previewRid), IMAGE_REL, 'the preview image rel survives the import')
		for (const rid of [modelRid, previewRid]) {
			const target_ = /<Relationship\b[^>]*\bId="(?:rId\d+)"[^>]*\/>/g
			const match = [...relsXml.matchAll(target_)].find((m) => m[0].includes(`Id="${rid}"`))
			const rel = /\bTarget="([^"]+)"/.exec(match[0])[1].replace('../', 'ppt/')
			assert(after.has(rel), `${rid} resolves to a real part (${rel})`)
		}

		// The .glb part came across with its bytes intact and its content type registered.
		const glbName = [...after.keys()].find((n) => n.endsWith('.glb'))
		assert(glbName, 'a .glb part was copied into the target')
		const originalGlb = (await partBodies(await readFile(fixturePath('model3d')))).get('ppt/media/model3d1.glb')
		assert(bytesEqual(originalGlb, after.get(glbName)), 'the .glb payload bytes are unchanged by the import')

		// The copied part must carry a content type, or PowerPoint reports the package as corrupt.
		// `OpcPackage.addPart` registers an `Override` rather than a `Default` — that is the read
		// path's general behaviour for every copied media part (an imported `.jpg` gets one too),
		// not something specific to `.glb`. Both spellings are legal OPC; assert the invariant that
		// matters rather than the mechanism, so this test does not pin unrelated import behaviour.
		const ct = text(after, '[Content_Types].xml')
		assert(
			ct.includes('<Default Extension="glb" ContentType="model/gltf.binary"/>') ||
				ct.includes(`<Override PartName="/${glbName}" ContentType="model/gltf.binary"/>`),
			`the copied .glb has a registered content type (${glbName})`
		)
	})
})

describe('3D model: ts-pptx-authored', () => {
	async function authored(options = {}) {
		const pptx = new TsPptx()
		pptx.layout = 'LAYOUT_WIDE'
		pptx.addSlide().addModel3d({
			path: path.join(FIXTURES, 'authoring', 'assets', 'cube.glb'),
			preview: { path: path.join(FIXTURES, 'media', 'poster.png') },
			x: 1,
			y: 1,
			w: 4,
			h: 3,
			...options,
		})
		// `write` is typed for every output target; `nodebuffer` resolves to a Buffer here.
		return partBodies(/** @type {Buffer} */ (await pptx.write({ outputType: 'nodebuffer' })))
	}

	test('emits the rel graph and content type the PowerPoint fixture pins', async () => {
		const bodies = await authored()
		const body = text(bodies, 'ppt/slides/slide1.xml')
		const rels = relTypes(text(bodies, 'ppt/slides/_rels/slide1.xml.rels'))

		const modelRid = /<am3d:model3d r:embed="(rId\d+)"/.exec(body)[1]
		const previewRid = /<am3d:blip r:embed="(rId\d+)"/.exec(body)[1]
		assertEqual(rels.get(modelRid), MODEL3D_REL, 'the payload rel uses the MS 2017/06 model3d type')
		assertEqual(rels.get(previewRid), IMAGE_REL, 'the preview rel is an ordinary image rel')
		assertEqual(
			/<mc:Fallback>[\s\S]*?<a:blip r:embed="(rId\d+)"/.exec(body)[1],
			previewRid,
			'the fallback picture reuses the preview rel'
		)
		assert(
			text(bodies, '[Content_Types].xml').includes('<Default Extension="glb" ContentType="model/gltf.binary"/>'),
			'glb is registered as a Default, matching PowerPoint'
		)
		assert(
			[...bodies.keys()].some((n) => /^ppt\/media\/model3d-1-1\.glb$/.test(n)),
			'the payload lands in ppt/media/ under a model3d name'
		)
	})

	test('emits the same am3d:model3d body PowerPoint does, modulo rIds and geometry', async () => {
		const authoredBody = text(await authored(), 'ppt/slides/slide1.xml')
		const fixtureBody = text(await partBodies(await readFile(fixturePath('model3d'))), 'ppt/slides/slide1.xml')

		// Normalize away the two things that legitimately differ: rIds, and the frame extent,
		// which appears in `am3d:spPr`'s frame-local xfrm.
		const normalize = (xml) =>
			model3dSubtree(xml)
				.replace(/rId\d+/g, 'rId#')
				.replace(/<a:ext cx="\d+" cy="\d+"\/>/, '<a:ext/>')
		assertEqual(
			normalize(authoredBody),
			normalize(fixtureBody),
			'camera, transform, raster, viewport and the whole lighting rig match PowerPoint byte-for-byte'
		)
	})

	test('a caller-supplied camera and scale reach the XML in am3d wire units', async () => {
		const body = text(
			await authored({ camera: { pos: { x: 1, y: 2, z: 3 }, fov: 60 }, meterPerModelUnit: 0.01 }),
			'ppt/slides/slide1.xml'
		)
		assert(body.includes('<am3d:pos x="36000000" y="72000000" z="108000000"/>'), 'camera.pos is scaled by 36,000,000')
		assert(body.includes('<am3d:perspective fov="3600000"/>'), 'fov is emitted in 60000ths of a degree')
		assert(body.includes('<am3d:meterPerModelUnit n="10000" d="1000000"/>'), 'meterPerModelUnit is emitted as n/1e6')
	})

	// READ THIS BEFORE TRUSTING THE NEXT TEST. The Open XML SDK validator does **not** check
	// anything inside an `mc:Choice` — it validates only the `mc:Fallback` branch. Measured at
	// `Microsoft365` on this very deck: a bogus attribute on the text shape outside the
	// `mc:AlternateContent` reports 1 error, and one inside `mc:Fallback` reports 1 error, but the
	// same mutation on `p:xfrm` or `am3d:model3d` inside the `mc:Choice` reports 0 — as do an
	// unknown `am3d` child, `am3d:camera` moved out of document order, a non-numeric
	// `perspective@fov`, and `am3d:model3d` with its required `r:embed` deleted. The SDK does model
	// `am3d` (Office2019+), but it never reaches it here. This is a property of
	// `mc:AlternateContent`, not of `am3d`, so the zoom and OLE emitters share the blind spot.
	//
	// So this test covers the FALLBACK picture, the rel graph and the content types — real
	// coverage, and the parts a strict consumer reads — but it is not evidence about the `am3d`
	// body. What covers that is the byte-for-byte comparison against PowerPoint's own output
	// above, plus the desktop check in `pnpm run test:com`.
	test.skipIf(!validatorInstalled)('the package and its mc:Fallback branch are schema-valid', async () => {
		const pptx = new TsPptx()
		pptx.layout = 'LAYOUT_WIDE'
		pptx.addSlide().addModel3d({
			path: path.join(FIXTURES, 'authoring', 'assets', 'cube.glb'),
			preview: { path: path.join(FIXTURES, 'media', 'poster.png') },
			x: 1,
			y: 1,
			w: 4,
			h: 3,
		})
		const errors = await validateBuf(/** @type {Buffer} */ (await pptx.write({ outputType: 'nodebuffer' })))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
