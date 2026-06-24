// Append-onto-existing tests for `pptxgenjs/read` (dn-append-onto-existing-deck).
//
// Contract under test: Presentation.appendSlides(source, { layout }) authors
// slides on a generator (PptxGenJS), serializes them via source.extractSlides(),
// and splices them into a loaded deck bound to an existing layout — keeping the
// deck's masters/layouts/theme (and every other untouched part) byte-identical,
// changing only presentation.xml, its .rels, [Content_Types].xml, and the new
// slide/media parts. Survives a save → reopen round-trip, resolves its layout to
// the *existing* layout (no new chrome), and stays schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'

// 1×1 transparent PNG.
const PNG_1PX =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

function bytesEqual(a, b) {
	return a && b && a.length === b.length && a.every((value, index) => value === b[index])
}

async function rejects(fn) {
	try {
		await fn()
		return false
	} catch {
		return true
	}
}

function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const matches = [...rels].filter((rel) => rel.type === type)
	if (matches.length === 0) return null
	return rels.resolveTarget(matches[0].id)
}

/** Resolve the absolute target a single `rId` points at, or `null`. */
function resolveRid(opc, partName, rId) {
	const rels = opc.relationshipsFor(partName)
	const match = [...rels].find((rel) => rel.id === rId)
	return match ? rels.resolveTarget(match.id) : null
}

/** The rel `type` of a single `rId` on a part, or `null`. */
function typeOfRid(opc, partName, rId) {
	const rels = opc.relationshipsFor(partName)
	const match = [...rels].find((rel) => rel.id === rId)
	return match ? match.type : null
}

async function mediaFixture(name) {
	const buf = await readFile(path.join(__dirname, 'fixtures', 'media', name))
	return buf.toString('base64')
}

/** A generator deck sized to LAYOUT_WIDE (12192000×6858000), matching theme-colors / image. */
function wideGenerator() {
	const pptx = new PptxGenJS()
	pptx.layout = 'LAYOUT_WIDE'
	return pptx
}

describe('Presentation.appendSlides', () => {
	test('appends a generated slide bound to an existing layout, keeping chrome byte-identical', async () => {
		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)
		const beforeSlideCount = pres.slides.length

		const target = pres.layouts().find((l) => l.name === 'Blank')
		assert(target, 'theme-colors has a "Blank" layout to bind to')

		const pptx = wideGenerator()
		const slide = pptx.addSlide()
		slide.addText('hello append', { x: 1, y: 1, w: 6, h: 1, color: 'FF0000' })
		slide.addImage({ data: PNG_1PX, x: 1, y: 3, w: 1, h: 1 })

		const added = await pres.appendSlides(pptx, { layout: 'Blank' })
		assertEqual(added.length, 1, 'one slide was appended')

		const out = await pres.save()
		const after = await partBodies(out)

		// Byte-stability: every original part is byte-identical except the three the
		// append legitimately touches.
		const expectedChanged = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, bytes] of before) {
			assert(after.has(name), `part ${name} survives the append`)
			if (expectedChanged.has(name)) continue
			assert(bytesEqual(bytes, after.get(name)), `part ${name} is byte-identical after the append`)
		}

		// The only new parts are the slide, its .rels, and its media.
		const newParts = [...after.keys()].filter((name) => !before.has(name))
		for (const name of newParts) {
			assert(/^ppt\/(slides|media)\//.test(name), `new part ${name} is a slide or media part`)
		}
		assert(
			newParts.some((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)),
			'a new slide part was added'
		)
		assert(
			newParts.some((n) => /^ppt\/media\/image\d+\.png$/.test(n)),
			'a new media part was added'
		)

		// Round-trip: the new slide is present, last, and carries the authored text.
		const reopened = await Presentation.load(out)
		assertEqual(reopened.slides.length, beforeSlideCount + 1, 'the deck gained exactly one slide')
		const zipPath = added[0].partName.slice(1)
		const body = new TextDecoder().decode(after.get(zipPath))
		assert(body.includes('hello append'), 'the appended slide carries the authored text')

		// Layout binding resolves to the EXISTING layout — no new chrome.
		const newSlide = reopened.slides[reopened.slides.length - 1]
		assertEqual(
			resolveSingle(reopened.opc, newSlide.partName, SLIDE_LAYOUT_REL),
			target.partName,
			'the appended slide binds to the existing layout part'
		)
		assert(before.has(target.partName.slice(1)), 'the bound layout existed in the original deck')

		// Media resolves; no notes were generated.
		const image = resolveSingle(reopened.opc, newSlide.partName, IMAGE_REL)
		assert(image && reopened.opc.part(image), `the appended slide's image rel resolves (${image})`)
		assertEqual(resolveSingle(reopened.opc, newSlide.partName, NOTES_SLIDE_REL), null, 'no notes slide was generated')
	})

	test('accepts a LayoutHandle and inserts at a chosen position', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('image'))) // LAYOUT_WIDE, 2 slides
		const handle = pres.layouts().find((l) => l.name === 'Leer') // German "Blank"
		assert(handle, 'image deck has a "Leer" layout')

		const pptx = wideGenerator()
		pptx.addSlide().addText('inserted first', { x: 1, y: 1, w: 6, h: 1 })

		const [added] = await pres.appendSlides(pptx, { layout: handle, at: 0 })
		const reopened = await Presentation.load(await pres.save())
		assertEqual(reopened.slides.length, 3, 'the slide was added')
		assertEqual(reopened.slides[0].slideId, added.slideId, 'the appended slide landed first (at: 0)')
	})

	test.skipIf(!validatorInstalled)(
		'the appended deck stays schema-valid (text, image, chart, internal link)',
		async () => {
			const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
			const pptx = wideGenerator()
			const slide = pptx.addSlide()
			slide.addText('valid', { x: 1, y: 1, w: 6, h: 1, color: '0000FF', hyperlink: { slide: 2 } })
			slide.addImage({ data: PNG_1PX, x: 1, y: 3, w: 1, h: 1 })
			slide.addChart(pptx.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
				x: 7,
				y: 1,
				w: 5,
				h: 3,
			})
			pptx.addSlide().addText('target', { x: 1, y: 1, w: 6, h: 1 })
			await pres.appendSlides(pptx, { layout: 'Blank' })

			const errors = await validateBuf(Buffer.from(await pres.save()))
			assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
		}
	)

	test('appends slides carrying embedded audio/video, reproducing the A/V rel graph (dn-append-av-media)', async () => {
		const mp4 = await mediaFixture('tiny.mp4')
		const mp3 = await mediaFixture('tiny.mp3')
		const poster = await mediaFixture('poster.png')

		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)

		const pptx = wideGenerator()
		pptx
			.addSlide()
			.addMedia({
				type: 'video',
				extn: 'mp4',
				data: `video/mp4;base64,${mp4}`,
				cover: `image/png;base64,${poster}`,
				x: 1,
				y: 1,
				w: 4,
				h: 3,
			})
		pptx
			.addSlide()
			.addMedia({
				type: 'audio',
				extn: 'mp3',
				data: `audio/mpeg;base64,${mp3}`,
				cover: `image/png;base64,${poster}`,
				x: 1,
				y: 1,
				w: 2,
				h: 2,
			})

		const added = await pres.appendSlides(pptx, { layout: 'Blank' })
		assertEqual(added.length, 2, 'two A/V slides were appended')

		const out = await pres.save()
		const after = await partBodies(out)

		// Byte-stability: only presentation.xml, its .rels, and [Content_Types].xml change among original parts.
		const expectedChanged = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, bytes] of before) {
			assert(after.has(name), `part ${name} survives the append`)
			if (expectedChanged.has(name)) continue
			assert(bytesEqual(bytes, after.get(name)), `part ${name} is byte-identical after the A/V append`)
		}

		// New parts are slides + the media binaries + the preview images only.
		const newParts = [...after.keys()].filter((name) => !before.has(name))
		for (const name of newParts) {
			assert(/^ppt\/(slides|media)\//.test(name), `new part ${name} is a slide or media part`)
		}
		assert(
			newParts.some((n) => /^ppt\/media\/media\d+\.mp4$/.test(n)),
			'an mp4 media part was added'
		)
		assert(
			newParts.some((n) => /^ppt\/media\/media\d+\.mp3$/.test(n)),
			'an mp3 media part was added'
		)
		assert(
			newParts.filter((n) => /^ppt\/media\/image\d+\.png$/.test(n)).length === 2,
			'two preview image parts were added'
		)

		// Content types: A/V parts are registered as Default extension entries (what PowerPoint
		// authors), matching the av-media.pptx oracle — NOT per-part Overrides.
		const ct = new TextDecoder().decode(after.get('[Content_Types].xml'))
		assert(ct.includes('<Default Extension="mp4" ContentType="video/mp4"/>'), 'mp4 Default content type was added')
		assert(
			ct.includes('<Default Extension="mp3" ContentType="audio/mpeg"/>'),
			'mp3 Default content type was added (audio/mpeg, not audio/mp3)'
		)
		assert(
			!/<Override PartName="\/ppt\/media\/[^"]*\.mp[34]"/.test(ct),
			'A/V parts use Defaults, not per-part Overrides'
		)

		// Reopen and assert the per-slide rel graph mirrors the oracle: one ECMA audio/video
		// rel + one MS-2007 media rel sharing the media Target, plus a separate image preview.
		const reopened = await Presentation.load(out)
		const avSlides = reopened.slides.slice(-2)
		const cases = [
			{ slide: avSlides[0], avRel: VIDEO_REL, mediaExt: '.mp4' },
			{ slide: avSlides[1], avRel: AUDIO_REL, mediaExt: '.mp3' },
		]
		for (const { slide, avRel, mediaExt } of cases) {
			const body = new TextDecoder().decode(after.get(slide.partName.slice(1)))

			// Body rId triple (mirrors gen-xml media markup).
			const fileRid = (body.match(/<a:(?:audio|video)File r:link="(rId\d+)"/) || [])[1]
			const embedRid = (body.match(/<p14:media[^>]*r:embed="(rId\d+)"/) || [])[1]
			const blipRid = (body.match(/<a:blip r:embed="(rId\d+)"/) || [])[1]
			assert(fileRid && embedRid && blipRid, `A/V body references the rId triple (${fileRid}/${embedRid}/${blipRid})`)

			// Each body rId resolves to the expected rel type.
			assertEqual(
				typeOfRid(reopened.opc, slide.partName, fileRid),
				avRel,
				'audioFile/videoFile r:link → ECMA audio/video rel'
			)
			assertEqual(
				typeOfRid(reopened.opc, slide.partName, embedRid),
				MS_MEDIA_REL,
				'p14:media r:embed → MS-2007 media rel'
			)
			assertEqual(typeOfRid(reopened.opc, slide.partName, blipRid), IMAGE_REL, 'blip r:embed → image preview rel')

			// The ECMA and MS rels share one media Target; the preview is a distinct image part.
			const mediaTarget = resolveRid(reopened.opc, slide.partName, fileRid)
			assertEqual(
				resolveRid(reopened.opc, slide.partName, embedRid),
				mediaTarget,
				'ECMA + MS rels share the media part Target'
			)
			assert(mediaTarget.endsWith(mediaExt), `media Target is the ${mediaExt} part (${mediaTarget})`)
			assert(reopened.opc.part(mediaTarget), 'the media part exists')
			const previewTarget = resolveRid(reopened.opc, slide.partName, blipRid)
			assert(
				previewTarget !== mediaTarget && reopened.opc.part(previewTarget),
				'the preview image part exists and is distinct'
			)
		}
	})

	test.skipIf(!validatorInstalled)('the appended deck with embedded audio/video stays schema-valid', async () => {
		const mp4 = await mediaFixture('tiny.mp4')
		const mp3 = await mediaFixture('tiny.mp3')
		const poster = await mediaFixture('poster.png')
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx
			.addSlide()
			.addMedia({
				type: 'video',
				extn: 'mp4',
				data: `video/mp4;base64,${mp4}`,
				cover: `image/png;base64,${poster}`,
				x: 1,
				y: 1,
				w: 4,
				h: 3,
			})
		pptx.addSlide().addMedia({ type: 'audio', extn: 'mp3', data: `audio/mpeg;base64,${mp3}`, x: 1, y: 1, w: 2, h: 2 })
		await pres.appendSlides(pptx, { layout: 'Blank' })

		const errors = await validateBuf(Buffer.from(await pres.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test('rejects an unknown layout name', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 })
		assert(
			await rejects(() => pres.appendSlides(pptx, { layout: 'Nonexistent Layout' })),
			'an unknown layout name throws'
		)
	})

	test('rejects a slide-size mismatch', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors'))) // LAYOUT_WIDE
		const pptx = new PptxGenJS() // default LAYOUT_16x9 — narrower
		pptx.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 })
		assert(await rejects(() => pres.appendSlides(pptx, { layout: 'Blank' })), 'a mismatched slide size throws')
	})

	test('appends a slide with a chart, injecting chart + workbook parts and keeping chrome byte-identical', async () => {
		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)

		const pptx = wideGenerator()
		pptx
			.addSlide()
			.addChart(pptx.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], { x: 1, y: 1, w: 6, h: 3 })

		const [added] = await pres.appendSlides(pptx, { layout: 'Blank' })
		const out = await pres.save()
		const after = await partBodies(out)

		// Byte-stability: only the three wiring parts plus brand-new parts change.
		const expectedChanged = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, bytes] of before) {
			assert(after.has(name), `part ${name} survives the append`)
			if (expectedChanged.has(name)) continue
			assert(bytesEqual(bytes, after.get(name)), `part ${name} is byte-identical after the append`)
		}

		// The chart's three parts were injected: chart XML, its .rels, and the workbook.
		const newParts = [...after.keys()].filter((name) => !before.has(name))
		assert(
			newParts.some((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)),
			`a chart part was added (${newParts.join(', ')})`
		)
		assert(
			newParts.some((n) => /^ppt\/charts\/_rels\/chart\d+\.xml\.rels$/.test(n)),
			'the chart .rels was added'
		)
		assert(
			newParts.some((n) => /^ppt\/embeddings\/.*\.xlsx$/.test(n)),
			'the embedded workbook was added'
		)

		// The slide's chart rel resolves, and the chart part's own rel resolves to the workbook.
		const reopened = await Presentation.load(out)
		const newSlide = reopened.slides.find((s) => s.partName === added.partName)
		const chartPart = resolveSingle(reopened.opc, added.partName, CHART_REL)
		assert(chartPart && reopened.opc.part(chartPart), `the slide's chart rel resolves (${chartPart})`)
		assert(newSlide, 'the appended chart slide is present after reopen')
		const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
		const workbook = resolveSingle(reopened.opc, chartPart, PACKAGE_REL)
		assert(workbook && reopened.opc.part(workbook), `the chart's workbook rel resolves (${workbook})`)
	})

	test('carries an internal slide-to-slide hyperlink across, repointed at the appended target', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))

		// Slide 1 links to slide 2 (source numbering); both are appended together.
		const pptx = wideGenerator()
		pptx.addSlide().addText('go to two', { x: 1, y: 1, w: 6, h: 1, hyperlink: { slide: 2 } })
		pptx.addSlide().addText('slide two', { x: 1, y: 1, w: 6, h: 1 })

		const added = await pres.appendSlides(pptx, { layout: 'Blank' })
		assertEqual(added.length, 2, 'both slides were appended')

		const reopened = await Presentation.load(await pres.save())
		const linkTarget = resolveSingle(reopened.opc, added[0].partName, SLIDE_REL)
		assertEqual(linkTarget, added[1].partName, 'the internal link resolves to the 2nd appended slide')
	})

	test('rejects an internal link to a source slide outside the appended batch', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx.addSlide().addText('dangling', { x: 1, y: 1, w: 6, h: 1, hyperlink: { slide: 9 } })
		assert(
			await rejects(() => pres.appendSlides(pptx, { layout: 'Blank' })),
			'a link to a non-appended source slide throws'
		)
	})
})
