// Append-onto-existing tests for `ts-pptx/read` (dn-append-onto-existing-deck).
//
// Contract under test: Presentation.appendSlides(source, { layout }) authors
// slides on a generator (TsPptx), serializes them via source.extractSlides(),
// and splices them into a loaded deck bound to an existing layout — keeping the
// deck's masters/layouts/theme (and every other untouched part) byte-identical,
// changing only presentation.xml, its .rels, [Content_Types].xml, and the new
// slide/media parts. Survives a save → reopen round-trip, resolves its layout to
// the *existing* layout (no new chrome), and stays schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import TsPptx, { ChartType } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { bytesEqual, PNG_1X1, assert, assertEqual, assertIncludes, partBodies } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { FIXTURES, fixturePath } from './corpus.js'
import { resolveSingle } from './opc.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await validatorAvailable()

const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx'
const CHART_STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
const CHART_COLOR_STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
const CHARTEX_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml'
const CHART_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartstyle+xml'
const CHART_COLOR_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

async function rejects(fn) {
	try {
		await fn()
		return false
	} catch {
		return true
	}
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
	const buf = await readFile(path.join(FIXTURES, 'media', name))
	return buf.toString('base64')
}

/** A generator deck sized to LAYOUT_WIDE (12192000×6858000), matching theme-colors / image. */
function wideGenerator() {
	const pptx = new TsPptx()
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
		slide.addImage({ data: PNG_1X1, x: 1, y: 3, w: 1, h: 1 })

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
			slide.addImage({ data: PNG_1X1, x: 1, y: 3, w: 1, h: 1 })
			slide.addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
				type: ChartType.bar,
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
		pptx.addSlide().addMedia({
			type: 'video',
			extn: 'mp4',
			data: `video/mp4;base64,${mp4}`,
			cover: `image/png;base64,${poster}`,
			x: 1,
			y: 1,
			w: 4,
			h: 3,
		})
		pptx.addSlide().addMedia({
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
		pptx.addSlide().addMedia({
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

	test('carries an online (external-link) video: two External rels + poster, no media part', async () => {
		// Mirrors the online-video.pptx oracle: an appended online video reproduces the
		// external-link rel graph PowerPoint authors — an ECMA video rel AND an MS-2007
		// media rel, both TargetMode="External" sharing one link Target, plus a poster
		// image part — with NO media binary part and NO A/V content-type entry.
		const link = 'https://example.com/online-video-sample.mp4'
		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)

		const pptx = wideGenerator()
		pptx.addSlide().addMedia({ type: 'online', link, cover: PNG_1X1, x: 1, y: 1, w: 4, h: 3 })

		const [added] = await pres.appendSlides(pptx, { layout: 'Blank' })
		const out = await pres.save()
		const after = await partBodies(out)

		// Only chrome changes among original parts.
		const expectedChanged = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, bytes] of before) {
			assert(after.has(name), `part ${name} survives the append`)
			if (expectedChanged.has(name)) continue
			assert(bytesEqual(bytes, after.get(name)), `part ${name} is byte-identical after the online-video append`)
		}

		// The only new media part is the poster image — no ppt/media/*.mp4 binary.
		const newParts = [...after.keys()].filter((name) => !before.has(name))
		assert(
			newParts.filter((n) => n.startsWith('ppt/media/')).every((n) => /^ppt\/media\/image\d+\.png$/.test(n)),
			`the only new media part is the poster image (${newParts.join(', ')})`
		)
		assert(!newParts.some((n) => /^ppt\/media\/.*\.mp4$/.test(n)), 'no mp4 media binary was added')

		// No A/V content-type entry — the video is a link, not a part.
		const ct = new TextDecoder().decode(after.get('[Content_Types].xml'))
		assert(!/Extension="mp4"/.test(ct), 'no mp4 content type was added for the linked video')

		// Body rId triple: a:videoFile r:link (ECMA), p14:media r:link (MS-2007), a:blip r:embed (poster).
		const reopened = await Presentation.load(out)
		const body = new TextDecoder().decode(after.get(added.partName.slice(1)))
		const fileRid = (body.match(/<a:videoFile r:link="(rId\d+)"/) || [])[1]
		const mediaRid = (body.match(/<p14:media[^>]*r:link="(rId\d+)"/) || [])[1]
		const blipRid = (body.match(/<a:blip r:embed="(rId\d+)"/) || [])[1]
		assert(
			fileRid && mediaRid && blipRid,
			`online-video body references the rId triple (${fileRid}/${mediaRid}/${blipRid})`
		)

		// Each body rId resolves to the expected rel type.
		assertEqual(typeOfRid(reopened.opc, added.partName, fileRid), VIDEO_REL, 'videoFile r:link → ECMA video rel')
		assertEqual(typeOfRid(reopened.opc, added.partName, mediaRid), MS_MEDIA_REL, 'p14:media r:link → MS-2007 media rel')
		assertEqual(typeOfRid(reopened.opc, added.partName, blipRid), IMAGE_REL, 'blip r:embed → poster image rel')

		// The two external rels share the link Target and are External; the poster is a real part.
		const rels = [...reopened.opc.relationshipsFor(added.partName)]
		const videoRel = rels.find((r) => r.id === fileRid)
		const msRel = rels.find((r) => r.id === mediaRid)
		assertEqual(videoRel.target, link, 'ECMA video rel Target is the external link')
		assertEqual(msRel.target, link, 'MS-2007 media rel shares the same link Target')
		assertEqual(videoRel.targetMode, 'External', 'ECMA video rel is External')
		assertEqual(msRel.targetMode, 'External', 'MS-2007 media rel is External')
		const posterTarget = resolveRid(reopened.opc, added.partName, blipRid)
		assert(reopened.opc.part(posterTarget), 'the poster image part exists')
	})

	test.skipIf(!validatorInstalled)('the appended deck with an online video stays schema-valid', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx
			.addSlide()
			.addMedia({ type: 'online', link: 'https://example.com/v.mp4', cover: PNG_1X1, x: 1, y: 1, w: 4, h: 3 })
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
		const pptx = new TsPptx() // default LAYOUT_16x9 — narrower
		pptx.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 })
		assert(await rejects(() => pres.appendSlides(pptx, { layout: 'Blank' })), 'a mismatched slide size throws')
	})

	test('appends a chartEx chart as a chartEx part, with its mandatory style/colors sidecars', async () => {
		// chartEx (Office 2016: waterfall, funnel, treemap, ...) is a different part in a different
		// namespace, reached through the MS chartEx rel, and PowerPoint reports it as corrupt without
		// its two style sidecars. It used to be serialized by the classic builder, which has no arm
		// for those types: the deck got a `<c:chartSpace>` with axes and no plot, behind a slide
		// still pointing at it through `<mc:AlternateContent><cx:chart>`. Every assertion below is
		// one coordinate of that: the part name, the content type, the rel type, the sidecars, and
		// a plot element that is actually there.
		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)

		const pptx = wideGenerator()
		pptx.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
			type: ChartType.waterfall,
			x: 1,
			y: 1,
			w: 6,
			h: 3,
		})

		const [added] = await pres.appendSlides(pptx, { layout: 'Blank' })
		const out = await pres.save()
		const after = await partBodies(out)

		// Byte-stability: only the three wiring parts plus brand-new parts change.
		const expectedChanged = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, bytes] of before) {
			assert(after.has(name), `part ${name} survives the append`)
			if (expectedChanged.has(name)) continue
			assert(bytesEqual(bytes, after.get(name)), `part ${name} is byte-identical after the chartEx append`)
		}

		// The chartEx part is its own name family, and both sidecars rode along.
		const newParts = [...after.keys()].filter((name) => !before.has(name))
		const chartZipPath = newParts.find((n) => /^ppt\/charts\/chartEx\d+\.xml$/.test(n))
		assert(chartZipPath, `a chartEx part was added (${newParts.join(', ')})`)
		assert(!newParts.some((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)), 'no classic chart part was written')
		assert(
			newParts.some((n) => /^ppt\/charts\/style\d+\.xml$/.test(n)),
			'the chart-style sidecar was added'
		)
		assert(
			newParts.some((n) => /^ppt\/charts\/colors\d+\.xml$/.test(n)),
			'the color-style sidecar was added'
		)

		// The plot is really there — the failure this test exists for was a chart part with none.
		const chartXml = new TextDecoder().decode(after.get(chartZipPath))
		assertIncludes(chartXml, '<cx:chartSpace', 'the chart part is a cx:chartSpace')
		assertIncludes(chartXml, 'layoutId="waterfall"', 'the waterfall plot element is present')

		// Three Overrides: PowerPoint resolves all three parts by their MS content types.
		const ct = new TextDecoder().decode(after.get('[Content_Types].xml'))
		for (const type of [CHARTEX_CONTENT_TYPE, CHART_STYLE_CONTENT_TYPE, CHART_COLOR_STYLE_CONTENT_TYPE]) {
			assertIncludes(ct, type, `[Content_Types].xml declares ${type}`)
		}

		// The slide reaches the chart through the MS chartEx rel, not the ECMA `chart` one, and the
		// body's `<cx:chart r:id>` is that rel.
		const reopened = await Presentation.load(out)
		const chartPart = resolveSingle(reopened.opc, added.partName, CHARTEX_REL)
		assertEqual(`/${chartZipPath}`, chartPart, "the slide's chartEx rel resolves to the chartEx part")
		assertEqual(resolveSingle(reopened.opc, added.partName, CHART_REL), null, 'no classic chart rel was written')
		const body = new TextDecoder().decode(after.get(added.partName.slice(1)))
		const chartRid = (body.match(/<cx:chart[^>]*r:id="(rId\d+)"/) || [])[1]
		assertEqual(typeOfRid(reopened.opc, added.partName, chartRid), CHARTEX_REL, 'cx:chart r:id → MS chartEx rel')

		// The chart part's own three rels resolve: workbook, colors, style.
		const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
		for (const [rel, label] of [
			[PACKAGE_REL, 'workbook'],
			[CHART_COLOR_STYLE_REL, 'color-style'],
			[CHART_STYLE_REL, 'chart-style'],
		]) {
			const target = resolveSingle(reopened.opc, chartPart, rel)
			assert(target && reopened.opc.part(target), `the chart's ${label} rel resolves (${target})`)
		}
	})

	test.skipIf(!validatorInstalled)('the appended deck with a chartEx chart stays schema-valid', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
			type: ChartType.waterfall,
			x: 1,
			y: 1,
			w: 6,
			h: 3,
		})
		await pres.appendSlides(pptx, { layout: 'Blank' })

		const errors = await validateBuf(Buffer.from(await pres.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test('appends a slide with a chart, injecting chart + workbook parts and keeping chrome byte-identical', async () => {
		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)

		const pptx = wideGenerator()
		pptx.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
			type: ChartType.bar,
			x: 1,
			y: 1,
			w: 6,
			h: 3,
		})

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

	test('round-trips & in hyperlink and online-video Targets without double-escaping', async () => {
		// `SlideRel.Target` is stored unescaped and escaped by whichever emitter writes it —
		// here `read/opc/relationships.ts`. Escaping at definition time instead would land a
		// literal `&amp;` in the URL on this path (the serializer escaping it a second time),
		// which is well-formed XML and so invisible to the validator. Reading the targets back
		// through the parser is what catches it: they must equal the URLs that went in.
		const LINK = 'https://example.com/?a=1&b=2'
		const VIDEO = 'https://example.com/embed/ID?rel=0&t=5'

		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		const slide = pptx.addSlide()
		slide.addText([{ text: 'link', options: { hyperlink: { url: LINK } } }], { x: 1, y: 1, w: 6, h: 1 })
		slide.addMedia({ type: 'online', link: VIDEO, cover: PNG_1X1, x: 1, y: 3, w: 4, h: 3 })

		const [added] = await pres.appendSlides(pptx, { layout: 'Blank' })
		const reopened = await Presentation.load(await pres.save())
		const rels = [...reopened.opc.relationshipsFor(added.partName)]

		const hyperlink = rels.find((r) => r.type === HYPERLINK_REL)
		assertEqual(hyperlink.target, LINK, 'hyperlink Target survives the round-trip unmangled')
		// Both halves of the online-video pair share the link and must agree.
		assertEqual(rels.find((r) => r.type === VIDEO_REL).target, VIDEO, 'ECMA video rel Target is unmangled')
		assertEqual(rels.find((r) => r.type === MS_MEDIA_REL).target, VIDEO, 'MS-2007 media rel Target is unmangled')
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
