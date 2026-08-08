#!/usr/bin/env node
/**
 * Append-path ceiling — what `appendSlides` itself costs, measured rather than quoted.
 *
 * Why this exists. The loss list in the `appendSlides` doc comment is prose, and prose has
 * already been wrong here (two limitations it advertised had been implemented and removed).
 * A converter built on this path can never beat the path's own fidelity, so that ceiling has
 * to be a number before anything is built against it.
 *
 * Method — a "null converter". Rather than reading a fixture and hand-authoring it back
 * (which measures the hand-authoring, not the append), each probe:
 *   1. authors one construct through the public write API,
 *   2. appends it onto a real PowerPoint-authored template via `fromTemplate` + `appendSlides`,
 *   3. saves, re-reads the output through `ts-pptx/read`,
 *   4. compares what comes back to what went in.
 *
 * Anything that fails to survive is a cost of the append path alone: no IR and no printer
 * are involved. `expected: 'loss'` marks a documented limitation — those failing is the
 * plan holding, and one of them *passing* is news.
 *
 * Usage: `node scripts/append-ceiling.mjs --help`.
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, parseCliOrExit } from './script-utils.mjs'
import TsPptx from '../dist/node.js'
import { Presentation } from '../dist/read.js'

const DEFAULT_TEMPLATE = path.join('test', 'read', 'fixtures', 'placeholder-inherit.pptx')

const USAGE = `Append-ceiling probe — what survives appending a slide to a template.

  node scripts/append-ceiling.mjs
  node scripts/append-ceiling.mjs --json
  node scripts/append-ceiling.mjs --template test/read/fixtures/mixed.pptx

Options:
  --template <path>  template deck, relative to the repo root (default ${DEFAULT_TEMPLATE})
  --json             machine-readable report on stdout
  -h, --help         show this message`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: { json: { type: 'boolean', default: false }, template: { type: 'string' } },
})
const asJson = values.json
const TEMPLATE = path.join(ROOT, values.template ?? DEFAULT_TEMPLATE)

const templateBytes = await fs.readFile(TEMPLATE)

/** Author one slide, append it to the template, save, and hand back the re-read slide. */
async function roundTrip(author) {
	const deck = await Presentation.fromTemplate(templateBytes)
	const size = deck.slideSize
	if (!size) throw new Error(`template ${TEMPLATE} declares no slide size`)
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'MATCH', width: size.widthEmu / 914400, height: size.heightEmu / 914400 })
	pptx.layout = 'MATCH'
	const slide = pptx.addSlide()
	author(slide, pptx)
	const layout = deck.layouts()[0]
	if (!layout) throw new Error(`template ${TEMPLATE} exposes no bindable layout`)
	await deck.appendSlides(pptx, { layout })
	const out = await deck.save()
	const reread = await Presentation.load(out)
	// fromTemplate stripped the original slides, so the appended one is the only slide.
	return { slide: reread.slides[0], deck: reread }
}

const firstText = (slide) => slide.shapes.find((s) => s.textFrame?.paragraphs?.length)
const firstRun = (slide) => firstText(slide)?.textFrame.paragraphs[0].runs[0]

/**
 * Each probe: author a construct, then assert what came back.
 * `expected: 'loss'` = a documented Tier B limitation; failing is the plan holding.
 */
const PROBES = [
	{
		construct: 'text content',
		author: (s) => s.addText('Hello ceiling', { x: 1, y: 1, w: 4, h: 1 }),
		check: (slide) => [firstRun(slide)?.text, 'Hello ceiling'],
	},
	{
		construct: 'run bold',
		author: (s) => s.addText([{ text: 'B', options: { bold: true } }], { x: 1, y: 1, w: 4, h: 1 }),
		check: (slide) => [firstRun(slide)?.bold, true],
	},
	{
		construct: 'run italic',
		author: (s) => s.addText([{ text: 'I', options: { italic: true } }], { x: 1, y: 1, w: 4, h: 1 }),
		check: (slide) => [firstRun(slide)?.italic, true],
	},
	{
		construct: 'run font size',
		author: (s) => s.addText('S', { x: 1, y: 1, w: 4, h: 1, fontSize: 33 }),
		check: (slide) => [firstRun(slide)?.fontSizePt, 33],
	},
	{
		construct: 'run font face',
		author: (s) => s.addText('F', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Georgia' }),
		check: (slide) => [firstRun(slide)?.fontName, 'Georgia'],
	},
	{
		construct: 'run hex colour',
		author: (s) => s.addText('C', { x: 1, y: 1, w: 4, h: 1, color: '1F3864' }),
		check: (slide) => [firstRun(slide)?.color, '1F3864'],
	},
	{
		construct: 'run scheme colour (mapped token)',
		author: (s) => s.addText('S', { x: 1, y: 1, w: 4, h: 1, color: 'accent2' }),
		check: (slide) => [firstRun(slide)?.schemeColor, 'accent2'],
	},
	{
		construct: 'paragraph align',
		author: (s) => s.addText('A', { x: 1, y: 1, w: 4, h: 1, align: 'right' }),
		check: (slide) => [firstText(slide)?.textFrame.paragraphs[0].align, 'r'],
	},
	{
		construct: 'bullet glyph',
		author: (s) => s.addText('L', { x: 1, y: 1, w: 4, h: 1, bullet: { characterCode: '2022' } }),
		check: (slide) => [firstText(slide)?.textFrame.paragraphs[0].bullet?.startsWith('char:'), true],
	},
	{
		construct: 'external hyperlink',
		author: (s) =>
			s.addText([{ text: 'link', options: { hyperlink: { url: 'https://example.com/' } } }], {
				x: 1,
				y: 1,
				w: 4,
				h: 1,
			}),
		check: (slide) => [firstRun(slide)?.hyperlink?.url, 'https://example.com/'],
	},
	{
		construct: 'geometry (EMU exact)',
		author: (s) => s.addText('G', { x: '914401emu', y: '523241emu', w: '2743201emu', h: '609601emu' }),
		check: (slide) => {
			const sp = firstText(slide)
			return [[sp?.left, sp?.top, sp?.width, sp?.height].join(','), '914401,523241,2743201,609601']
		},
	},
	{
		construct: 'shape rotation',
		author: (s) => s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, rotate: 45 }),
		check: (slide) => [slide.shapes.find((sp) => sp.rotation)?.rotation, 45],
	},
	{
		construct: 'shape solid fill',
		author: (s) => s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: 'C00000' } }),
		check: (slide) => [slide.shapes.find((sp) => sp.fillColor)?.fillColor, 'C00000'],
	},
	{
		construct: 'shape name',
		author: (s) => s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, objectName: 'MyRect' }),
		check: (slide) => [slide.shapes.some((sp) => sp.name === 'MyRect'), true],
	},
	{
		construct: 'image (embedded png)',
		author: (s) =>
			s.addImage({
				x: 1,
				y: 1,
				w: 2,
				h: 2,
				data:
					'image/png;base64,' +
					'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
			}),
		check: (slide) => [slide.shapes.some((sp) => sp.constructor.name === 'Picture'), true],
	},
	{
		construct: 'table (cell text)',
		author: (s) =>
			s.addTable(
				[
					[{ text: 'r1c1' }, { text: 'r1c2' }],
					[{ text: 'r2c1' }, { text: 'r2c2' }],
				],
				{ x: 1, y: 1, w: 6, colW: [3, 3] }
			),
		check: (slide) => {
			const gf = slide.shapes.find((sp) => sp.table)
			return [gf?.table?.rows?.[1]?.cells?.[0]?.text, 'r2c1']
		},
	},
	{
		construct: 'chart (graphic frame carried)',
		author: (s) =>
			s.addChart([{ name: 'S1', labels: ['a', 'b'], values: [1, 2] }], { type: 'bar', x: 1, y: 1, w: 5, h: 3 }),
		check: (slide) => [slide.shapes.some((sp) => sp.chart), true],
	},
	{
		// The append path repoints `slide:N` at the Nth appended slide's new partname. Two
		// appended slides, the first linking to the second.
		construct: 'internal slide:N hyperlink (repointed)',
		author: (s, pptx) => {
			s.addText([{ text: 'go', options: { hyperlink: { slide: 2 } } }], { x: 1, y: 1, w: 4, h: 1 })
			pptx.addSlide().addText('target', { x: 1, y: 1, w: 4, h: 1 })
		},
		check: (slide, deck) => {
			const link = firstRun(slide)?.hyperlink
			// Resolves to a real slide in the saved deck, not a dangling target.
			const target = link?.slidePartName ?? link?.targetPartName ?? link?.url
			return [Boolean(target && deck.slides.some((sl) => sl.partName === target)), true]
		},
	},
	{
		construct: 'gradient fill',
		author: (s) =>
			s.addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'FF0000' },
							{ position: 100, color: '0000FF' },
						],
					},
				},
			}),
		check: (slide) => {
			const stops = slide.shapes.map((sp) => sp.gradientStops).find((g) => g?.length)
			return [stops?.map((st) => st.color ?? st.hex).join(','), 'FF0000,0000FF']
		},
	},
	{
		construct: 'shape outline (colour + width)',
		author: (s) => s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, line: { color: '00B050', width: 3 } }),
		check: (slide) => [slide.shapes.find((sp) => sp.lineColor)?.lineColor, '00B050'],
	},
	{
		construct: 'speaker notes',
		author: (s) => s.addNotes('These are the notes'),
		check: (slide) => [slide.notesText, 'These are the notes'],
	},
	{
		construct: 'placeholder type (title)',
		expected: 'partial',
		author: (s) => s.addText('T', { placeholder: 'title', x: 1, y: 1, w: 4, h: 1 }),
		check: (slide) => [firstText(slide)?.placeholder?.type, 'title'],
	},
	{
		construct: 'placeholder idx',
		expected: 'loss',
		author: (s) => s.addText('T', { placeholder: 'body', x: 1, y: 1, w: 4, h: 1 }),
		// No public setter for idx, so it can never round-trip a source value.
		check: (slide) => [firstText(slide)?.placeholder?.idx, '7'],
	},
	{
		construct: 'slide background colour',
		author: (s) => {
			s.background = { color: 'EEEEEE' }
			s.addText('bg', { x: 1, y: 1, w: 2, h: 1 })
		},
		// SlideBackground is a discriminated union; the solid variant carries a ResolvedColor.
		check: (slide) => {
			const bg = slide.background
			return [
				bg?.type === 'solid' ? (bg.color?.hex ?? bg.color?.value ?? JSON.stringify(bg.color)) : bg?.type,
				'EEEEEE',
			]
		},
	},
]

const results = []
for (const probe of PROBES) {
	let status
	let detail = ''
	try {
		const { slide, deck } = await roundTrip(probe.author)
		const [actual, expected] = probe.check(slide, deck)
		const ok = String(actual) === String(expected)
		status = ok ? 'survives' : 'lost'
		if (!ok) detail = `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
	} catch (err) {
		status = 'threw'
		detail = err instanceof Error ? err.message : String(err)
	}
	results.push({ construct: probe.construct, status, expected: probe.expected ?? 'survives', detail })
}

if (asJson) {
	console.log(JSON.stringify({ template: path.relative(ROOT, TEMPLATE), results }, null, 2))
} else {
	const survives = results.filter((r) => r.status === 'survives')
	const documented = results.filter((r) => r.status !== 'survives' && r.expected !== 'survives')
	const surprises = results.filter((r) => r.status !== 'survives' && r.expected === 'survives')
	const unexpectedPass = results.filter((r) => r.status === 'survives' && r.expected === 'loss')

	console.log(`append-path ceiling — template: ${path.relative(ROOT, TEMPLATE)}`)
	console.log(
		`  ${survives.length}/${results.length} constructs survive the append round-trip ` +
			`(${documented.length} documented losses, ${surprises.length} undocumented)`
	)
	const line = (r) => {
		const mark = r.status === 'survives' ? 'ok  ' : r.status === 'threw' ? 'THREW' : 'LOST'
		console.log(`  ${mark.padEnd(6)} ${r.construct.padEnd(36)} ${r.detail}`)
	}
	console.log('\nSurvives:')
	survives.forEach(line)
	if (documented.length) {
		console.log('\nDocumented losses (failing here means the plan holds):')
		documented.forEach(line)
	}
	if (unexpectedPass.length) {
		console.log('\n!! Documented as a loss but SURVIVED — the doc comment is stale again:')
		unexpectedPass.forEach(line)
	}
	if (surprises.length) {
		console.log('\n!! UNDOCUMENTED losses — these are the ones that reprice the plan:')
		surprises.forEach(line)
	}
}
